//! Port of test/session.test.mjs: cookie import, silent sign-in with PKCE, refresh
//! token with Origin, fallback, day-old refresh token skipped, rotation, status without
//! secrets, needs_sign_in, cookies.txt, logout.

mod common;

use base64::Engine;
use common::*;
use outlook_web_cli::commands;
use outlook_web_cli::session::{self, *};
use regex::Regex;
use serde_json::{json, Value};
use sha2::{Digest, Sha256};
use std::collections::HashMap;
use std::sync::atomic::{AtomicBool, AtomicU32, Ordering};
use std::sync::{Arc, Mutex};

const TENANT: &str = "11111111-2222-3333-4444-555555555555";

/// A fake login.microsoftonline.com: the silent authorize endpoint and the token endpoint.
fn fake_login(login_required: Arc<AtomicBool>) -> Fake {
    let codes: Arc<Mutex<HashMap<String, String>>> = Arc::default();
    let counter = Arc::new(AtomicU32::new(0));
    let cookie_re = Regex::new(r"ESTSAUTHPERSISTENT=cookie-v(\d+)").unwrap();
    Fake::start(move |req| {
        if req.path.ends_with("/oauth2/v2.0/authorize") {
            let redirect = req.q("redirect_uri").unwrap_or_default();
            let cookie = req.header("cookie").unwrap_or_default();
            let caps = cookie_re.captures(&cookie);
            if login_required.load(Ordering::SeqCst) || caps.is_none() {
                return Reply::json(302, json!({})).header(
                    "location",
                    &format!("{redirect}#error=login_required&error_description=AADSTS50058%3A+no+session"),
                );
            }
            let mut codes = codes.lock().unwrap();
            let code = format!("code-{}", codes.len() + 1);
            codes.insert(code.clone(), req.q("code_challenge").unwrap_or_default());
            let n: u32 = caps.unwrap()[1].parse().unwrap();
            return Reply::json(302, json!({}))
                .header("location", &format!("{redirect}#code={code}&state={}", req.q("state").unwrap_or_default()))
                .header("set-cookie", &format!("ESTSAUTHPERSISTENT=cookie-v{}; domain=.login.microsoftonline.com; path=/; expires=Wed, 10 Feb 2027 00:00:00 GMT; secure; HttpOnly", n + 1))
                .header("set-cookie", "buid=; domain=login.microsoftonline.com; path=/; max-age=0");
        }
        if req.path.ends_with("/oauth2/v2.0/token") {
            if req.header("origin").as_deref() != Some("https://outlook.cloud.microsoft") {
                return Reply::json(
                    400,
                    json!({ "error": "invalid_request", "error_description": "AADSTS9002327: Tokens issued for the 'Single-Page Application' client-type may only be redeemed via cross-origin requests." }),
                );
            }
            match req.form("grant_type").as_deref() {
                Some("authorization_code") => {
                    let challenge = codes
                        .lock()
                        .unwrap()
                        .get(&req.form("code").unwrap_or_default())
                        .cloned();
                    let expected = base64::engine::general_purpose::URL_SAFE_NO_PAD.encode(
                        Sha256::digest(req.form("code_verifier").unwrap_or_default().as_bytes()),
                    );
                    if challenge.as_deref() != Some(expected.as_str()) {
                        return Reply::json(
                            400,
                            json!({ "error": "invalid_grant", "error_description": "PKCE mismatch" }),
                        );
                    }
                }
                Some("refresh_token")
                    if !req
                        .form("refresh_token")
                        .unwrap_or_default()
                        .starts_with("rt-") =>
                {
                    return Reply::json(
                        400,
                        json!({ "error": "invalid_grant", "error_description": "AADSTS700084: The refresh token was issued to a single page app (SPA), and therefore has a fixed, limited lifetime of 1.00:00:00" }),
                    );
                }
                _ => {}
            }
            let n = counter.fetch_add(1, Ordering::SeqCst) + 1;
            let exp = chrono::Utc::now().timestamp() + 3600;
            return Reply::json(
                200,
                json!({
                    "token_type": "Bearer",
                    "expires_in": 3600,
                    "access_token": jwt(json!({ "aud": "https://outlook.office.com", "tid": TENANT, "scp": "Mail.ReadWrite Calendars.ReadWrite", "exp": exp, "n": n })),
                    "refresh_token": format!("rt-{n}"),
                    "id_token": jwt(json!({ "tid": TENANT, "oid": "oid-1", "preferred_username": "me@contoso.com" })),
                }),
            );
        }
        Reply::json(404, json!({}))
    })
}

fn cookie_value(s: &serde_json::Map<String, Value>, name: &str) -> Option<String> {
    s["cookies"]
        .as_array()?
        .iter()
        .find(|c| c["name"] == json!(name))
        .map(|c| c["value"].as_str().unwrap_or("").to_string())
}

#[test]
fn session_flow() {
    let login_required = Arc::new(AtomicBool::new(false));
    let server = fake_login(login_required.clone());
    let dir = config_dir(json!([{ "name": "contoso", "username": "me@contoso.com" }]));
    let ctx = ctx(&dir, &[("OUTLOOK_WEB_LOGIN_BASE_URL", &server.base)]);

    // 1. Import cookies exported from a browser (Cookie-Editor JSON); only Microsoft sign-in cookies are kept.
    let exported = json!([
        { "name": "ESTSAUTHPERSISTENT", "value": "cookie-v1", "domain": ".login.microsoftonline.com", "path": "/", "expirationDate": 1802000000 },
        { "name": "buid", "value": "b", "domain": "login.microsoftonline.com", "path": "/" },
        { "name": "SID", "value": "google", "domain": ".google.com", "path": "/" },
    ])
    .to_string();
    let imported = import_cookies(&ctx, "contoso", &exported, "me@contoso.com", "").unwrap();
    assert_eq!(imported["ok"], json!(true));
    assert_eq!(imported["imported_cookies"], json!(2));
    let file = dir.path().join("sessions").join("contoso.json");
    #[cfg(unix)]
    {
        use std::os::unix::fs::PermissionsExt;
        assert_eq!(
            std::fs::metadata(&file).unwrap().permissions().mode() & 0o777,
            0o600,
            "session file is private"
        );
    }

    // 2. Cookies -> silent sign-in (prompt=none + PKCE) -> token; cookies rotate, tenant and RT are saved.
    let first = refresh_session(&ctx, "contoso", None, false).unwrap();
    assert!(first.ok, "{:?}", first);
    assert_eq!(first.via, "cookies");
    let authorize = server
        .requests()
        .into_iter()
        .find(|r| r.path.ends_with("/authorize"))
        .unwrap();
    assert_eq!(authorize.q("prompt").as_deref(), Some("none"));
    assert_eq!(
        authorize.q("code_challenge_method").as_deref(),
        Some("S256")
    );
    assert_eq!(authorize.q("login_hint").as_deref(), Some("me@contoso.com"));
    assert!(
        authorize.path.starts_with("/organizations/"),
        "first sign-in uses the organizations endpoint"
    );
    let saved = load_session(&ctx, "contoso").unwrap();
    assert_eq!(saved["tenant_id"], json!(TENANT));
    assert_eq!(saved["refresh_token"], json!("rt-1"));
    assert_eq!(
        cookie_value(&saved, "ESTSAUTHPERSISTENT").as_deref(),
        Some("cookie-v2"),
        "rotated cookie saved"
    );
    assert_eq!(cookie_value(&saved, "buid"), None, "expired cookie dropped");

    // 3. A valid saved access token is reused without calling Microsoft.
    server.clear();
    let reused = get_session_token(&ctx, "contoso", true).unwrap().unwrap();
    assert_eq!(reused.source, "saved");
    assert_eq!(reused.api, "outlook_rest");
    assert_eq!(server.count(), 0);

    // 4. An expired access token is renewed with the refresh token (Origin header required).
    let mut s = load_session(&ctx, "contoso").unwrap();
    s["access_token"]["expires_on"] = json!(1);
    save_session(&ctx, "contoso", &s).unwrap();
    let renewed = get_session_token(&ctx, "contoso", true).unwrap().unwrap();
    assert_eq!(renewed.source, "refresh_token");
    let last = server.last();
    assert_eq!(last.form("grant_type").as_deref(), Some("refresh_token"));
    assert_eq!(last.form("refresh_token").as_deref(), Some("rt-1"));
    assert!(last.path.starts_with(&format!("/{TENANT}/")));
    assert_eq!(
        load_session(&ctx, "contoso").unwrap()["refresh_token"],
        json!("rt-2")
    );

    // 5. A rejected refresh token falls back to the cookies.
    let mut s = load_session(&ctx, "contoso").unwrap();
    s.insert("refresh_token".into(), json!("expired"));
    s.insert("access_token".into(), Value::Null);
    save_session(&ctx, "contoso", &s).unwrap();
    let fallback = refresh_session(&ctx, "contoso", None, false).unwrap();
    assert!(fallback.ok);
    assert_eq!(fallback.via, "cookies");

    // 6. A day-old refresh token is not tried at all.
    let mut s = load_session(&ctx, "contoso").unwrap();
    let old = (chrono::Utc::now() - chrono::Duration::hours(25))
        .to_rfc3339_opts(chrono::SecondsFormat::Millis, true);
    s.insert("refresh_token_obtained_at".into(), json!(old));
    save_session(&ctx, "contoso", &s).unwrap();
    server.clear();
    assert_eq!(
        refresh_session(&ctx, "contoso", None, false).unwrap().via,
        "cookies"
    );
    assert!(!server
        .requests()
        .iter()
        .any(|r| r.form("grant_type").as_deref() == Some("refresh_token")));

    // 7. --rotate (force cookies) always renews the cookies.
    let before = cookie_value(
        &load_session(&ctx, "contoso").unwrap(),
        "ESTSAUTHPERSISTENT",
    );
    assert_eq!(
        refresh_session(&ctx, "contoso", None, true).unwrap().via,
        "cookies"
    );
    assert_ne!(
        cookie_value(
            &load_session(&ctx, "contoso").unwrap(),
            "ESTSAUTHPERSISTENT"
        ),
        before
    );

    // 8. Status shows lifetimes, never secrets; the outlook layer reports where the token came from.
    let secrets = Regex::new(r"rt-|cookie-v|eyJ").unwrap();
    let status = session_status(&ctx, "contoso");
    assert_eq!(status["saved"], json!(true));
    assert_eq!(status["sign_in_cookie"], json!(true));
    assert_eq!(status["sign_in_cookie_expires"], json!("2027-02-10"));
    assert!(status["access_token_valid_min"].as_i64().unwrap() > 50);
    assert!(
        !secrets.is_match(&status.to_string()),
        "no secrets in status"
    );
    let live = commands::run(
        &ctx,
        "status",
        &body(json!({ "account": "contoso" })),
        false,
    )
    .unwrap();
    assert_eq!(live["accounts"][0]["status"], json!("signed_in"));
    assert_eq!(live["accounts"][0]["token_source"], json!("saved"));

    // 8b. refresh_accounts renews every configured account and never returns secrets.
    let all = refresh_accounts(&ctx, None, true).unwrap();
    assert_eq!(all["ok"], json!(true));
    assert_eq!(all["accounts"][0]["account"], json!("contoso"));
    assert_eq!(all["accounts"][0]["via"], json!("cookies"));
    assert!(!secrets.is_match(&all.to_string()));

    // 9. A signed-out session reports needs_sign_in with the reason.
    login_required.store(true, Ordering::SeqCst);
    let mut s = load_session(&ctx, "contoso").unwrap();
    s.insert("refresh_token".into(), json!(""));
    s.insert("access_token".into(), Value::Null);
    save_session(&ctx, "contoso", &s).unwrap();
    let dead = refresh_session(&ctx, "contoso", None, false).unwrap();
    assert!(!dead.ok);
    assert_eq!(dead.status, "needs_sign_in");
    assert!(dead.errors.join(" ").contains("login_required"));
    assert_eq!(get_session_token(&ctx, "contoso", true).unwrap(), None);

    // 10. Netscape cookies.txt exports work too; logout removes the session.
    let txt = "# Netscape HTTP Cookie File\n#HttpOnly_.login.microsoftonline.com\tTRUE\t/\tTRUE\t1802000000\tESTSAUTHPERSISTENT\tcookie-v9\n";
    let parsed = parse_cookie_export(txt).unwrap();
    assert_eq!(
        parsed
            .iter()
            .map(|c| (c["name"].clone(), c["value"].clone(), c["domain"].clone()))
            .collect::<Vec<_>>(),
        vec![(
            json!("ESTSAUTHPERSISTENT"),
            json!("cookie-v9"),
            json!(".login.microsoftonline.com")
        )]
    );
    assert_eq!(
        import_cookies(&ctx, "contoso", "[]", "", "").unwrap()["status"],
        json!("no_sign_in_cookie")
    );
    assert!(delete_session(&ctx, "contoso"));
    assert_eq!(session_status(&ctx, "contoso")["saved"], json!(false));
}

/// Session files written by the Node.js version (v0.3.0) load and are rewritten with
/// the same field names and order, keeping fields this version doesn't know.
#[test]
fn node_session_files_stay_compatible() {
    let dir = config_dir(json!([]));
    let ctx = ctx(&dir, &[]);
    let node_file = r#"{
  "source": "chrome:9230",
  "client_id": "9199bf20-a13f-4107-85dc-02114787ef48",
  "origin": "https://outlook.cloud.microsoft",
  "redirect_uri": "https://outlook.cloud.microsoft/mail/",
  "tenant_id": "t-1",
  "home_account_id": "oid.t-1",
  "username": "me@contoso.com",
  "refresh_token": "rt-x",
  "refresh_token_obtained_at": "2026-11-16T10:00:00.000Z",
  "cookies": [
    {
      "name": "ESTSAUTHPERSISTENT",
      "value": "v",
      "domain": ".login.microsoftonline.com",
      "path": "/",
      "expires": 1802000000
    }
  ],
  "access_token": {
    "token": "a.b.c",
    "expires_on": 4102444800,
    "audience": "https://outlook.office.com"
  },
  "future_field": {
    "kept": true
  },
  "version": 1,
  "account": "work",
  "updated_at": "2026-11-16T10:00:00.000Z"
}"#;
    std::fs::create_dir_all(dir.path().join("sessions")).unwrap();
    std::fs::write(dir.path().join("sessions/work.json"), node_file).unwrap();
    let tok = get_session_token(&ctx, "work", false).unwrap().unwrap();
    assert_eq!(tok.token, "a.b.c");
    assert_eq!(tok.source, "saved");
    let s = load_session(&ctx, "work").unwrap();
    save_session(&ctx, "work", &s).unwrap();
    let rewritten = std::fs::read_to_string(dir.path().join("sessions/work.json")).unwrap();
    let strip = |t: &str| {
        Regex::new(r#""updated_at": "[^"]+""#)
            .unwrap()
            .replace(t, "")
            .to_string()
    };
    assert_eq!(
        strip(&rewritten),
        strip(node_file),
        "byte-identical apart from updated_at"
    );
    // The account shows up from its saved session alone.
    assert_eq!(
        outlook_web_cli::accounts::list_accounts(&ctx)[0].name,
        "work"
    );
    let _ = session::OUTLOOK_WEB_CLIENT_ID;
}

#[test]
fn set_cookie_merge() {
    let existing = vec![
        json!({ "name": "A", "value": "1", "domain": ".login.microsoftonline.com", "path": "/", "expires": null }),
        json!({ "name": "B", "value": "2", "domain": "login.microsoftonline.com", "path": "/", "expires": null }),
    ];
    let merged = merge_set_cookies(
        &existing,
        &[
            "A=new; Domain=login.microsoftonline.com; Path=/".into(),
            "B=; max-age=0".into(),
            "C=3; expires=Wed, 10-Feb-2027 00:00:00 GMT".into(),
            "D=4; expires=Wed, 10 Feb 2010 00:00:00 GMT".into(),
        ],
    );
    let names: Vec<_> = merged
        .iter()
        .map(|c| (c["name"].as_str().unwrap(), c["value"].as_str().unwrap()))
        .collect();
    assert_eq!(names, vec![("A", "new"), ("C", "3")]);
    assert_eq!(merged[1]["expires"], json!(1802217600));
}
