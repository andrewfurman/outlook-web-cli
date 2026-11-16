//! Saved Outlook web sessions, used like a browser would: no Chrome needed after import.
//!
//! A session file (`sessions/<account>.json`, mode 0600) holds what Outlook web itself keeps:
//! - the MSAL refresh token of the Outlook web app (a single-page app, so the token is
//!   redeemed with Outlook's Origin header and lasts 24 hours from sign-in),
//! - the login.microsoftonline.com sign-in cookies (ESTSAUTHPERSISTENT and friends, which
//!   last months when "Stay signed in" was chosen),
//! - the current access token for <https://outlook.office.com> and its expiry.
//!
//! The file format is the same one the Node.js version (v0.3.0) wrote; files are kept as
//! JSON objects, so fields this version does not know about survive a rewrite.
//!
//! Getting a token tries, in order: the saved access token, the refresh token, then a
//! silent sign-in with the cookies (prompt=none + PKCE, as Outlook web's own hidden
//! iframe does), which also yields a fresh refresh token and rotated cookies. When the
//! account has a Chrome DevTools port, a dead session is re-imported from that Chrome.
//! Token and cookie values are never logged or returned by status calls.

use crate::accounts::{get_account, list_accounts, NAME_RE};
use crate::cdp::{connect_cdp, find_outlook_tab};
use crate::ctx::{now_iso, now_ms, now_sec_f, Ctx};
use crate::js::{self, num, set_opt, Body};
use anyhow::{bail, Result};
use base64::Engine;
use regex::Regex;
use serde_json::{json, Map, Value};
use sha2::{Digest, Sha256};
use std::path::PathBuf;
use std::sync::LazyLock;
use std::time::Duration;

pub const OUTLOOK_WEB_CLIENT_ID: &str = "9199bf20-a13f-4107-85dc-02114787ef48";
pub const OUTLOOK_WEB_ORIGIN: &str = "https://outlook.cloud.microsoft";
pub const OUTLOOK_WEB_REDIRECT_URI: &str = "https://outlook.cloud.microsoft/mail/";
pub const OUTLOOK_SCOPE: &str = "https://outlook.office.com/.default openid profile offline_access";

const ACCESS_TOKEN_MARGIN_SEC: f64 = 120.0;
/// Single-page-app refresh tokens expire 24 hours after the sign-in that issued them.
const REFRESH_TOKEN_MAX_AGE_MS: f64 = 23.5 * 60.0 * 60.0 * 1000.0;
const LOGIN_TIMEOUT: Duration = Duration::from_secs(15);
const BROWSER_UA: &str =
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/141.0.0.0 Safari/537.36";

static LOGIN_DOMAIN_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"(^|\.)login\.microsoftonline\.com$").unwrap());

/// A token for the Outlook REST API and where it came from.
#[derive(Clone, Debug, PartialEq)]
pub struct SessionToken {
    pub token: String,
    pub api: String,
    pub source: String,
}

fn login_base(ctx: &Ctx) -> String {
    ctx.var("OUTLOOK_WEB_LOGIN_BASE_URL")
        .unwrap_or("https://login.microsoftonline.com")
        .trim_end_matches('/')
        .to_string()
}

fn session_path(ctx: &Ctx, name: &str) -> Result<PathBuf> {
    if !NAME_RE.is_match(name) {
        bail!("Invalid account name '{name}'.");
    }
    Ok(ctx.sessions_dir().join(format!("{name}.json")))
}

pub fn load_session(ctx: &Ctx, name: &str) -> Option<Body> {
    let text = std::fs::read_to_string(session_path(ctx, name).ok()?).ok()?;
    match serde_json::from_str::<Value>(&text).ok()? {
        Value::Object(m) => Some(m),
        _ => None,
    }
}

/// Writes the session atomically with mode 0600, adding version, account and updated_at.
pub fn save_session(ctx: &Ctx, name: &str, session: &Body) -> Result<()> {
    let file = session_path(ctx, name)?;
    let dir = file.parent().expect("sessions dir");
    create_private_dir(dir)?;
    let mut out = session.clone();
    out.insert("version".into(), json!(1));
    out.insert("account".into(), json!(name));
    out.insert("updated_at".into(), json!(now_iso()));
    let tmp = PathBuf::from(format!(
        "{}.{}.{}.tmp",
        file.display(),
        std::process::id(),
        now_ms()
    ));
    write_private(
        &tmp,
        serde_json::to_string_pretty(&Value::Object(out))?.as_bytes(),
    )?;
    std::fs::rename(&tmp, &file)?;
    Ok(())
}

pub fn create_private_dir(dir: &std::path::Path) -> Result<()> {
    let mut b = std::fs::DirBuilder::new();
    b.recursive(true);
    #[cfg(unix)]
    std::os::unix::fs::DirBuilderExt::mode(&mut b, 0o700);
    b.create(dir)?;
    Ok(())
}

/// Writes a file that only the owner can read (mode 0600 when it is created).
pub fn write_private(path: &std::path::Path, bytes: &[u8]) -> Result<()> {
    use std::io::Write;
    let mut o = std::fs::OpenOptions::new();
    o.write(true).create(true).truncate(true);
    #[cfg(unix)]
    std::os::unix::fs::OpenOptionsExt::mode(&mut o, 0o600);
    let mut f = o.open(path)?;
    f.write_all(bytes)?;
    Ok(())
}

pub fn delete_session(ctx: &Ctx, name: &str) -> bool {
    session_path(ctx, name)
        .map(|p| std::fs::remove_file(p).is_ok())
        .unwrap_or(false)
}

fn parse_iso_ms(v: Option<&Value>) -> f64 {
    match v.and_then(Value::as_str) {
        Some(s) => chrono::DateTime::parse_from_rfc3339(s)
            .map(|d| d.timestamp_millis() as f64)
            .unwrap_or(f64::NAN),
        None => f64::NAN,
    }
}

fn cookies_of(session: &Body) -> Vec<Value> {
    session
        .get("cookies")
        .and_then(Value::as_array)
        .cloned()
        .unwrap_or_default()
}

fn cookie_name(c: &Value) -> &str {
    c.get("name").and_then(Value::as_str).unwrap_or("")
}

/// `s.x || ""` keeping the saved value's own type.
fn or_empty(session: &Body, key: &str) -> Value {
    match session.get(key) {
        Some(v) if js::truthy(v) => v.clone(),
        _ => json!(""),
    }
}

/// Metadata only: safe to print.
pub fn session_status(ctx: &Ctx, name: &str) -> Value {
    let Some(s) = load_session(ctx, name) else {
        return json!({ "account": name, "saved": false });
    };
    let now = now_ms() as f64;
    let cookies = cookies_of(&s);
    let persistent = cookies
        .iter()
        .find(|c| cookie_name(c) == "ESTSAUTHPERSISTENT");
    let rt_obtained = s.get("refresh_token_obtained_at").filter(|v| js::truthy(v));
    let rt_age = rt_obtained.map(|v| now - parse_iso_ms(Some(v)));
    let has_rt = s.get("refresh_token").is_some_and(js::truthy);
    let expires_on = s
        .get("access_token")
        .and_then(|a| a.get("expires_on"))
        .filter(|v| js::truthy(v));
    let valid_min = match expires_on {
        Some(e) => num(0f64.max(((js::number(Some(e)) * 1000.0 - now) / 60000.0).round())),
        None => json!(0),
    };
    let hours_left = match rt_age {
        Some(age) if has_rt => {
            let h = ((REFRESH_TOKEN_MAX_AGE_MS - age) / 360000.0).round() / 10.0;
            if h.is_nan() {
                Value::Null
            } else {
                num(0f64.max(h))
            }
        }
        _ => json!(0),
    };
    let cookie_expires = match persistent {
        Some(p) => match p.get("expires").filter(|v| js::truthy(v)) {
            Some(e) => {
                let ms = (js::number(Some(e)) * 1000.0) as i64;
                chrono::DateTime::from_timestamp_millis(ms)
                    .map(|d| d.format("%Y-%m-%d").to_string())
                    .unwrap_or_default()
            }
            None => "session".into(),
        },
        None => String::new(),
    };
    json!({
        "account": name,
        "saved": true,
        "username": or_empty(&s, "username"),
        "tenant_id": or_empty(&s, "tenant_id"),
        "source": or_empty(&s, "source"),
        "access_token_valid_min": valid_min,
        "refresh_token": has_rt,
        "refresh_token_hours_left": hours_left,
        "sign_in_cookie": persistent.is_some(),
        "sign_in_cookie_expires": cookie_expires,
        "updated_at": or_empty(&s, "updated_at"),
    })
}

/// A token from the saved session (renewing it when needed), else from a re-import
/// from the account's Chrome. None when the account needs signing in again.
pub fn get_session_token(
    ctx: &Ctx,
    name: &str,
    allow_chrome_import: bool,
) -> Result<Option<SessionToken>> {
    let token = |t: String, source: String| SessionToken {
        token: t,
        api: "outlook_rest".into(),
        source,
    };
    if let Some(session) = load_session(ctx, name) {
        let at = session.get("access_token");
        let saved_token = at.and_then(|a| a.get("token")).filter(|v| js::truthy(v));
        let expires_on = js::number(at.and_then(|a| a.get("expires_on")));
        if let Some(t) = saved_token {
            if expires_on > now_sec_f() + ACCESS_TOKEN_MARGIN_SEC {
                return Ok(Some(token(js::string(t), "saved".into())));
            }
        }
        let fresh = refresh_session(ctx, name, Some(session), false)?;
        if fresh.ok {
            return Ok(Some(token(fresh.token, fresh.via)));
        }
    }
    if allow_chrome_import {
        if let Some(port) = get_account(ctx, name).and_then(|a| a.cdp_port.map(|p| (p, a))) {
            let (cdp_port, account) = port;
            let imported = import_from_chrome(ctx, name, cdp_port, Some(&account.main_url))?;
            if imported.get("ok") == Some(&json!(true)) {
                let fresh = refresh_session(ctx, name, None, false)?;
                if fresh.ok {
                    return Ok(Some(token(fresh.token, format!("chrome+{}", fresh.via))));
                }
            }
        }
    }
    Ok(None)
}

/// The outcome of a refresh. `to_json` never includes the token.
#[derive(Clone, Debug, Default)]
pub struct Refresh {
    pub ok: bool,
    pub account: String,
    pub token: String,
    pub via: String,
    pub status: String,
    pub errors: Vec<String>,
}

impl Refresh {
    pub fn to_json(&self) -> Value {
        if self.ok {
            json!({ "ok": true, "account": self.account, "via": self.via })
        } else if self.errors.is_empty() {
            json!({ "ok": false, "account": self.account, "status": self.status })
        } else {
            json!({ "ok": false, "account": self.account, "status": self.status, "errors": self.errors })
        }
    }
}

/// Mints a new access token from the refresh token, else from the sign-in cookies.
/// `force_cookies` skips the refresh token (the keepalive uses it to rotate cookies).
pub fn refresh_session(
    ctx: &Ctx,
    name: &str,
    session: Option<Body>,
    force_cookies: bool,
) -> Result<Refresh> {
    let Some(session) = session.or_else(|| load_session(ctx, name)) else {
        return Ok(Refresh {
            account: name.into(),
            status: "no_session".into(),
            ..Default::default()
        });
    };
    let mut errors = Vec::new();
    let rt = session.get("refresh_token").filter(|v| js::truthy(v));
    let rt_fresh = rt.is_some()
        && session
            .get("refresh_token_obtained_at")
            .is_some_and(js::truthy)
        && (now_ms() as f64 - parse_iso_ms(session.get("refresh_token_obtained_at")))
            < REFRESH_TOKEN_MAX_AGE_MS;
    if rt_fresh && !force_cookies {
        let rt = js::string(rt.unwrap());
        match redeem(
            ctx,
            &session,
            &[("grant_type", "refresh_token"), ("refresh_token", &rt)],
        ) {
            Ok(json) => return finish(ctx, name, session, &json, "refresh_token", true),
            Err(e) => errors.push(format!("refresh_token: {e}")),
        }
    }
    let cookies = cookies_of(&session);
    if cookies
        .iter()
        .any(|c| matches!(cookie_name(c), "ESTSAUTHPERSISTENT" | "ESTSAUTH"))
    {
        let r = silent_sign_in(ctx, &session);
        match r.result {
            Ok(json) => {
                let mut next = session.clone();
                next.insert(
                    "cookies".into(),
                    Value::Array(r.cookies.unwrap_or_default()),
                );
                return finish(ctx, name, next, &json, "cookies", false);
            }
            Err(e) => {
                if let Some(c) = r.cookies {
                    let mut next = session.clone();
                    next.insert("cookies".into(), Value::Array(c));
                    save_session(ctx, name, &next)?;
                }
                errors.push(format!("cookies: {e}"));
            }
        }
    }
    if errors.is_empty() {
        errors.push("nothing usable saved".into());
    }
    Ok(Refresh {
        account: name.into(),
        status: "needs_sign_in".into(),
        errors,
        ..Default::default()
    })
}

fn finish(
    ctx: &Ctx,
    name: &str,
    session: Body,
    json: &Value,
    via: &str,
    keep_refresh_token_age: bool,
) -> Result<Refresh> {
    let access_token = js::string(json.get("access_token").unwrap_or(&Value::Null));
    let claims = decode_jwt(&access_token);
    let id = decode_jwt(&json.get("id_token").map(js::string).unwrap_or_default());
    let t = |v: &Value, k: &str| v.get(k).filter(|x| js::truthy(x)).cloned();
    let mut next = session.clone();
    set_opt(
        &mut next,
        "tenant_id",
        t(&id, "tid")
            .or_else(|| t(&claims, "tid"))
            .or_else(|| session.get("tenant_id").cloned()),
    );
    let home = match (t(&id, "oid"), t(&id, "tid")) {
        (Some(oid), Some(tid)) => Some(json!(format!("{}.{}", js::string(&oid), js::string(&tid)))),
        _ => session.get("home_account_id").cloned(),
    };
    set_opt(&mut next, "home_account_id", home);
    let username = t(&id, "preferred_username")
        .or_else(|| t(&claims, "upn"))
        .or_else(|| t(&claims, "unique_name"))
        .or_else(|| session.get("username").filter(|v| js::truthy(v)).cloned())
        .unwrap_or_else(|| json!(""));
    next.insert("username".into(), username);
    let new_rt = t(json, "refresh_token");
    set_opt(
        &mut next,
        "refresh_token",
        new_rt
            .clone()
            .or_else(|| session.get("refresh_token").cloned()),
    );
    let saved_at = session.get("refresh_token_obtained_at").cloned();
    let obtained = if new_rt.is_some() {
        match &saved_at {
            Some(v) if keep_refresh_token_age && js::truthy(v) => Some(v.clone()),
            _ => Some(json!(now_iso())),
        }
    } else {
        saved_at
    };
    set_opt(&mut next, "refresh_token_obtained_at", obtained);
    let exp = js::number(claims.get("exp"));
    let expires_on = if exp.is_finite() && exp != 0.0 {
        num(exp)
    } else {
        let expires_in = match json.get("expires_in").filter(|v| js::truthy(v)) {
            Some(v) => js::number(Some(v)),
            None => 3600.0,
        };
        num((now_ms() / 1000) as f64 + expires_in)
    };
    next.insert(
        "access_token".into(),
        json!({
            "token": access_token,
            "expires_on": expires_on,
            "audience": js::str_or_empty(claims.get("aud")),
        }),
    );
    save_session(ctx, name, &next)?;
    Ok(Refresh {
        ok: true,
        account: name.into(),
        token: access_token,
        via: via.into(),
        ..Default::default()
    })
}

fn session_str(session: &Body, key: &str, default: &str) -> String {
    match session.get(key) {
        Some(v) if js::truthy(v) => js::string(v),
        _ => default.to_string(),
    }
}

/// Redeems a grant at the tenant's token endpoint with Outlook web's Origin.
fn redeem(ctx: &Ctx, session: &Body, params: &[(&str, &str)]) -> Result<Value, String> {
    let tenant = session_str(session, "tenant_id", "organizations");
    let client_id = session_str(session, "client_id", OUTLOOK_WEB_CLIENT_ID);
    let mut form = url::form_urlencoded::Serializer::new(String::new());
    form.append_pair("client_id", &client_id);
    form.append_pair("scope", OUTLOOK_SCOPE);
    for (k, v) in params {
        form.append_pair(k, v);
    }
    let res = ctx
        .http
        .post(format!("{}/{tenant}/oauth2/v2.0/token", login_base(ctx)))
        .header("origin", session_str(session, "origin", OUTLOOK_WEB_ORIGIN))
        .header("content-type", "application/x-www-form-urlencoded")
        .header("user-agent", BROWSER_UA)
        .body(form.finish())
        .timeout(LOGIN_TIMEOUT)
        .send()
        .map_err(|e| e.to_string())?;
    let status = res.status();
    if status.is_redirection() {
        return Err("redirect not allowed".into());
    }
    let json: Value = res
        .text()
        .ok()
        .and_then(|t| serde_json::from_str(&t).ok())
        .unwrap_or_else(|| json!({}));
    if status.is_success() && json.get("access_token").is_some_and(js::truthy) {
        return Ok(json);
    }
    let code = match json.get("error").filter(|v| js::truthy(v)) {
        Some(e) => js::string(e),
        None => format!("http_{}", status.as_u16()),
    };
    let desc = js::first_line(&js::str_or_empty(json.get("error_description")));
    Err(format!("{code} {desc}").trim().to_string())
}

struct SignIn {
    result: Result<Value, String>,
    cookies: Option<Vec<Value>>,
}

fn random_bytes(n: usize) -> Vec<u8> {
    let mut b = vec![0u8; n];
    getrandom::fill(&mut b).expect("random bytes");
    b
}

fn b64url(bytes: &[u8]) -> String {
    base64::engine::general_purpose::URL_SAFE_NO_PAD.encode(bytes)
}

fn hex(bytes: &[u8]) -> String {
    bytes.iter().map(|b| format!("{b:02x}")).collect()
}

/// The prompt=none authorization-code + PKCE flow with the saved sign-in cookies, exactly
/// as Outlook web renews in a hidden iframe. Returns rotated cookies either way.
fn silent_sign_in(ctx: &Ctx, session: &Body) -> SignIn {
    let tenant = session_str(session, "tenant_id", "organizations");
    let redirect_uri = session_str(session, "redirect_uri", OUTLOOK_WEB_REDIRECT_URI);
    let verifier = b64url(&random_bytes(32));
    let challenge = b64url(&Sha256::digest(verifier.as_bytes()));
    let mut q = url::form_urlencoded::Serializer::new(String::new());
    q.append_pair(
        "client_id",
        &session_str(session, "client_id", OUTLOOK_WEB_CLIENT_ID),
    );
    q.append_pair("response_type", "code");
    q.append_pair("redirect_uri", &redirect_uri);
    q.append_pair("scope", OUTLOOK_SCOPE);
    q.append_pair("prompt", "none");
    q.append_pair("response_mode", "fragment");
    q.append_pair("code_challenge", &challenge);
    q.append_pair("code_challenge_method", "S256");
    q.append_pair("state", &hex(&random_bytes(8)));
    q.append_pair("nonce", &hex(&random_bytes(8)));
    let username = session_str(session, "username", "");
    if !username.is_empty() {
        q.append_pair("login_hint", &username);
    }
    let cookies = cookies_of(session);
    let res = ctx
        .http
        .get(format!(
            "{}/{tenant}/oauth2/v2.0/authorize?{}",
            login_base(ctx),
            q.finish()
        ))
        .header("cookie", cookie_header(&cookies))
        .header("user-agent", BROWSER_UA)
        .timeout(LOGIN_TIMEOUT)
        .send();
    let res = match res {
        Ok(r) => r,
        Err(e) => {
            return SignIn {
                result: Err(e.to_string()),
                cookies: None,
            }
        }
    };
    let set_cookies: Vec<String> = res
        .headers()
        .get_all("set-cookie")
        .iter()
        .filter_map(|v| v.to_str().ok().map(str::to_string))
        .collect();
    let merged = merge_set_cookies(&cookies, &set_cookies);
    let location = res
        .headers()
        .get("location")
        .and_then(|v| v.to_str().ok())
        .unwrap_or("")
        .to_string();
    let part = if location.contains('#') {
        location.split('#').nth(1).unwrap_or("")
    } else {
        location.split('?').nth(1).unwrap_or("")
    };
    let params: Vec<(String, String)> = url::form_urlencoded::parse(part.as_bytes())
        .into_owned()
        .collect();
    let param = |k: &str| {
        params
            .iter()
            .find(|(n, _)| n == k)
            .map(|(_, v)| v.clone())
            .filter(|v| !v.is_empty())
    };
    let Some(code) = param("code") else {
        let err = param("error").unwrap_or_else(|| format!("http_{}", res.status().as_u16()));
        let desc = js::first_line(&param("error_description").unwrap_or_default());
        return SignIn {
            result: Err(format!("{err} {desc}").trim().to_string()),
            cookies: Some(merged),
        };
    };
    let r = redeem(
        ctx,
        session,
        &[
            ("grant_type", "authorization_code"),
            ("code", &code),
            ("redirect_uri", &redirect_uri),
            ("code_verifier", &verifier),
        ],
    );
    SignIn {
        result: r,
        cookies: Some(merged),
    }
}

const MSAL_EXPR: &str = r#"(() => {
      const out = { refresh: [], accounts: [] };
      for (let i = 0; i < localStorage.length; i++) {
        const k = localStorage.key(i) || "";
        if (!k.startsWith("msal.")) continue;
        let o; try { o = JSON.parse(localStorage.getItem(k) || ""); } catch { continue; }
        if (!o || typeof o !== "object") continue;
        if (o.credentialType === "RefreshToken" && typeof o.secret === "string") out.refresh.push({ secret: o.secret, clientId: o.clientId, homeAccountId: o.homeAccountId, environment: o.environment });
        else if (o.homeAccountId && o.username && o.realm) out.accounts.push({ homeAccountId: o.homeAccountId, username: o.username, realm: o.realm });
      }
      return out;
    })()"#;

fn origin_of(url: &str) -> String {
    match url::Url::parse(url) {
        Ok(u) if u.origin().is_tuple() => u.origin().ascii_serialization(),
        _ => OUTLOOK_WEB_ORIGIN.into(),
    }
}

/// Reads the refresh token and sign-in cookies from a signed-in Outlook web tab.
pub fn import_from_chrome(
    ctx: &Ctx,
    name: &str,
    cdp_port: u16,
    main_url: Option<&str>,
) -> Result<Value> {
    let Some(tab) = find_outlook_tab(ctx, cdp_port, main_url) else {
        return Ok(json!({
            "ok": false, "account": name, "status": "no_outlook_tab",
            "message": format!("No Outlook web tab on Chrome DevTools port {cdp_port}."),
        }));
    };
    let Some(mut cdp) = connect_cdp(&tab.web_socket_debugger_url) else {
        return Ok(json!({ "ok": false, "account": name, "status": "cdp_unreachable" }));
    };
    let msal = cdp.evaluate(MSAL_EXPR)?;
    let all = cdp.send("Network.getAllCookies", json!({}))?;
    drop(cdp);
    let cookies: Vec<Value> = all
        .get("cookies")
        .and_then(Value::as_array)
        .cloned()
        .unwrap_or_default()
        .iter()
        .filter(|c| LOGIN_DOMAIN_RE.is_match(&js::get_str(c, &["domain"])))
        .map(|c| {
            let expires = js::number(c.get("expires"));
            json!({
                "name": js::string(c.get("name").unwrap_or(&Value::Null)),
                "value": js::string(c.get("value").unwrap_or(&Value::Null)),
                "domain": js::get_str(c, &["domain"]),
                "path": match c.get("path") { Some(p) if js::truthy(p) => js::string(p), _ => "/".into() },
                "expires": if expires > 0.0 { num(expires.floor()) } else { Value::Null },
            })
        })
        .collect();
    let list = |k: &str| -> Vec<Value> {
        msal.get(k)
            .and_then(Value::as_array)
            .cloned()
            .unwrap_or_default()
    };
    let refreshes = list("refresh");
    let accounts = list("accounts");
    let refresh = refreshes
        .iter()
        .find(|r| r.get("clientId").and_then(Value::as_str) == Some(OUTLOOK_WEB_CLIENT_ID))
        .or_else(|| refreshes.first())
        .cloned();
    let rget = |k: &str| {
        refresh
            .as_ref()
            .and_then(|r| r.get(k))
            .filter(|v| js::truthy(v))
            .cloned()
    };
    let home = rget("homeAccountId");
    let acct = accounts
        .iter()
        .find(|a| home.is_some() && a.get("homeAccountId") == home.as_ref())
        .or_else(|| accounts.first())
        .cloned();
    let aget = |k: &str| {
        acct.as_ref()
            .and_then(|a| a.get(k))
            .filter(|v| js::truthy(v))
            .cloned()
    };
    if refresh.is_none()
        && !cookies
            .iter()
            .any(|c| cookie_name(c).starts_with("ESTSAUTH"))
    {
        return Ok(json!({
            "ok": false, "account": name, "status": "not_signed_in",
            "message": "The Outlook tab has no saved sign-in. Sign in to Outlook web in that Chrome first.",
        }));
    }
    let previous = load_session(ctx, name).unwrap_or_default();
    let prev = |k: &str| previous.get(k).filter(|v| js::truthy(v)).cloned();
    let tenant_from_home = home
        .as_ref()
        .map(js::string)
        .and_then(|h| h.split('.').nth(1).map(str::to_string))
        .filter(|t| !t.is_empty())
        .map(Value::String);
    let secret = rget("secret");
    let mut next = previous.clone();
    next.insert("source".into(), json!(format!("chrome:{cdp_port}")));
    next.insert(
        "client_id".into(),
        rget("clientId").unwrap_or_else(|| json!(OUTLOOK_WEB_CLIENT_ID)),
    );
    next.insert("origin".into(), json!(origin_of(&tab.url)));
    next.insert("redirect_uri".into(), json!(OUTLOOK_WEB_REDIRECT_URI));
    next.insert(
        "tenant_id".into(),
        aget("realm")
            .or(tenant_from_home)
            .or_else(|| prev("tenant_id"))
            .unwrap_or_else(|| json!("")),
    );
    next.insert(
        "home_account_id".into(),
        home.clone()
            .or_else(|| prev("home_account_id"))
            .unwrap_or_else(|| json!("")),
    );
    next.insert(
        "username".into(),
        aget("username")
            .or_else(|| prev("username"))
            .unwrap_or_else(|| json!("")),
    );
    next.insert(
        "refresh_token".into(),
        secret.clone().unwrap_or_else(|| json!("")),
    );
    // Outlook web renews its refresh token often, so a freshly read one is at most a day
    // old; count it as new and let the cookie sign-in take over if Microsoft rejects it.
    next.insert(
        "refresh_token_obtained_at".into(),
        if secret.is_some() {
            json!(now_iso())
        } else {
            Value::Null
        },
    );
    next.insert("cookies".into(), Value::Array(cookies));
    next.insert("access_token".into(), Value::Null);
    save_session(ctx, name, &next)?;
    let mut out = Map::new();
    out.insert("ok".into(), json!(true));
    out.insert("account".into(), json!(name));
    js::assign(&mut out, session_status(ctx, name));
    Ok(Value::Object(out))
}

/// Imports cookies exported from a browser (Cookie-Editor / EditThisCookie JSON, or a
/// Netscape cookies.txt). Only login.microsoftonline.com cookies are kept.
pub fn import_cookies(
    ctx: &Ctx,
    name: &str,
    text: &str,
    username: &str,
    tenant: &str,
) -> Result<Value> {
    let cookies: Vec<Value> = parse_cookie_export(text)?
        .into_iter()
        .filter(|c| LOGIN_DOMAIN_RE.is_match(&js::get_str(c, &["domain"])))
        .collect();
    if !cookies
        .iter()
        .any(|c| matches!(cookie_name(c), "ESTSAUTHPERSISTENT" | "ESTSAUTH"))
    {
        return Ok(json!({
            "ok": false, "account": name, "status": "no_sign_in_cookie",
            "message": "No ESTSAUTHPERSISTENT or ESTSAUTH cookie for login.microsoftonline.com in that export.",
        }));
    }
    let previous = load_session(ctx, name).unwrap_or_default();
    let pick = |given: &str, key: &str| {
        if given.is_empty() {
            previous
                .get(key)
                .filter(|v| js::truthy(v))
                .cloned()
                .unwrap_or_else(|| json!(""))
        } else {
            json!(given)
        }
    };
    let mut next = previous.clone();
    next.insert("source".into(), json!("cookies"));
    next.insert("client_id".into(), json!(OUTLOOK_WEB_CLIENT_ID));
    next.insert("origin".into(), json!(OUTLOOK_WEB_ORIGIN));
    next.insert("redirect_uri".into(), json!(OUTLOOK_WEB_REDIRECT_URI));
    next.insert("tenant_id".into(), pick(tenant, "tenant_id"));
    next.insert("username".into(), pick(username, "username"));
    next.insert("refresh_token".into(), json!(""));
    next.insert("refresh_token_obtained_at".into(), Value::Null);
    let count = cookies.len();
    next.insert("cookies".into(), Value::Array(cookies));
    next.insert("access_token".into(), Value::Null);
    save_session(ctx, name, &next)?;
    Ok(json!({ "ok": true, "account": name, "imported_cookies": count }))
}

/// Parses Cookie-Editor JSON (an array, or `{cookies: [...]}`) or a Netscape cookies.txt
/// into `{name, value, domain, path, expires}` objects.
pub fn parse_cookie_export(text: &str) -> Result<Vec<Value>> {
    let s = text.trim();
    if s.starts_with('[') || s.starts_with('{') {
        let parsed: Value = serde_json::from_str(s)?;
        let list = match &parsed {
            Value::Array(a) => a.clone(),
            other => other
                .get("cookies")
                .and_then(Value::as_array)
                .cloned()
                .unwrap_or_default(),
        };
        return Ok(list
            .iter()
            .map(|c| {
                let exp = match js::or(
                    c.as_object().unwrap_or(&Map::new()),
                    &["expirationDate", "expires"],
                ) {
                    Some(v) => js::number(Some(&v)),
                    None => 0.0,
                };
                json!({
                    "name": js::get_str(c, &["name"]),
                    "value": js::get_str(c, &["value"]),
                    "domain": js::get_str(c, &["domain"]).to_lowercase(),
                    "path": match c.get("path") { Some(p) if js::truthy(p) => js::string(p), _ => "/".into() },
                    "expires": if exp > 0.0 { num(exp.floor()) } else { Value::Null },
                })
            })
            .filter(|c| !cookie_name(c).is_empty())
            .collect());
    }
    Ok(s.split('\n')
        .map(|l| l.strip_suffix('\r').unwrap_or(l))
        .map(|l| l.strip_prefix("#HttpOnly_").unwrap_or(l))
        .filter(|l| !l.is_empty() && !l.starts_with('#'))
        .map(|l| l.split('\t').collect::<Vec<_>>())
        .filter(|f| f.len() >= 7)
        .map(|f| {
            let exp = js::number(Some(&json!(f[4])));
            json!({
                "name": f[5],
                "value": f[6],
                "domain": f[0].to_lowercase(),
                "path": if f[2].is_empty() { "/" } else { f[2] },
                "expires": if exp > 0.0 { num(exp) } else { Value::Null },
            })
        })
        .collect())
}

fn cookie_header(cookies: &[Value]) -> String {
    let now = now_sec_f();
    cookies
        .iter()
        .filter(|c| c.get("value").is_some_and(js::truthy))
        .filter(|c| match c.get("expires").filter(|v| js::truthy(v)) {
            Some(e) => js::number(Some(e)) > now,
            None => true,
        })
        .map(|c| format!("{}={}", cookie_name(c), js::get_str(c, &["value"])))
        .collect::<Vec<_>>()
        .join("; ")
}

fn parse_cookie_date(s: &str) -> Option<i64> {
    let s = s.trim();
    chrono::DateTime::parse_from_rfc2822(s)
        .or_else(|_| chrono::DateTime::parse_from_rfc2822(&s.replace('-', " ")))
        .or_else(|_| chrono::DateTime::parse_from_rfc3339(s))
        .ok()
        .map(|d| d.timestamp())
}

/// Applies Set-Cookie headers to the saved cookies: replaced values, new cookies,
/// and removal of cleared or expired ones.
pub fn merge_set_cookies(existing: &[Value], headers: &[String]) -> Vec<Value> {
    let key = |c: &Value| format!("{}|{}", cookie_name(c), js::get_str(c, &["domain"]));
    let mut jar: Vec<(String, Value)> = existing.iter().map(|c| (key(c), c.clone())).collect();
    for header in headers {
        let mut parts = header.split(';');
        let pair = parts.next().unwrap_or("");
        let Some(eq) = pair.find('=').filter(|i| *i > 0) else {
            continue;
        };
        let name = pair[..eq].trim().to_string();
        let value = pair[eq + 1..].trim().to_string();
        let mut attrs: Vec<(String, Option<String>)> = Vec::new();
        for a in parts {
            let entry = match a.find('=') {
                None => (a.trim().to_lowercase(), None),
                Some(i) => (
                    a[..i].trim().to_lowercase(),
                    Some(a[i + 1..].trim().to_string()),
                ),
            };
            attrs.retain(|(k, _)| *k != entry.0);
            attrs.push(entry);
        }
        let attr = |k: &str| attrs.iter().find(|(n, _)| n == k).map(|(_, v)| v.clone());
        let domain = match attr("domain") {
            Some(Some(d)) if !d.is_empty() => d.to_lowercase(),
            _ => "login.microsoftonline.com".into(),
        };
        let now = now_sec_f();
        let expires: Option<i64> = if let Some(max_age) = attr("max-age") {
            max_age
                .and_then(|m| m.trim().parse::<f64>().ok())
                .map(|m| (now.floor() + m) as i64)
        } else {
            match attr("expires") {
                Some(Some(e)) if !e.is_empty() => parse_cookie_date(&e).filter(|t| *t != 0),
                _ => None,
            }
        };
        // Same name on the bare or dotted domain is the same cookie for our purposes.
        let bare = domain.trim_start_matches('.').to_string();
        jar.retain(|(k, _)| {
            let mut it = k.splitn(2, '|');
            let (n, d) = (it.next().unwrap_or(""), it.next().unwrap_or(""));
            !(n == name && d.strip_prefix('.').unwrap_or(d) == bare)
        });
        if value.is_empty() || expires.is_some_and(|e| (e as f64) <= now) {
            continue;
        }
        let path = match attr("path") {
            Some(Some(p)) if !p.is_empty() => p,
            Some(None) => "true".into(),
            _ => "/".into(),
        };
        jar.push((
            format!("{name}|{domain}"),
            json!({ "name": name, "value": value, "domain": domain, "path": path, "expires": expires }),
        ));
    }
    jar.into_iter().map(|(_, c)| c).collect()
}

/// The claims of a JWT (unverified), or `{}`.
pub fn decode_jwt(token: &str) -> Value {
    let Some(payload) = token.split('.').nth(1) else {
        return json!({});
    };
    let normalized: String = payload
        .chars()
        .map(|c| match c {
            '+' => '-',
            '/' => '_',
            c => c,
        })
        .filter(|c| *c != '=')
        .collect();
    base64::engine::general_purpose::URL_SAFE_NO_PAD
        .decode(normalized.as_bytes())
        .ok()
        .and_then(|b| serde_json::from_slice::<Value>(&b).ok())
        .filter(Value::is_object)
        .unwrap_or_else(|| json!({}))
}

/// Renews every account (or the named ones): a fresh token, and with `rotate` also
/// renewed sign-in cookies. An account whose session is dead is re-imported from its
/// Chrome port. Safe to print: no tokens or cookies in the result.
pub fn refresh_accounts(ctx: &Ctx, accounts: Option<Vec<String>>, rotate: bool) -> Result<Value> {
    let names = match accounts {
        Some(a) if !a.is_empty() => a,
        _ => list_accounts(ctx).into_iter().map(|a| a.name).collect(),
    };
    let mut results = Vec::new();
    for name in &names {
        let mut r = refresh_session(ctx, name, None, rotate)?.to_json();
        if r.get("ok") != Some(&json!(true)) {
            if let Some(account) = get_account(ctx, name) {
                if let Some(port) = account.cdp_port {
                    let imported = import_from_chrome(ctx, name, port, Some(&account.main_url))?;
                    if imported.get("ok") == Some(&json!(true)) {
                        r = refresh_session(ctx, name, None, false)?.to_json();
                        r["reimported_from_chrome"] = json!(true);
                    }
                }
            }
        }
        let mut entry = r.as_object().cloned().unwrap_or_default();
        js::assign(&mut entry, session_status(ctx, name));
        results.push(Value::Object(entry));
    }
    let all_ok = !results.is_empty() && results.iter().all(|r| r.get("ok") == Some(&json!(true)));
    Ok(json!({ "ok": all_ok, "status": "ok", "accounts": results }))
}
