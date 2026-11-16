//! The `outlook-web-cli` binary: help, flags and `--json` bodies, JSON output and exit
//! codes, and the auth commands end to end (cookies, Chrome over a fake DevTools
//! endpoint, status without secrets, refresh, logout, accounts).

mod common;

use assert_cmd::Command;
use common::*;
use serde_json::{json, Value};
use std::net::TcpListener;

fn cli(dir: &tempfile::TempDir) -> Command {
    let mut c = Command::new(assert_cmd::cargo::cargo_bin!("outlook-web-cli"));
    c.env_clear()
        .env("HOME", dir.path())
        .env("OUTLOOK_WEB_CONFIG_DIR", dir.path());
    c
}

fn json_out(out: &std::process::Output) -> Value {
    serde_json::from_slice(&out.stdout)
        .unwrap_or_else(|_| panic!("{}", String::from_utf8_lossy(&out.stdout)))
}

#[test]
fn help_lists_every_command() {
    let dir = config_dir(json!([]));
    for args in [vec![], vec!["--help"], vec!["help"]] {
        let out = cli(&dir).args(&args).output().unwrap();
        assert!(out.status.success(), "{args:?}");
        let text = String::from_utf8(out.stdout).unwrap();
        for phrase in [
            "auth chrome",
            "auth cookies",
            "freebusy",
            "draft-reply",
            "attachment-inspect",
            "--json",
            "nothing ever sends mail",
        ] {
            assert!(text.contains(phrase), "{phrase} missing from {args:?}");
        }
    }
    for sub in ["messages", "freebusy", "draft-new", "auth"] {
        cli(&dir).args([sub, "--help"]).assert().success();
    }
    let v = cli(&dir).arg("--version").output().unwrap();
    assert!(String::from_utf8_lossy(&v.stdout).contains("1.0.0"));
}

#[test]
fn flags_and_json_bodies_reach_the_api() {
    let server = Fake::start(|req| {
        if req.path.starts_with("/api/v2.0/me/MailFolders/") {
            return Reply::json(200, json!({ "value": [] }));
        }
        Reply::not_found()
    });
    let dir = two_accounts();
    let token = rest_token();
    let run = |args: &[&str]| {
        cli(&dir)
            .env("OUTLOOK_REST_BASE_URL", &server.base)
            .env("OUTLOOK_TEST_TOKEN", &token)
            .args(args)
            .output()
            .unwrap()
    };
    // Flags: --kebab-case, --flag=value, bare booleans.
    let out = run(&[
        "messages",
        "--account",
        "contoso",
        "--limit=3",
        "--unread",
        "--from",
        "pat@contoso.com",
        "--folder",
        "sent",
    ]);
    assert!(out.status.success());
    let v = json_out(&out);
    assert_eq!(v["folder"], json!("SentItems"));
    let req = server.last();
    assert_eq!(req.path, "/api/v2.0/me/MailFolders/SentItems/messages");
    assert_eq!(req.q("$top").as_deref(), Some("3"));
    assert_eq!(
        req.q("$filter").as_deref(),
        Some("ReceivedDateTime ge 1900-01-01T00:00:00Z and IsRead eq false and From/EmailAddress/Address eq 'pat@contoso.com'")
    );
    // --json is merged over the flags (snake_case fields, real types).
    run(&[
        "messages",
        "--account",
        "contoso",
        "--limit",
        "9",
        "--json",
        r#"{"limit": 4, "folder": "drafts"}"#,
    ]);
    assert_eq!(server.last().q("$top").as_deref(), Some("4"));
    assert_eq!(
        server.last().path,
        "/api/v2.0/me/MailFolders/Drafts/messages"
    );
    // Aliases.
    run(&[
        "messages",
        "--account",
        "fabrikam.com",
        "--max-results",
        "2",
    ]);
    assert_eq!(server.last().q("$top").as_deref(), Some("2"));

    // ok=false -> exit 1, and the JSON still prints.
    let bad = run(&["messages", "--account", "contoso", "--folder", "everything"]);
    assert_eq!(bad.status.code(), Some(1));
    assert_eq!(json_out(&bad)["status"], json!("invalid_folder"));
    // Draft commands do nothing without --yes.
    let pending = run(&[
        "draft-new",
        "--account",
        "contoso",
        "--to",
        "a@contoso.com",
        "--subject",
        "s",
        "--body",
        "b",
    ]);
    assert_eq!(pending.status.code(), Some(1));
    assert_eq!(json_out(&pending)["status"], json!("confirmation_required"));
    // Errors become {ok:false, status:"error"}.
    let err = run(&["messages", "--account", "gmail"]);
    assert_eq!(err.status.code(), Some(1));
    let e = json_out(&err);
    assert_eq!(e["status"], json!("error"));
    assert!(e["message"]
        .as_str()
        .unwrap()
        .contains("Unknown account 'gmail'"));
    let bad_json = run(&["status", "--json", "{nope"]);
    assert_eq!(json_out(&bad_json)["status"], json!("error"));
    assert!(server.requests().iter().all(|r| r.method == "GET"));
}

/// A fake Chrome: /json/list over HTTP and a DevTools websocket answering
/// Runtime.evaluate (the MSAL cache) and Network.getAllCookies.
fn fake_chrome() -> (Fake, std::thread::JoinHandle<()>) {
    let ws = TcpListener::bind("127.0.0.1:0").unwrap();
    let ws_port = ws.local_addr().unwrap().port();
    let handle = std::thread::spawn(move || {
        let (stream, _) = ws.accept().unwrap();
        let mut sock = tungstenite::accept(stream).unwrap();
        loop {
            let msg = match sock.read() {
                Ok(tungstenite::Message::Text(t)) => t.to_string(),
                Ok(_) => continue,
                Err(_) => break,
            };
            let req: Value = serde_json::from_str(&msg).unwrap();
            let result = match req["method"].as_str().unwrap() {
                "Runtime.evaluate" => json!({ "result": { "type": "object", "value": {
                    "refresh": [{ "secret": "rt-chrome", "clientId": "9199bf20-a13f-4107-85dc-02114787ef48", "homeAccountId": "oid-1.tenant-1", "environment": "login.microsoftonline.com" }],
                    "accounts": [{ "homeAccountId": "oid-1.tenant-1", "username": "me@contoso.com", "realm": "tenant-1" }],
                } } }),
                "Network.getAllCookies" => json!({ "cookies": [
                    { "name": "ESTSAUTHPERSISTENT", "value": "cookie-v1", "domain": ".login.microsoftonline.com", "path": "/", "expires": 1802000000.5 },
                    { "name": "SID", "value": "x", "domain": ".google.com", "path": "/", "expires": -1 },
                ] }),
                _ => json!({}),
            };
            sock.send(tungstenite::Message::text(
                json!({ "id": req["id"], "result": result }).to_string(),
            ))
            .unwrap();
        }
    });
    let http = Fake::start(move |req| {
        if req.path == "/json/list" {
            return Reply::json(
                200,
                json!([
                    { "type": "service_worker", "url": "https://outlook.cloud.microsoft/sw.js", "webSocketDebuggerUrl": "ws://127.0.0.1:1/x" },
                    { "type": "page", "url": "https://outlook.cloud.microsoft/mail/inbox", "webSocketDebuggerUrl": format!("ws://127.0.0.1:{ws_port}/devtools/page/1") },
                ]),
            );
        }
        Reply::not_found()
    });
    (http, handle)
}

fn fake_login() -> Fake {
    Fake::start(|req| {
        if req.path.ends_with("/oauth2/v2.0/token")
            && req.header("origin").as_deref() == Some("https://outlook.cloud.microsoft")
        {
            let exp = chrono::Utc::now().timestamp() + 3600;
            return Reply::json(
                200,
                json!({
                    "access_token": jwt(json!({ "aud": "https://outlook.office.com", "tid": "tenant-1", "exp": exp })),
                    "refresh_token": "rt-next",
                    "id_token": jwt(json!({ "tid": "tenant-1", "oid": "oid-1", "preferred_username": "me@contoso.com" })),
                }),
            );
        }
        if req.path.ends_with("/oauth2/v2.0/authorize") {
            return Reply::json(302, json!({})).header(
                "location",
                "https://outlook.cloud.microsoft/mail/#error=login_required",
            );
        }
        Reply::json(400, json!({ "error": "invalid_request" }))
    })
}

#[test]
fn auth_commands_end_to_end() {
    let login = fake_login();
    let (chrome, ws_thread) = fake_chrome();
    let chrome_port = chrome.base.rsplit(':').next().unwrap().to_string();
    let dir = config_dir(json!([]));
    std::fs::remove_file(dir.path().join("config.json")).unwrap();
    let run = |args: &[&str]| {
        cli(&dir)
            .env("OUTLOOK_WEB_LOGIN_BASE_URL", &login.base)
            .args(args)
            .output()
            .unwrap()
    };

    // auth chrome: reads the MSAL refresh token + login cookies, then verifies with a token.
    let out = run(&[
        "auth",
        "chrome",
        "--account",
        "Work",
        "--cdp-port",
        &chrome_port,
    ]);
    let v = json_out(&out);
    assert!(out.status.success(), "{v}");
    assert_eq!(v["account"], json!("work"));
    assert_eq!(v["verified"], json!("token minted via refresh_token"));
    assert_eq!(v["username"], json!("me@contoso.com"));
    assert_eq!(v["source"], json!(format!("chrome:{chrome_port}")));
    ws_thread.join().unwrap();
    let saved: Value = serde_json::from_str(
        &std::fs::read_to_string(dir.path().join("sessions/work.json")).unwrap(),
    )
    .unwrap();
    let keys: Vec<&str> = saved
        .as_object()
        .unwrap()
        .keys()
        .map(String::as_str)
        .collect();
    assert_eq!(
        keys,
        [
            "source",
            "client_id",
            "origin",
            "redirect_uri",
            "tenant_id",
            "home_account_id",
            "username",
            "refresh_token",
            "refresh_token_obtained_at",
            "cookies",
            "access_token",
            "version",
            "account",
            "updated_at"
        ]
    );
    assert_eq!(saved["refresh_token"], json!("rt-next"));
    assert_eq!(saved["tenant_id"], json!("tenant-1"));
    assert_eq!(saved["home_account_id"], json!("oid-1.tenant-1"));
    assert_eq!(
        saved["cookies"],
        json!([{ "name": "ESTSAUTHPERSISTENT", "value": "cookie-v1", "domain": ".login.microsoftonline.com", "path": "/", "expires": 1802000000 }])
    );
    assert_eq!(
        saved["access_token"]["audience"],
        json!("https://outlook.office.com")
    );
    // The Chrome port is remembered in config.json (mode 0600).
    let config: Value =
        serde_json::from_str(&std::fs::read_to_string(dir.path().join("config.json")).unwrap())
            .unwrap();
    assert_eq!(
        config,
        json!({ "accounts": [{ "name": "work", "cdp_port": chrome_port.parse::<u16>().unwrap(), "username": "me@contoso.com" }] })
    );
    #[cfg(unix)]
    {
        use std::os::unix::fs::PermissionsExt;
        assert_eq!(
            std::fs::metadata(dir.path().join("config.json"))
                .unwrap()
                .permissions()
                .mode()
                & 0o777,
            0o600
        );
        assert_eq!(
            std::fs::metadata(dir.path().join("sessions/work.json"))
                .unwrap()
                .permissions()
                .mode()
                & 0o777,
            0o600
        );
    }

    // auth status and accounts never print secrets.
    let status = run(&["auth", "status"]);
    let s = json_out(&status);
    assert_eq!(s["accounts"][0]["account"], json!("work"));
    assert_eq!(s["accounts"][0]["refresh_token"], json!(true));
    assert_eq!(
        s["accounts"][0]["sign_in_cookie_expires"],
        json!("2027-02-07")
    );
    assert_eq!(
        s["accounts"][0]["cdp_port"],
        json!(chrome_port.parse::<u16>().unwrap())
    );
    let accounts = json_out(&run(&["accounts"]));
    assert_eq!(
        accounts["accounts"][0],
        json!({ "account": "work", "username": "me@contoso.com", "aliases": [], "cdp_port": chrome_port.parse::<u16>().unwrap(), "saved_session": true })
    );
    let secrets = regex::Regex::new(r"rt-|cookie-v|eyJ").unwrap();
    for out in [&status.stdout, &run(&["auth", "refresh", "--all"]).stdout] {
        assert!(
            !secrets.is_match(&String::from_utf8_lossy(out)),
            "no secrets"
        );
    }

    // auth cookies from a Netscape cookies.txt on stdin: the silent sign-in is refused here,
    // so it reports why and exits 1, but the cookies are saved.
    let txt = "#HttpOnly_.login.microsoftonline.com\tTRUE\t/\tTRUE\t1802000000\tESTSAUTHPERSISTENT\tcookie-v5\n";
    let out = cli(&dir)
        .env("OUTLOOK_WEB_LOGIN_BASE_URL", &login.base)
        .args([
            "auth",
            "cookies",
            "--account",
            "other",
            "--file",
            "-",
            "--username",
            "me@fabrikam.com",
        ])
        .write_stdin(txt)
        .output()
        .unwrap();
    let v = json_out(&out);
    assert_eq!(out.status.code(), Some(1));
    assert_eq!(v["sign_in_cookie"], json!(true));
    assert!(v["verified"][0]
        .as_str()
        .unwrap()
        .contains("login_required"));
    assert!(!secrets.is_match(&v.to_string()));

    // logout removes the session.
    let gone = json_out(&run(&["auth", "logout", "--account", "other"]));
    assert_eq!(
        gone,
        json!({ "ok": true, "account": "other", "removed": true })
    );
    // --account is required for logout.
    assert_eq!(
        json_out(&run(&["auth", "logout"]))["status"],
        json!("error")
    );
}
