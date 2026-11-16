//! Port of test/msal-parse.test.mjs: picking a live token from Outlook web's MSAL cache.

use outlook_web_cli::outlook::pick_best_token;
use serde_json::{json, Value};

fn mk(target: &str, exp_offset: i64) -> Value {
    let now = chrono::Utc::now().timestamp();
    json!({ "key": "msal.3|...|accesstoken|...", "source": "msal_v3", "target": target, "expiresOn": now + exp_offset, "token": "x".repeat(200) })
}

#[test]
fn prefers_outlook_rest_for_mail() {
    let e = [
        mk(
            "https://graph.microsoft.com/Mail.Read email openid profile",
            3600,
        ),
        mk(
            "https://outlook.office.com/Mail.Read https://outlook.office.com/Calendars.Read",
            7200,
        ),
    ];
    assert_eq!(pick_best_token(&e, "mail").unwrap().api, "outlook_rest");
}

#[test]
fn falls_back_to_graph_when_outlook_lacks_scope() {
    let e = [
        mk("https://outlook.office.com/Calendars.Read", 7200),
        mk("https://graph.microsoft.com/Mail.Read", 3600),
    ];
    assert_eq!(pick_best_token(&e, "mail").unwrap().api, "graph");
    // Drafts are written through Outlook REST only.
    assert_eq!(pick_best_token(&e, "mail_write"), None);
}

#[test]
fn ignores_expired_and_short_tokens() {
    let e = [
        mk("https://outlook.office.com/Mail.Read", -10),
        mk("https://graph.microsoft.com/Mail.Read", -10),
    ];
    assert_eq!(pick_best_token(&e, "mail"), None);
    let mut short = mk("https://outlook.office.com/Mail.Read", 3600);
    short["token"] = json!("abc");
    assert_eq!(pick_best_token(&[short], "mail"), None);
}

#[test]
fn calendar_scope() {
    let e = [
        mk("https://outlook.office.com/Calendars.Read", 600),
        mk("https://graph.microsoft.com/Calendars.Read", 3600),
    ];
    assert_eq!(pick_best_token(&e, "calendar").unwrap().api, "outlook_rest");
}
