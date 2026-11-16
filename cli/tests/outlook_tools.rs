//! Port of test/outlook-tools.test.mjs: status, messages (Graph and Outlook REST,
//! folders, filters), read, search, agenda, and account aliases.

mod common;

use common::*;
use outlook_web_cli::accounts::resolve_account_alias;
use outlook_web_cli::commands::run;
use serde_json::{json, Value};

fn graph_msg(id: &str, subject: &str, name: &str, at: &str, read: bool) -> Value {
    json!({
        "id": id, "subject": subject,
        "from": { "emailAddress": { "name": name, "address": format!("{}@example.com", name.to_lowercase()) } },
        "receivedDateTime": at, "isRead": read, "conversationId": "C",
        "webLink": format!("https://outlook.office.com/mail/deeplink/read/{id}"),
    })
}

fn rest_msg(id: &str, subject: &str, name: &str, at: &str, read: bool) -> Value {
    json!({
        "Id": id, "Subject": subject,
        "From": { "EmailAddress": { "Name": name, "Address": format!("{}@example.com", name.to_lowercase()) } },
        "ReceivedDateTime": at, "IsRead": read, "ConversationId": "C",
        "WebLink": format!("https://outlook.office.com/mail/deeplink/read/{id}"),
    })
}

/// A fake Graph + Outlook REST server.
fn fake() -> Fake {
    Fake::start(|req| {
        let p = req.path.as_str();
        if p == "/v1.0/me/mailFolders/Inbox/messages" {
            return Reply::json(
                200,
                json!({ "value": [
                graph_msg("MSG_A", "Hello A", "Alice", "2026-10-09T16:00:00Z", false),
                graph_msg("MSG_B", "Hello B", "Bob", "2026-10-09T15:00:00Z", true),
            ] }),
            );
        }
        if let Some(id) = p.strip_prefix("/v1.0/me/messages/") {
            return Reply::json(
                200,
                json!({
                    "id": id, "subject": "Read Test",
                    "from": { "emailAddress": { "name": "Carol", "address": "carol@example.com" } },
                    "receivedDateTime": "2026-10-09T17:00:00Z",
                    "body": { "contentType": "html", "content": "<p>Line 1<br>Line 2</p>" },
                    "webLink": "https://outlook.office.com/mail/deeplink/read/READ_ID",
                }),
            );
        }
        if p == "/v1.0/me/messages" && req.q("$search").is_some() {
            return Reply::json(
                200,
                json!({ "value": [graph_msg("MSG_S", "Search Hit", "Dana", "2026-10-09T18:00:00Z", false)] }),
            );
        }
        if p == "/v1.0/me/calendarView" {
            return Reply::json(
                200,
                json!({ "value": [{
                "id": "EV1", "subject": "Standup", "isAllDay": false,
                "start": { "dateTime": "2026-10-09T09:00:00", "timeZone": "America/New_York" },
                "end": { "dateTime": "2026-10-09T09:30:00", "timeZone": "America/New_York" },
                "organizer": { "emailAddress": { "name": "Pat", "address": "pat@example.com" } },
                "location": { "displayName": "Zoom" },
                "webLink": "https://outlook.office.com/calendar/item/EV1",
            }] }),
            );
        }
        // Outlook REST folder listing: Sent Items holds a sent reply, Drafts an unsent one.
        if let Some(folder) = p
            .strip_prefix("/api/v2.0/me/MailFolders/")
            .and_then(|r| r.strip_suffix("/messages"))
        {
            return Reply::json(
                200,
                json!({ "value": [{
                "Id": format!("MSG_{folder}"), "Subject": "RE: Duplicate checks",
                "From": { "EmailAddress": { "Name": "AF (Shared)", "Address": "me@fabrikam.com" } },
                "ReceivedDateTime": "2026-11-11T19:06:05Z", "IsRead": true, "IsDraft": folder == "Drafts",
                "ConversationId": "C9", "WebLink": format!("https://outlook.office.com/mail/deeplink/read/MSG_{folder}"),
            }] }),
            );
        }
        if p == "/api/v2.0/me/messages" {
            return Reply::json(
                200,
                json!({ "value": [
                rest_msg("MSG_A", "Hello A", "Alice", "2026-10-09T16:00:00Z", false),
                rest_msg("MSG_B", "Hello B", "Bob", "2026-10-09T15:00:00Z", true),
            ] }),
            );
        }
        if let Some(id) = p.strip_prefix("/api/v2.0/me/messages/") {
            return Reply::json(
                200,
                json!({
                    "Id": id, "Subject": "Read Test",
                    "From": { "EmailAddress": { "Name": "Carol", "Address": "carol@example.com" } },
                    "ReceivedDateTime": "2026-10-09T17:00:00Z",
                    "Body": { "ContentType": "HTML", "Content": "<p>Line 1<br>Line 2</p>" },
                }),
            );
        }
        if p == "/api/v2.0/me/calendarview" {
            return Reply::json(
                200,
                json!({ "value": [{
                "Id": "EV1", "Subject": "Standup", "IsAllDay": false,
                "Start": { "DateTime": "2026-10-09T09:00:00", "TimeZone": "America/New_York" },
                "End": { "DateTime": "2026-10-09T09:30:00", "TimeZone": "America/New_York" },
                "Organizer": { "EmailAddress": { "Name": "Pat", "Address": "pat@example.com" } },
                "Location": { "DisplayName": "Zoom" },
            }] }),
            );
        }
        Reply::not_found()
    })
}

#[test]
fn mail_and_calendar() {
    let server = fake();
    let dir = two_accounts();
    let graph = format!(
        "x.{}.y",
        b64(&json!({ "aud": "https://graph.microsoft.com", "scp": "Mail.Read" }))
    );
    let rest = format!(
        "x.{}.y",
        b64(&json!({ "aud": "https://outlook.office.com", "scp": "Mail.Read" }))
    );
    let base = server.base.clone();
    let vars = [
        ("OUTLOOK_GRAPH_BASE_URL", base.as_str()),
        ("OUTLOOK_REST_BASE_URL", base.as_str()),
        ("OUTLOOK_TEST_TOKEN", graph.as_str()),
    ];
    let ctx = ctx(&dir, &vars);
    let r = |cmd: &str, b: Value| run(&ctx, cmd, &body(b), false).unwrap();

    // status
    let s = r("status", json!({ "account": "contoso" }));
    assert_eq!(s["ok"], json!(true));
    assert_eq!(s["accounts"][0]["status"], json!("signed_in"));
    assert_eq!(s["answer_text"], json!("Signed in."));

    // messages through Graph (flags arrive as strings)
    let m = r(
        "messages",
        json!({ "account": "contoso", "limit": "2", "unread": "false" }),
    );
    assert_eq!(m["ok"], json!(true));
    assert_eq!(m["returned_count"], json!(2));
    assert_eq!(m["items"][0]["id"], json!("MSG_A"));
    assert_eq!(
        m["items"][0]["received_at"],
        json!("2026-10-09T12:00:00"),
        "Eastern wall time"
    );
    assert_eq!(m["items"][0]["is_unread"], json!(true));
    assert_eq!(server.last().q("$top").as_deref(), Some("2"));
    let keys: Vec<&str> = m["items"][0]
        .as_object()
        .unwrap()
        .keys()
        .map(String::as_str)
        .collect();
    assert_eq!(
        keys,
        [
            "account",
            "id",
            "subject",
            "from",
            "is_unread",
            "is_draft",
            "received_datetime",
            "received_at",
            "web_link",
            "conversation_id",
            "folder"
        ]
    );

    // Outlook REST lists one folder, Inbox by default, and flags drafts.
    let mut vars_rest = vars.to_vec();
    vars_rest.push(("OUTLOOK_TEST_TOKEN_FABRIKAM", rest.as_str()));
    let ctx2 = common::ctx(&dir, &vars_rest);
    let r2 = |cmd: &str, b: Value| run(&ctx2, cmd, &body(b), false).unwrap();
    let inbox = r2("messages", json!({ "account": "fabrikam", "limit": 5 }));
    assert_eq!(inbox["ok"], json!(true));
    assert_eq!(inbox["folder"], json!("Inbox"));
    assert_eq!(
        server.last().path,
        "/api/v2.0/me/MailFolders/Inbox/messages"
    );
    assert_eq!(inbox["items"][0]["folder"], json!("Inbox"));
    assert_eq!(inbox["items"][0]["is_draft"], json!(false));
    let sent = r2(
        "messages",
        json!({ "account": "fabrikam", "folder": "sent" }),
    );
    assert_eq!(
        server.last().path,
        "/api/v2.0/me/MailFolders/SentItems/messages"
    );
    assert_eq!(sent["items"][0]["folder"], json!("SentItems"));
    let drafts = r2(
        "messages",
        json!({ "account": "fabrikam", "folder": "Drafts" }),
    );
    assert_eq!(drafts["items"][0]["is_draft"], json!(true));
    let bad = r2(
        "messages",
        json!({ "account": "fabrikam", "folder": "everything" }),
    );
    assert_eq!(bad["ok"], json!(false));
    assert_eq!(bad["status"], json!("invalid_folder"));
    // A sender filter alone still sorts by date: the date clause leads the filter.
    r2(
        "messages",
        json!({ "account": "fabrikam", "from": "dia@fabrikam.com" }),
    );
    assert_eq!(
        server.last().q("$orderby").as_deref(),
        Some("ReceivedDateTime desc")
    );
    assert_eq!(
        server.last().q("$filter").as_deref(),
        Some("ReceivedDateTime ge 1900-01-01T00:00:00Z and From/EmailAddress/Address eq 'dia@fabrikam.com'")
    );
    r2(
        "messages",
        json!({ "account": "fabrikam", "unread": "true", "since": "2026-11-01" }),
    );
    assert_eq!(
        server.last().q("$filter").as_deref(),
        Some("ReceivedDateTime ge 2026-11-01T00:00:00.000Z and IsRead eq false")
    );
    // $search can't be combined with $orderby.
    r2(
        "messages",
        json!({ "account": "fabrikam", "search": "duplicate" }),
    );
    assert_eq!(server.last().q("$search").as_deref(), Some("\"duplicate\""));
    assert_eq!(server.last().q("$orderby"), None);

    // Graph lists the requested folder too, and searches without $orderby.
    r(
        "messages",
        json!({ "account": "contoso", "search": "invoice" }),
    );
    assert_eq!(server.last().path, "/v1.0/me/mailFolders/Inbox/messages");
    assert_eq!(server.last().q("$orderby"), None);
    assert_eq!(
        server.last().header("consistency-level").as_deref(),
        Some("eventual")
    );
    r(
        "search",
        json!({ "account": "contoso", "query": "invoice" }),
    );
    assert_eq!(server.last().q("$orderby"), None);

    // read
    let rd = r(
        "read",
        json!({ "account": "contoso", "id": "READ_ID", "max_body_chars": 2000 }),
    );
    assert_eq!(rd["ok"], json!(true));
    assert!(rd["body_text"].as_str().unwrap().contains("Line 1\nLine 2"));
    assert_eq!(
        rd["answer_text"],
        json!("Read Test from carol@example.com.")
    );
    assert_eq!(
        r("read", json!({ "account": "contoso" }))["status"],
        json!("missing_field")
    );
    let _rd2 = r2(
        "read",
        json!({ "account": "fabrikam", "message_id": "R/2+=" }),
    );
    assert!(
        server
            .requests()
            .iter()
            .any(|q| q.path == "/api/v2.0/me/messages/R%2F2%2B%3D"),
        "ids are URL-encoded"
    );

    // search
    let sr = r(
        "search",
        json!({ "account": "contoso", "query": "search words", "limit": 1 }),
    );
    assert_eq!(sr["ok"], json!(true));
    assert_eq!(sr["returned_count"], json!(1));
    assert_eq!(sr["items"][0]["id"], json!("MSG_S"));
    assert_eq!(r("search", json!({}))["status"], json!("missing_field"));

    // agenda today uses Eastern day bounds (00:00..23:59) with the Prefer header
    let ag = run(
        &common::ctx(
            &dir,
            &[
                vars[0],
                vars[1],
                vars[2],
                ("OUTLOOK_TEST_NOW_MS", "1794805200000"),
            ],
        ),
        "agenda",
        &body(json!({ "account": "contoso", "today": "true" })),
        false,
    )
    .unwrap();
    assert_eq!(ag["ok"], json!(true));
    let acc = &ag["accounts"][0];
    assert_eq!(acc["status"], json!("ok"));
    assert_eq!(acc["items"][0]["subject"], json!("Standup"));
    assert_eq!(
        acc["items"][0]["start"]["dateTime"],
        json!("2026-10-09T09:00:00")
    );
    // 1794805200000 = 2026-11-16T05:00:00Z = 00:00 Eastern
    assert_eq!(ag["start_date"], json!("2026-11-16T00:00:00"));
    assert_eq!(ag["end_date"], json!("2026-11-16T23:59:59"));
    assert_eq!(
        server.last().header("prefer").as_deref(),
        Some("outlook.timezone=\"America/New_York\"")
    );
    assert_eq!(
        server.last().q("startDateTime").as_deref(),
        Some("2026-11-16T00:00:00")
    );
    let ag2 = r2(
        "agenda",
        json!({ "account": "fabrikam", "start_date": "2026-11-17", "end_date": "2026-11-18" }),
    );
    assert_eq!(ag2["start_date"], json!("2026-11-17T00:00:00"));
    assert_eq!(ag2["end_date"], json!("2026-11-18T23:59:59"));
    assert_eq!(server.last().path, "/api/v2.0/me/calendarview");
    assert_eq!(ag2["accounts"][0]["items"][0]["location"], json!("Zoom"));

    // "all" covers both accounts, newest first across them
    let both = r2("messages", json!({}));
    assert_eq!(both["scope"], json!("all"));
    assert_eq!(both["accounts"].as_array().unwrap().len(), 2);
    assert_eq!(
        both["items"][0]["account"],
        json!("fabrikam"),
        "2026-11-11 is newer than 2026-10-09"
    );

    // An unknown account is an error, not a silent fallback.
    assert!(run(
        &ctx,
        "messages",
        &body(json!({ "account": "gmail" })),
        false
    )
    .unwrap_err()
    .to_string()
    .contains("Unknown account 'gmail'. Use contoso, fabrikam, all."));
}

#[test]
fn account_aliases() {
    let dir = two_accounts();
    let ctx = ctx(&dir, &[]);
    let res = |a: Option<Value>| resolve_account_alias(&ctx, a.as_ref());
    for a in [
        "contoso",
        "Contoso",
        "contoso work",
        "me@contoso.com",
        "contoso.com",
        "anyone@contoso.com",
    ] {
        assert_eq!(res(Some(json!(a))).as_deref(), Some("contoso"), "{a}");
    }
    for a in ["fabrikam", "FABRIKAM", "fabrikam.com"] {
        assert_eq!(res(Some(json!(a))).as_deref(), Some("fabrikam"), "{a}");
    }
    for a in ["", "all", "gmail", "someone@gmail.com"] {
        assert_eq!(res(Some(json!(a))), None, "{a}");
    }
    assert_eq!(res(None), None);
}

#[test]
fn aliases_reach_commands() {
    let server = fake();
    let dir = two_accounts();
    let rest = rest_token();
    let ctx = ctx(
        &dir,
        &[
            ("OUTLOOK_REST_BASE_URL", &server.base),
            ("OUTLOOK_TEST_TOKEN", &rest),
        ],
    );
    let via_alias = run(
        &ctx,
        "messages",
        &body(json!({ "account": "fabrikam.com", "limit": 1 })),
        false,
    )
    .unwrap();
    assert_eq!(via_alias["ok"], json!(true));
    let listed = run(
        &ctx,
        "messages",
        &body(json!({ "account": "contoso", "max_results": "1" })),
        false,
    )
    .unwrap();
    assert_eq!(listed["ok"], json!(true));
    assert_eq!(server.last().q("$top").as_deref(), Some("1"));
    let pending = run(
        &ctx,
        "draft-new",
        &body(json!({ "account": "contoso", "to": "a@b.c", "subject": "s", "body": "b" })),
        false,
    )
    .unwrap();
    assert_eq!(pending["status"], json!("confirmation_required"));
}

#[test]
fn needs_sign_in_without_session() {
    let dir = two_accounts();
    let ctx = ctx(&dir, &[]);
    let s = run(&ctx, "status", &body(json!({})), false).unwrap();
    assert_eq!(s["accounts_checked"], json!(2));
    assert_eq!(s["accounts"][0]["status"], json!("needs_sign_in"));
    assert_eq!(
        s["answer_text"],
        json!("0/2 Outlook sessions are signed in.")
    );
    let m = run(&ctx, "messages", &body(json!({})), false).unwrap();
    assert_eq!(m["ok"], json!(false));
    assert_eq!(m["accounts"][0]["status"], json!("needs_sign_in"));
}
