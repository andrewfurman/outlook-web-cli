//! Port of test/drafts.test.mjs: confirmation, reply above the quoted original,
//! reply-all, new drafts, delete only drafts, inline images, extra recipients, and
//! proof that nothing ever reaches a send endpoint.

mod common;

use base64::Engine;
use common::*;
use outlook_web_cli::commands::run;
use outlook_web_cli::drafts::rest_write;
use serde_json::{json, Value};

const QUOTED: &str = "<html><head></head><body><hr><div><b>From:</b> Halie Reyes<br>Vendors are correct.</div></body></html>";

fn draft(extra: Value) -> Value {
    let mut d = json!({
        "Id": "DRAFT_1",
        "Subject": "RE: Claims data fix",
        "IsDraft": true,
        "ToRecipients": [{ "EmailAddress": { "Name": "Halie Reyes", "Address": "halie@example.com" } }],
        "CcRecipients": [{ "EmailAddress": { "Name": "Oscar", "Address": "oscar@example.com" } }],
        "WebLink": "https://outlook.office.com/owa/?ItemID=DRAFT_1",
        "Body": { "ContentType": "HTML", "Content": QUOTED },
    });
    for (k, v) in extra.as_object().unwrap() {
        d[k] = v.clone();
    }
    d
}

/// A fake Outlook REST API that records every request.
fn fake() -> Fake {
    Fake::start(|req| {
        let (m, p) = (req.method.as_str(), req.path.as_str());
        let re = regex::Regex::new(r"^/api/v2\.0/me/messages/[^/]+/createreply(all)?$").unwrap();
        if m == "POST" && re.is_match(p) {
            return Reply::json(201, draft(json!({})));
        }
        if m == "PATCH" && p == "/api/v2.0/me/messages/DRAFT_1" {
            return Reply::json(200, draft(json!({ "Body": req.json()["Body"] })));
        }
        if m == "GET"
            && (p == "/api/v2.0/me/messages/DRAFT_OLD" || p == "/api/v2.0/me/messages/MSG_SENT")
        {
            let is_draft = p.ends_with("DRAFT_OLD");
            return Reply::json(
                200,
                json!({ "Id": if is_draft { "DRAFT_OLD" } else { "MSG_SENT" }, "Subject": "RE: Status", "IsDraft": is_draft }),
            );
        }
        if m == "POST" && p == "/api/v2.0/me/messages/DRAFT_OLD/move" {
            return Reply::json(201, json!({ "Id": "DRAFT_OLD_MOVED" }));
        }
        if m == "GET" && p == "/img/shot.png" {
            return Reply {
                status: 200,
                headers: vec![("content-type".into(), "image/png".into())],
                body: b"PNGDATA".to_vec(),
            };
        }
        if m == "POST"
            && (p == "/api/v2.0/me/messages/NEW_1/attachments"
                || p == "/api/v2.0/me/messages/DRAFT_1/attachments")
        {
            return Reply::json(201, json!({ "Id": "ATT_1" }));
        }
        if m == "POST" && p == "/api/v2.0/me/messages" {
            let b = req.json();
            return Reply::json(
                201,
                json!({ "Id": "NEW_1", "Subject": b["Subject"], "IsDraft": true, "ToRecipients": b["ToRecipients"], "CcRecipients": b["CcRecipients"], "WebLink": "" }),
            );
        }
        Reply::json(404, json!({ "error": { "message": "not found" } }))
    })
}

#[test]
fn drafts() {
    let server = fake();
    let dir = two_accounts();
    let token = jwt(json!({ "aud": "https://outlook.office.com", "scp": "Mail.ReadWrite" }));
    let ctx = ctx(
        &dir,
        &[
            ("OUTLOOK_TEST_TOKEN_FABRIKAM", &token),
            ("OUTLOOK_REST_BASE_URL", &server.base),
            ("OUTLOOK_INLINE_IMAGE_HOSTS", "127.0.0.1"),
        ],
    );
    let r = |cmd: &str, b: Value, yes: bool| run(&ctx, cmd, &body(b), yes).unwrap();

    // 1. Nothing happens until it's confirmed.
    let pending = r(
        "draft-reply",
        json!({ "account": "fabrikam", "id": "MSG_1", "body": "hi" }),
        false,
    );
    assert_eq!(pending["status"], json!("confirmation_required"));
    assert_eq!(server.count(), 0);

    // 2. A work account must be named: no "all".
    assert_eq!(
        r(
            "draft-reply",
            json!({ "account": "all", "id": "MSG_1", "body": "hi" }),
            true
        )["ok"],
        json!(false)
    );
    assert_eq!(
        r("draft-reply", json!({ "id": "MSG_1", "body": "hi" }), true)["status"],
        json!("missing_field")
    );
    assert_eq!(server.count(), 0);

    // 3. Reply: createreply, then the text goes above the quoted original (escaped).
    let reply = r(
        "draft-reply",
        json!({ "account": "Fabrikam", "id": "MSG_1", "body": "Thanks, Halie.\n\nAll 79 are linked <now>." }),
        true,
    );
    assert_eq!(reply["ok"], json!(true));
    assert_eq!(reply["status"], json!("draft_saved"));
    assert_eq!(reply["id"], json!("DRAFT_1"));
    assert_eq!(reply["is_draft"], json!(true));
    assert!(reply["answer_text"]
        .as_str()
        .unwrap()
        .contains("has not been sent"));
    let reqs = server.requests();
    assert_eq!(reqs[0].method, "POST");
    assert!(reqs[0].path.ends_with("/me/messages/MSG_1/createreply"));
    assert_eq!(reqs[1].method, "PATCH");
    let patched = reqs[1].json()["Body"]["Content"]
        .as_str()
        .unwrap()
        .to_string();
    assert!(
        patched.starts_with("<html><head></head><body><div><p>Thanks, Halie.</p><p>All 79 are linked &lt;now&gt;.</p></div><br><hr>"),
        "{patched}"
    );
    assert!(
        patched.contains("Vendors are correct."),
        "quoted original is kept"
    );

    // 4. Reply all uses createreplyall; HTML bodies go in as-is.
    server.clear();
    let all = r(
        "draft-reply",
        json!({ "account": "fabrikam", "id": "MSG_2", "reply_all": true, "body_html": "<p><b>Done.</b></p>" }),
        true,
    );
    assert_eq!(all["ok"], json!(true));
    let reqs = server.requests();
    assert!(reqs[0].path.ends_with("/me/messages/MSG_2/createreplyall"));
    assert!(reqs[1].json()["Body"]["Content"]
        .as_str()
        .unwrap()
        .contains("<div><p><b>Done.</b></p></div><br>"));

    // 6. New message: saved to Drafts with parsed recipients.
    server.clear();
    let fresh = r(
        "draft-new",
        json!({ "account": "fabrikam", "to": "Oscar Buitrago <oscar@example.com>, dia@example.com", "cc": ["halie@example.com"], "subject": "Status", "body": "Hi" }),
        true,
    );
    assert_eq!(fresh["status"], json!("draft_saved"));
    let req = &server.requests()[0];
    assert_eq!(req.method, "POST");
    assert_eq!(req.path, "/api/v2.0/me/messages");
    assert_eq!(
        req.json()["ToRecipients"],
        json!([
            { "EmailAddress": { "Address": "oscar@example.com", "Name": "Oscar Buitrago" } },
            { "EmailAddress": { "Address": "dia@example.com" } },
        ])
    );
    assert_eq!(
        req.json()["CcRecipients"],
        json!([{ "EmailAddress": { "Address": "halie@example.com" } }])
    );
    assert_eq!(
        fresh["to"][0],
        json!({ "name": "Oscar Buitrago", "address": "oscar@example.com" })
    );

    // 7. Deleting a draft moves it to Deleted Items, only after confirmation, and only drafts.
    server.clear();
    assert_eq!(
        r(
            "draft-delete",
            json!({ "account": "fabrikam", "id": "DRAFT_OLD" }),
            false
        )["status"],
        json!("confirmation_required")
    );
    assert_eq!(server.count(), 0);
    let removed = r(
        "draft-delete",
        json!({ "account": "fabrikam", "id": "DRAFT_OLD" }),
        true,
    );
    assert_eq!(removed["status"], json!("draft_deleted"));
    assert_eq!(removed["subject"], json!("RE: Status"));
    assert_eq!(server.last().path, "/api/v2.0/me/messages/DRAFT_OLD/move");
    assert_eq!(
        server.last().json(),
        json!({ "DestinationId": "DeletedItems" })
    );
    server.clear();
    let refused = r(
        "draft-delete",
        json!({ "account": "fabrikam", "id": "MSG_SENT" }),
        true,
    );
    assert_eq!(refused["status"], json!("not_a_draft"));
    assert!(
        server.requests().iter().all(|q| q.method == "GET"),
        "a non-draft is never moved"
    );

    // 8. Inline images are fetched from the allowed host and attached with IsInline + ContentId.
    server.clear();
    let img_url = format!("{}/img/shot.png", server.base);
    let with_image = r(
        "draft-new",
        json!({ "account": "fabrikam", "to": "oscar@example.com", "subject": "Shots", "body_html": "<img src=\"cid:shot1\">", "inline_images": [{ "url": img_url, "content_id": "shot1" }] }),
        true,
    );
    assert_eq!(with_image["status"], json!("draft_saved"), "{with_image}");
    assert_eq!(with_image["answer_text"], json!("Draft saved with 1 inline image(s). It's in fabrikam Outlook Drafts and has not been sent."));
    let att = server
        .requests()
        .into_iter()
        .find(|q| q.path.ends_with("/attachments"))
        .unwrap()
        .json();
    assert_eq!(att["IsInline"], json!(true));
    assert_eq!(att["ContentId"], json!("shot1"));
    assert_eq!(att["ContentType"], json!("image/png"));
    assert_eq!(att["Name"], json!("shot.png"));
    let bytes = base64::engine::general_purpose::STANDARD
        .decode(att["ContentBytes"].as_str().unwrap())
        .unwrap();
    assert_eq!(bytes, b"PNGDATA");
    // A host that isn't allowed is refused before anything is fetched.
    let refused_img = r(
        "draft-new",
        json!({ "account": "fabrikam", "to": "oscar@example.com", "subject": "x", "body_html": "x", "inline_images": [{ "url": "https://evil.example/x.png", "content_id": "a" }] }),
        true,
    );
    assert_eq!(refused_img["status"], json!("inline_image_failed"));
    assert_eq!(refused_img["ok"], json!(false));

    // 9. Reply-all drafts take inline images too, attached to the new draft.
    server.clear();
    let reply_img = r(
        "draft-reply",
        json!({ "account": "fabrikam", "id": "MSG_4", "reply_all": true, "body_html": "<img src=\"cid:r1\">", "inline_images": [{ "url": img_url, "content_id": "r1" }] }),
        true,
    );
    assert_eq!(reply_img["status"], json!("draft_saved"), "{reply_img}");
    let ra = server
        .requests()
        .into_iter()
        .find(|q| q.path == "/api/v2.0/me/messages/DRAFT_1/attachments")
        .unwrap()
        .json();
    assert_eq!(ra["ContentId"], json!("r1"));
    assert_eq!(ra["IsInline"], json!(true));

    // 10. Reply-all can add recipients; anyone already on the draft stays once.
    server.clear();
    let widened = r(
        "draft-reply",
        json!({ "account": "fabrikam", "id": "MSG_5", "reply_all": true, "body": "ok", "add_to": "oscar@example.com, rebecca@example.com", "add_cc": ["Halie@example.com", "phil@example.com"] }),
        true,
    );
    assert_eq!(widened["status"], json!("draft_saved"), "{widened}");
    let patch = server
        .requests()
        .into_iter()
        .find(|q| q.method == "PATCH")
        .unwrap()
        .json();
    let addrs = |k: &str| {
        patch[k]
            .as_array()
            .unwrap()
            .iter()
            .map(|r| r["EmailAddress"]["Address"].as_str().unwrap().to_string())
            .collect::<Vec<_>>()
    };
    assert_eq!(
        addrs("ToRecipients"),
        [
            "halie@example.com",
            "oscar@example.com",
            "rebecca@example.com"
        ]
    );
    assert_eq!(addrs("CcRecipients"), ["phil@example.com"]);
    // Without add_to/add_cc the recipients aren't touched.
    server.clear();
    r(
        "draft-reply",
        json!({ "account": "fabrikam", "id": "MSG_6", "reply_all": true, "body": "ok" }),
        true,
    );
    assert!(server
        .requests()
        .into_iter()
        .find(|q| q.method == "PATCH")
        .unwrap()
        .json()
        .get("ToRecipients")
        .is_none());

    // 11. The only write path refuses /send, and nothing ever went to a send endpoint.
    let err = rest_write(&ctx, &token, "POST", "me/messages/DRAFT_1/send", None).unwrap_err();
    assert!(err.to_string().contains("Refusing to send mail"));
    assert!(!server
        .requests()
        .iter()
        .any(|q| q.path.to_lowercase().contains("/send")));
}

#[test]
fn inline_images_off_without_allowed_hosts() {
    let server = fake();
    let dir = two_accounts();
    let token = rest_token();
    let ctx = ctx(
        &dir,
        &[
            ("OUTLOOK_TEST_TOKEN", &token),
            ("OUTLOOK_REST_BASE_URL", &server.base),
        ],
    );
    let img_url = format!("{}/img/shot.png", server.base);
    let res = run(&ctx, "draft-new", &body(json!({ "account": "contoso", "to": "a@example.com", "body": "x", "inline_images": [{ "url": img_url, "content_id": "c" }] })), true).unwrap();
    assert_eq!(res["status"], json!("inline_image_failed"));
    assert!(res["error"]
        .as_str()
        .unwrap()
        .contains("set OUTLOOK_INLINE_IMAGE_HOSTS"));
    assert!(
        !server.requests().iter().any(|q| q.path == "/img/shot.png"),
        "image never fetched"
    );
}
