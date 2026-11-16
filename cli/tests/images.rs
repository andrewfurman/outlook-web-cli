//! Port of test/images.test.mjs: attachment listing (inline screenshots, files, items),
//! base64 data, byte ranges joined back together, the in-memory cache, Graph shapes,
//! attachment-inspect with an injected and an OpenRouter-shaped vision model.

mod common;

use base64::Engine;
use common::*;
use outlook_web_cli::commands::run;
use serde_json::{json, Value};
use std::sync::{Arc, Mutex};

const PNG_1X1: &str = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

fn b64std() -> base64::engine::GeneralPurpose {
    base64::engine::general_purpose::STANDARD
}

/// image001 is padded past one 1,000-byte range so the byte-range test needs several pieces.
fn shot() -> String {
    let mut bytes = b64std().decode(PNG_1X1).unwrap();
    bytes.extend(std::iter::repeat_n(7u8, 2500));
    b64std().encode(bytes)
}

fn fake() -> Fake {
    let shot = shot();
    Fake::start(move |req| {
        let a0 = json!({ "@odata.type": "#Microsoft.OutlookServices.FileAttachment", "Id": "A0", "Name": "image001.png", "ContentType": "image/png", "Size": 2568, "IsInline": true, "ContentId": "image001.png@01DD0000.00000000", "ContentBytes": shot });
        match (req.method.as_str(), req.path.as_str()) {
            ("GET", "/api/v2.0/me/messages/MSG_SHOT/attachments") => Reply::json(
                200,
                json!({ "value": [
                a0,
                { "@odata.type": "#Microsoft.OutlookServices.FileAttachment", "Id": "A1", "Name": "../../logo.png", "ContentType": "image/png", "Size": 68, "IsInline": true, "ContentId": "image002.png@01DD0000.00000000", "ContentBytes": PNG_1X1 },
                { "@odata.type": "#Microsoft.OutlookServices.FileAttachment", "Id": "A2", "Name": "Statement.pdf", "ContentType": "application/pdf", "Size": 9, "IsInline": false, "ContentBytes": b64std().encode("%PDF-1.4\n") },
                { "@odata.type": "#Microsoft.OutlookServices.ItemAttachment", "Id": "A3", "Name": "Fwd: earlier", "ContentType": null, "Size": 2000, "IsInline": false },
            ] }),
            ),
            ("GET", "/api/v2.0/me/messages/MSG_SHOT/attachments/A0") => Reply::json(200, a0),
            ("GET", "/v1.0/me/messages/MSG_GRAPH/attachments") => Reply::json(
                200,
                json!({ "value": [
                { "@odata.type": "#microsoft.graph.fileAttachment", "id": "G0", "name": "chart.jpg", "contentType": "image/jpeg", "size": 68, "isInline": false, "contentBytes": PNG_1X1 },
            ] }),
            ),
            ("POST", "/chat/completions") => {
                assert!(req
                    .header("authorization")
                    .unwrap()
                    .starts_with("Bearer sk-test"));
                let b = req.json();
                assert_eq!(b["messages"][0]["content"][1]["type"], json!("image_url"));
                assert!(b["messages"][0]["content"][1]["image_url"]["url"]
                    .as_str()
                    .unwrap()
                    .starts_with("data:image/png;base64,"));
                Reply::json(
                    200,
                    json!({ "choices": [{ "message": { "content": "```json\n{\"description\": \"A logo.\", \"ocr_text\": \"ACME\"}\n```" } }] }),
                )
            }
            _ => Reply::json(404, json!({ "error": { "message": "not found" } })),
        }
    })
}

#[test]
fn attachments_and_inspect() {
    let server = fake();
    let dir = two_accounts();
    let (rest, graph) = (rest_token(), graph_token());
    let vars = [
        ("OUTLOOK_TEST_TOKEN_FABRIKAM", rest.as_str()),
        ("OUTLOOK_TEST_TOKEN_CONTOSO", graph.as_str()),
        ("OUTLOOK_REST_BASE_URL", server.base.as_str()),
        ("OUTLOOK_GRAPH_BASE_URL", server.base.as_str()),
    ];
    let calls: Arc<Mutex<Vec<Value>>> = Arc::default();
    let seen = calls.clone();
    let ctx = ctx(&dir, &vars).with_vision(Box::new(move |args| {
        seen.lock().unwrap().push(args.clone());
        Ok(json!({ "ok": true, "provider": "openrouter", "model": "google/gemini-3.8-flash", "description": "A reports page.", "ocr_text": "Reports\nSeptember 2026", "answer_text": "A reports page." }))
    }));
    let r = |cmd: &str, b: Value| run(&ctx, cmd, &body(b), false).unwrap();
    let shot = shot();

    // 1. List: inline screenshots show up with content ids; no base64 unless asked.
    let listed = r(
        "attachments",
        json!({ "account": "fabrikam", "id": "MSG_SHOT" }),
    );
    assert_eq!(listed["ok"], json!(true));
    assert_eq!(listed["returned_count"], json!(4));
    assert_eq!(listed["image_count"], json!(2));
    let a = listed["attachments"].as_array().unwrap();
    assert_eq!(a[0]["name"], json!("image001.png"));
    assert_eq!(a[0]["is_inline"], json!(true));
    assert_eq!(a[0]["content_id"], json!("image001.png@01DD0000.00000000"));
    assert_eq!(a[0]["is_image"], json!(true));
    assert_eq!(a[2]["is_image"], json!(false));
    assert_eq!(a[3]["kind"], json!("item"));
    assert!(a[0].get("data_base64").is_none());
    let keys: Vec<&str> = a[0]
        .as_object()
        .unwrap()
        .keys()
        .map(String::as_str)
        .collect();
    assert_eq!(
        keys,
        [
            "index",
            "attachment_id",
            "name",
            "content_type",
            "size",
            "is_inline",
            "content_id",
            "kind",
            "is_image"
        ]
    );
    assert!(listed["answer_text"]
        .as_str()
        .unwrap()
        .contains("4 attachments, 2 images"));

    // 2. include_data + images_only (as the CLI passes them: strings).
    let with_data = r(
        "attachments",
        json!({ "account": "fabrikam", "id": "MSG_SHOT", "include_data": "true", "images_only": "true" }),
    );
    assert_eq!(with_data["returned_count"], json!(2));
    assert_eq!(with_data["attachments"][0]["data_base64"], json!(shot));

    // 3. One attachment, whole or as byte ranges joined back together; nothing on disk.
    let png = b64std().decode(&shot).unwrap();
    let one = r(
        "attachments",
        json!({ "account": "fabrikam", "id": "MSG_SHOT", "content_id": "<image001.png@01DD0000.00000000>" }),
    );
    assert_eq!(one["attachment"]["name"], json!("image001.png"));
    assert_eq!(one["total_bytes"], json!(png.len()));
    assert_eq!(one["data_base64"], json!(shot));
    let first_image = r(
        "attachments",
        json!({ "account": "fabrikam", "id": "MSG_SHOT", "index": 1, "include_data": false }),
    );
    assert_eq!(first_image["attachment"]["name"], json!("../../logo.png"));
    assert!(first_image.get("data_base64").is_none());
    assert_eq!(
        r(
            "attachments",
            json!({ "account": "fabrikam", "id": "MSG_SHOT", "index": "9" })
        )["status"],
        json!("attachment_not_found")
    );
    let mut pieces = Vec::new();
    let mut next = json!(0);
    let mut piece = 0;
    while !next.is_null() {
        let mut b =
            json!({ "account": "fabrikam", "id": "MSG_SHOT", "offset": next, "length": "1000" });
        if piece > 0 {
            b["attachment_id"] = json!("A0");
        } else {
            b["index"] = json!(0);
        }
        let res = r("attachments", b);
        assert_eq!(res["ok"], json!(true), "{res}");
        pieces.extend(
            b64std()
                .decode(res["data_base64"].as_str().unwrap())
                .unwrap(),
        );
        next = res["next_offset"].clone();
        piece += 1;
    }
    assert_eq!(pieces, png);
    assert!(piece > 2);
    assert!(
        server
            .requests()
            .iter()
            .any(|q| q.path.ends_with("/attachments/A0")),
        "later pieces fetch only that attachment"
    );

    // 3b. Attachments stay in memory: repeat calls in one process don't hit Outlook again.
    let before = server.count();
    r(
        "attachments",
        json!({ "account": "fabrikam", "id": "MSG_SHOT", "index": 0, "offset": 0, "length": 1000 }),
    );
    r(
        "attachments",
        json!({ "account": "fabrikam", "id": "MSG_SHOT", "index": 0, "offset": 1000, "length": 1000 }),
    );
    assert_eq!(server.count(), before);

    // 4. Graph token shape (camelCase) works too.
    let g = r(
        "attachments",
        json!({ "account": "contoso", "id": "MSG_GRAPH" }),
    );
    assert_eq!(g["ok"], json!(true));
    assert_eq!(g["attachments"][0]["content_type"], json!("image/jpeg"));
    assert_eq!(g["attachments"][0]["is_image"], json!(true));

    // 5. Inspect: first image by default, by content id, by index.
    let first = r(
        "attachment-inspect",
        json!({ "account": "fabrikam", "id": "MSG_SHOT" }),
    );
    assert_eq!(first["ok"], json!(true));
    assert_eq!(first["image"]["name"], json!("image001.png"));
    assert_eq!(first["ocr_text"], json!("Reports\nSeptember 2026"));
    {
        let c = calls.lock().unwrap();
        assert_eq!(c[0]["provider"], json!("openrouter"));
        assert!(c[0]["prompt"]
            .as_str()
            .unwrap()
            .contains("transcribe ALL visible text"));
        assert_eq!(c[0]["images"][0]["data_base64"], json!(shot));
    }
    let by_cid = r(
        "attachment-inspect",
        json!({ "account": "fabrikam", "id": "MSG_SHOT", "cid": "<image002.png@01DD0000.00000000>", "prompt": "Read it" }),
    );
    assert_eq!(by_cid["image"]["index"], json!(1));
    assert_eq!(calls.lock().unwrap()[1]["prompt"], json!("Read it"));
    let not_image = r(
        "attachment-inspect",
        json!({ "account": "fabrikam", "id": "MSG_SHOT", "index": 2 }),
    );
    assert_eq!(not_image["ok"], json!(false));
    assert_eq!(not_image["status"], json!("image_not_found"));
    assert_eq!(calls.lock().unwrap().len(), 2);

    // 6. Reads only.
    assert!(server.requests().iter().all(|q| q.method == "GET"));

    // 7. Without an image model, inspect explains what to set up.
    let plain = common::ctx(&dir, &vars);
    let none = run(
        &plain,
        "attachment-inspect",
        &body(json!({ "account": "fabrikam", "id": "MSG_SHOT" })),
        false,
    )
    .unwrap();
    assert_eq!(none["status"], json!("vision_not_configured"));
    assert!(none["message"]
        .as_str()
        .unwrap()
        .contains("OPENROUTER_API_KEY"));

    // 8. With OPENROUTER_API_KEY the OpenRouter analyzer is used (fake base URL).
    let mut v = vars.to_vec();
    v.push(("OPENROUTER_API_KEY", "sk-test"));
    v.push(("OPENROUTER_BASE_URL", server.base.as_str()));
    let mut with_key = common::ctx(&dir, &v);
    with_key.vision = outlook_web_cli::vision::openrouter_analyzer(&with_key);
    let o = run(
        &with_key,
        "attachment-inspect",
        &body(json!({ "account": "fabrikam", "id": "MSG_SHOT", "index": 1 })),
        false,
    )
    .unwrap();
    assert_eq!(o["ok"], json!(true), "{o}");
    assert_eq!(o["provider"], json!("openrouter"));
    assert_eq!(o["model"], json!("google/gemini-2.5-flash"));
    assert_eq!(o["description"], json!("A logo."));
    assert_eq!(o["ocr_text"], json!("ACME"));
}
