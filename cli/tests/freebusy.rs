//! Port of test/freebusy.test.mjs: getSchedule with the Eastern Prefer header, account
//! picked from the email domain, merged busy blocks, common free slots, weekends,
//! minimum length, working day, and input errors.

mod common;

use common::*;
use outlook_web_cli::commands::run;
use serde_json::{json, Value};

fn fake() -> Fake {
    Fake::start(|req| {
        if req.method == "POST" && req.path == "/api/v2.0/me/calendar/getschedule" {
            let payload = req.json();
            let item = |s: &str, e: &str, status: &str, extra: Value| {
                let mut v = json!({ "Status": status, "Start": { "DateTime": format!("{s}:00.0000000"), "TimeZone": "Eastern Standard Time" }, "End": { "DateTime": format!("{e}:00.0000000"), "TimeZone": "Eastern Standard Time" } });
                for (k, x) in extra.as_object().unwrap() {
                    v[k] = x.clone();
                }
                v
            };
            let data = |id: &str| -> Value {
                match id {
                    "alice@contoso.com" => json!({ "ScheduleItems": [
                        item("2026-11-17T10:00", "2026-11-17T10:30", "Busy", json!({})),
                        item("2026-11-17T10:15", "2026-11-17T11:00", "Tentative", json!({})),
                        item("2026-11-17T13:00", "2026-11-17T13:30", "Free", json!({})),
                        item("2026-11-17T19:00", "2026-11-17T20:00", "Busy", json!({})),
                    ] }),
                    "bob@contoso.com" => json!({ "ScheduleItems": [
                        item("2026-11-17T11:30", "2026-11-17T12:00", "Busy", json!({ "Subject": "Shared calendar standup", "Location": "Room 2" })),
                        item("2026-11-17T14:00", "2026-11-17T16:45", "Oof", json!({})),
                    ] }),
                    "nobody@elsewhere.com" => {
                        json!({ "Error": { "Message": "The user wasn't found." } })
                    }
                    _ => json!({ "ScheduleItems": [] }),
                }
            };
            let value: Vec<Value> = payload["Schedules"]
                .as_array()
                .unwrap()
                .iter()
                .map(|id| {
                    let mut v = data(id.as_str().unwrap());
                    v["ScheduleId"] = id.clone();
                    v
                })
                .collect();
            return Reply::json(200, json!({ "value": value }));
        }
        Reply::json(404, json!({}))
    })
}

#[test]
fn freebusy() {
    let server = fake();
    let dir = two_accounts();
    let token = jwt(json!({ "aud": "https://outlook.office.com", "scp": "Calendars.ReadWrite" }));
    let ctx = ctx(
        &dir,
        &[
            ("OUTLOOK_REST_BASE_URL", &server.base),
            ("OUTLOOK_TEST_TOKEN_CONTOSO", &token),
        ],
    );
    let fb = |b: Value| run(&ctx, "freebusy", &body(b), false).unwrap();
    let slots = |v: &Value| -> Vec<(String, String, i64)> {
        v["free_for_all"]
            .as_array()
            .unwrap()
            .iter()
            .map(|s| {
                (
                    s["start"].as_str().unwrap()[11..].to_string(),
                    s["end"].as_str().unwrap()[11..].to_string(),
                    s["minutes"].as_i64().unwrap(),
                )
            })
            .collect()
    };

    // 1. Two people for one day: the account is picked from their email domain.
    let r = fb(
        json!({ "emails": "Alice <alice@contoso.com>, bob@contoso.com; ALICE@contoso.com", "start_date": "2026-11-17", "end_date": "2026-11-17" }),
    );
    assert_eq!(r["ok"], json!(true), "{r}");
    assert_eq!(r["account"], json!("contoso"));
    let post = server.last();
    let payload = post.json();
    assert_eq!(
        payload["Schedules"],
        json!(["alice@contoso.com", "bob@contoso.com"]),
        "deduplicated, lower-cased"
    );
    assert_eq!(
        payload["StartTime"]["DateTime"],
        json!("2026-11-17T00:00:00")
    );
    assert_eq!(
        payload["StartTime"]["TimeZone"],
        json!("Eastern Standard Time")
    );
    assert_eq!(
        post.header("prefer").as_deref(),
        Some("outlook.timezone=\"America/New_York\""),
        "answers in Eastern, not UTC"
    );
    let person = |email: &str| {
        r["people"]
            .as_array()
            .unwrap()
            .iter()
            .find(|p| p["email"] == json!(email))
            .unwrap()
            .clone()
    };
    let alice = person("alice@contoso.com");
    let busy: Vec<(String, String, String)> = alice["busy"]
        .as_array()
        .unwrap()
        .iter()
        .map(|b| {
            (
                b["start"].as_str().unwrap()[11..].into(),
                b["end"].as_str().unwrap()[11..].into(),
                b["status"].as_str().unwrap().into(),
            )
        })
        .collect();
    assert_eq!(
        busy,
        vec![
            ("10:00".into(), "11:00".into(), "busy".into()),
            ("19:00".into(), "20:00".into(), "busy".into())
        ],
        "overlaps merged, Free items dropped"
    );
    let bob = person("bob@contoso.com");
    assert_eq!(bob["busy"][0]["subject"], json!("Shared calendar standup"));
    assert_eq!(bob["busy"][0]["location"], json!("Room 2"));
    assert_eq!(bob["busy"][1]["status"], json!("oof"));
    assert_eq!(
        slots(&r)
            .into_iter()
            .map(|(a, b, _)| (a, b))
            .collect::<Vec<_>>(),
        vec![
            ("09:00".into(), "10:00".into()),
            ("11:00".into(), "11:30".into()),
            ("12:00".into(), "14:00".into())
        ],
        "16:45-17:00 is shorter than 30 minutes"
    );
    assert!(r["answer_text"]
        .as_str()
        .unwrap()
        .contains("3 slots when everyone is free"));
    assert_eq!(r["working_day"], json!("09:00-17:00"));
    assert_eq!(r["timezone"], json!("America/New_York"));

    // 2. A date-time window, a shorter minimum and a custom working day (CLI-style strings).
    let w = fb(
        json!({ "emails": ["alice@contoso.com", "bob@contoso.com"], "start": "2026-11-17T11:00", "end": "2026-11-17T18:00", "min_minutes": "15", "day_end": "18:00" }),
    );
    assert_eq!(
        server.last().json()["StartTime"]["DateTime"],
        json!("2026-11-17T11:00:00"),
        "seconds added"
    );
    assert_eq!(
        slots(&w),
        vec![
            ("11:00".into(), "11:30".into(), 30),
            ("12:00".into(), "14:00".into(), 120),
            ("16:45".into(), "18:00".into(), 75)
        ]
    );

    // 3. Someone without free/busy is reported, the rest still answer.
    let x = fb(
        json!({ "account": "contoso", "emails": "alice@contoso.com,nobody@elsewhere.com", "tomorrow": true }),
    );
    assert_eq!(x["ok"], json!(true));
    let nobody = x["people"]
        .as_array()
        .unwrap()
        .iter()
        .find(|p| p["email"] == json!("nobody@elsewhere.com"))
        .unwrap()
        .clone();
    assert_eq!(nobody["ok"], json!(false));
    assert_eq!(nobody["error"], json!("The user wasn't found."));
    assert!(x["answer_text"]
        .as_str()
        .unwrap()
        .contains("No free/busy for nobody@elsewhere.com"));

    // 4. Weekends are skipped unless asked for.
    let sat = fb(
        json!({ "emails": "alice@contoso.com", "start_date": "2026-11-21", "end_date": "2026-11-21" }),
    );
    assert_eq!(sat["free_for_all"].as_array().unwrap().len(), 0);
    let sat_on = fb(
        json!({ "emails": "alice@contoso.com", "start_date": "2026-11-21", "end_date": "2026-11-21", "weekends": "true" }),
    );
    assert_eq!(sat_on["free_for_all"].as_array().unwrap().len(), 1);
    assert!(sat_on["answer_text"]
        .as_str()
        .unwrap()
        .contains("when they are free"));

    // 5. Input errors.
    assert_eq!(
        fb(json!({ "emails": "" }))["status"],
        json!("missing_emails")
    );
    assert_eq!(
        fb(json!({ "emails": "someone@gmail.com" }))["status"],
        json!("account_required"),
        "two accounts and no domain match"
    );
    let many: Vec<String> = (0..21).map(|i| format!("p{i}@contoso.com")).collect();
    assert_eq!(
        fb(json!({ "emails": many }))["status"],
        json!("too_many_emails")
    );
}

#[test]
fn single_account_is_the_default() {
    let server = fake();
    let dir = config_dir(json!([{ "name": "work", "username": "me@fabrikam.com" }]));
    let token = rest_token();
    let ctx = ctx(
        &dir,
        &[
            ("OUTLOOK_REST_BASE_URL", &server.base),
            ("OUTLOOK_TEST_TOKEN", &token),
        ],
    );
    let r = run(
        &ctx,
        "freebusy",
        &body(json!({ "emails": "someone@gmail.com", "today": "true" })),
        false,
    )
    .unwrap();
    assert_eq!(r["account"], json!("work"));
    assert_eq!(r["people"][0]["busy"], json!([]));
}
