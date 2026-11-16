//! Free/busy for one or more people, like Outlook's Scheduling Assistant. Uses the
//! signed-in account's getSchedule, so it works for anyone in the same organization
//! (and other organizations that publish free/busy) without them sharing their calendar.
//! Returns each person's busy blocks (subject and location only when they share more
//! than free/busy) and the slots inside the working day when everyone is free.
//! Read-only. All times are US Eastern wall-clock times.

use crate::accounts::{
    get_account, list_accounts, resolve_account_alias, resolve_alias_str, Account,
};
use crate::ctx::Ctx;
use crate::drafts::parse_recipients;
use crate::js::{self, normalize_string, plural};
use crate::outlook::{
    account_names, api_url, normalize_et_local_range, rest_base, safe_get_best_token, send,
    weekday_sun0, RangeArgs, DEFAULT_TIMEOUT, ET_TIMEZONE,
};
use anyhow::Result;
use chrono::{NaiveDateTime, TimeZone, Utc};
use serde_json::{json, Map, Value};

const FREEBUSY_MAX_PEOPLE: usize = 20;
const WINDOWS_TZ: &str = "Eastern Standard Time";

pub struct FreeBusyArgs<'a> {
    pub account: Option<&'a Value>,
    pub emails: Option<&'a Value>,
    pub range: RangeArgs<'a>,
    pub day_start: Option<&'a Value>,
    pub day_end: Option<&'a Value>,
    pub min_minutes: Option<&'a Value>,
    pub weekends: Option<&'a Value>,
}

/// The account to query from: the one named, else the one whose domain matches the
/// people being checked, else the only account.
fn free_busy_account(ctx: &Ctx, account: Option<&Value>, people: &[String]) -> Option<Account> {
    let named = normalize_string(account);
    if !named.is_empty() && named.to_lowercase() != "all" {
        return resolve_account_alias(ctx, account).and_then(|k| get_account(ctx, &k));
    }
    let all = list_accounts(ctx);
    for email in people {
        let domain = email.rsplit('@').next().unwrap_or(email);
        if let Some(key) = resolve_alias_str(ctx, domain) {
            return get_account(ctx, &key);
        }
    }
    if all.len() == 1 {
        all.into_iter().next()
    } else {
        None
    }
}

/// "2026-11-17T11:00:00.0000000" -> "2026-11-17T11:00"
fn wall_time(v: Option<&Value>) -> String {
    js::str_or_empty(v).chars().take(16).collect()
}

/// Wall-clock "YYYY-MM-DDTHH:MM" -> minutes since the epoch, with no time zone math.
fn wall_minutes(v: &str) -> Option<i64> {
    let head: String = v.chars().take(16).collect();
    NaiveDateTime::parse_from_str(&head, "%Y-%m-%dT%H:%M")
        .ok()
        .map(|d| d.and_utc().timestamp() / 60)
}

fn minutes_to_wall(m: i64) -> String {
    Utc.timestamp_opt(m * 60, 0)
        .single()
        .map(|d| d.format("%Y-%m-%dT%H:%M").to_string())
        .unwrap_or_default()
}

/// Joins overlapping or touching busy blocks for one person.
fn merge_busy(blocks: Vec<Map<String, Value>>) -> Vec<Value> {
    let mut out: Vec<Map<String, Value>> = Vec::new();
    for b in blocks {
        let s = |m: &Map<String, Value>, k: &str| js::str_or_empty(m.get(k));
        if let Some(last) = out.last_mut() {
            if s(&b, "start") <= s(last, "end") {
                if s(&b, "end") > s(last, "end") {
                    last.insert("end".into(), json!(s(&b, "end")));
                }
                let (ls, bs) = (s(last, "status"), s(&b, "status"));
                if ls != bs && (ls == "busy" || bs == "busy") {
                    last.insert("status".into(), json!("busy"));
                }
                if b.get("subject").is_some_and(js::truthy)
                    && !last.get("subject").is_some_and(js::truthy)
                {
                    last.insert("subject".into(), b["subject"].clone());
                }
                continue;
            }
        }
        out.push(b);
    }
    out.into_iter().map(Value::Object).collect()
}

fn hm(v: &str) -> i64 {
    let mut parts = v.split(':').map(|p| js::number(Some(&json!(p))));
    let h = parts.next().unwrap_or(f64::NAN);
    let m = parts.next().unwrap_or(f64::NAN);
    let z = |x: f64| if x.is_finite() && x != 0.0 { x } else { 0.0 };
    (z(h) * 60.0 + z(m)) as i64
}

struct SlotOptions<'a> {
    start: &'a str,
    end: &'a str,
    day_start: &'a str,
    day_end: &'a str,
    min_minutes: i64,
    weekends: bool,
}

fn common_free_slots(busy: &[Value], o: SlotOptions) -> Vec<Value> {
    let (ds, de) = (hm(o.day_start), hm(o.day_end));
    let (Some(range_start), Some(range_end)) = (wall_minutes(o.start), wall_minutes(o.end)) else {
        return Vec::new();
    };
    if de <= ds {
        return Vec::new();
    }
    let mut blocks: Vec<(i64, i64)> = busy
        .iter()
        .filter_map(|b| {
            Some((
                wall_minutes(&js::get_str(b, &["start"]))?,
                wall_minutes(&js::get_str(b, &["end"]))?,
            ))
        })
        .filter(|(a, b)| b > a)
        .collect();
    blocks.sort_by_key(|x| x.0);
    let mut slots = Vec::new();
    let mut day = range_start.div_euclid(1440) * 1440;
    while day < range_end {
        let date = Utc
            .timestamp_opt(day * 60, 0)
            .single()
            .map(|d| d.date_naive());
        let weekday = date.map(weekday_sun0).unwrap_or(1);
        if o.weekends || (weekday != 0 && weekday != 6) {
            let mut cursor = (day + ds).max(range_start);
            let stop = (day + de).min(range_end);
            for &(bs, be) in &blocks {
                if be <= cursor || bs >= stop {
                    continue;
                }
                if bs - cursor >= o.min_minutes {
                    slots.push((cursor, bs));
                }
                cursor = cursor.max(be);
            }
            if stop - cursor >= o.min_minutes {
                slots.push((cursor, stop));
            }
        }
        day += 1440;
    }
    slots
        .into_iter()
        .map(|(a, b)| json!({ "start": minutes_to_wall(a), "end": minutes_to_wall(b), "minutes": b - a }))
        .collect()
}

pub fn outlook_free_busy(ctx: &Ctx, a: FreeBusyArgs) -> Result<Value> {
    let mut people: Vec<String> = Vec::new();
    for r in parse_recipients(a.emails) {
        let e = js::get_str(&r, &["EmailAddress", "Address"]).to_lowercase();
        if !people.contains(&e) {
            people.push(e);
        }
    }
    if people.is_empty() {
        return Ok(json!({
            "ok": false, "status": "missing_emails",
            "answer_text": "Give one or more email addresses, separated by commas.",
        }));
    }
    if people.len() > FREEBUSY_MAX_PEOPLE {
        return Ok(json!({
            "ok": false, "status": "too_many_emails",
            "answer_text": format!("Check at most {FREEBUSY_MAX_PEOPLE} people at a time."),
        }));
    }
    let Some(target) = free_busy_account(ctx, a.account, &people) else {
        let names = account_names(ctx);
        let list = if names.is_empty() {
            "none set up".to_string()
        } else {
            names.join(", ")
        };
        return Ok(json!({
            "ok": false, "status": "account_required",
            "answer_text": format!("Say which account to check from: {list}."),
        }));
    };
    let with_seconds = |v: String| {
        let b = v.as_bytes();
        let n = b.len();
        if n >= 6
            && b[n - 6] == b'T'
            && b[n - 5].is_ascii_digit()
            && b[n - 4].is_ascii_digit()
            && b[n - 3] == b':'
            && b[n - 2].is_ascii_digit()
            && b[n - 1].is_ascii_digit()
        {
            format!("{v}:00")
        } else {
            v
        }
    };
    let (s, e) = normalize_et_local_range(
        ctx,
        a.range.today,
        a.range.tomorrow,
        a.range.start_date,
        a.range.end_date,
    );
    let (start, end) = (with_seconds(s), with_seconds(e));
    let Some(auth) = safe_get_best_token(ctx, &target, "calendar")? else {
        return Ok(json!({
            "ok": false, "status": "needs_sign_in", "account": target.label,
            "answer_text": format!("The {} Outlook session needs signing in again.", target.label),
        }));
    };
    let url = api_url(&rest_base(ctx), "/api/v2.0/me/calendar/getschedule", &[])?;
    let payload = json!({
        "Schedules": people,
        "StartTime": { "DateTime": start, "TimeZone": WINDOWS_TZ },
        "EndTime": { "DateTime": end, "TimeZone": WINDOWS_TZ },
        "AvailabilityViewInterval": 15,
    });
    // Without the Prefer header, getSchedule answers in UTC whatever time zone the request used.
    let req = ctx
        .http
        .post(url)
        .header("authorization", format!("Bearer {}", auth.token))
        .header("accept", "application/json")
        .header("content-type", "application/json")
        .header("prefer", format!("outlook.timezone=\"{ET_TIMEZONE}\""))
        .body(payload.to_string());
    let res = send(req, DEFAULT_TIMEOUT, "outlook_rest")?;
    if !res.ok {
        return Ok(json!({
            "ok": false, "status": "freebusy_failed", "account": target.label,
            "upstream_http_status": res.status, "upstream_error": res.error,
            "answer_text": "Outlook didn't return free/busy times.",
        }));
    }
    let schedules: Vec<Value> = res
        .body
        .as_ref()
        .and_then(|b| b.get("value"))
        .and_then(Value::as_array)
        .cloned()
        .unwrap_or_default();
    let empty = json!({});
    let results: Vec<Value> = people
        .iter()
        .map(|email| {
            let sched = schedules
                .iter()
                .find(|x| js::get_str(x, &["ScheduleId"]).to_lowercase() == *email)
                .unwrap_or(&empty);
            if let Some(err) = sched.get("Error").filter(|v| js::truthy(v)) {
                let msg = normalize_string(err.get("Message"));
                return json!({
                    "email": email, "ok": false,
                    "error": if msg.is_empty() { "unavailable".to_string() } else { msg },
                    "busy": [],
                });
            }
            let mut busy: Vec<Map<String, Value>> = sched
                .get("ScheduleItems")
                .and_then(Value::as_array)
                .into_iter()
                .flatten()
                .filter(|i| {
                    i.get("Status").is_some_and(js::truthy)
                        && i.get("Status") != Some(&json!("Free"))
                })
                .map(|i| {
                    let mut m = Map::new();
                    m.insert(
                        "start".into(),
                        json!(wall_time(js::get(i, &["Start", "DateTime"]))),
                    );
                    m.insert(
                        "end".into(),
                        json!(wall_time(js::get(i, &["End", "DateTime"]))),
                    );
                    m.insert(
                        "status".into(),
                        json!(js::string(&i["Status"]).to_lowercase()),
                    );
                    if let Some(s) = i.get("Subject").filter(|v| js::truthy(v)) {
                        m.insert("subject".into(), s.clone());
                    }
                    if let Some(l) = i.get("Location").filter(|v| js::truthy(v)) {
                        m.insert("location".into(), l.clone());
                    }
                    m
                })
                .collect();
            busy.sort_by_key(|x| js::str_or_empty(x.get("start")));
            json!({ "email": email, "ok": true, "busy": merge_busy(busy) })
        })
        .collect();
    let known: Vec<&Value> = results.iter().filter(|r| r["ok"] == json!(true)).collect();
    let day_start = match a.day_start {
        Some(v) => js::string(v),
        None => "09:00".into(),
    };
    let day_end = match a.day_end {
        Some(v) => js::string(v),
        None => "17:00".into(),
    };
    let free_for_all = if known.is_empty() {
        Vec::new()
    } else {
        let all_busy: Vec<Value> = known
            .iter()
            .flat_map(|r| r["busy"].as_array().cloned().unwrap_or_default())
            .collect();
        common_free_slots(
            &all_busy,
            SlotOptions {
                start: &start,
                end: &end,
                day_start: &day_start,
                day_end: &day_end,
                min_minutes: js::clamp_integer(a.min_minutes, 5, 480, 30),
                weekends: js::to_boolean(a.weekends, false),
            },
        )
    };
    let unavailable: Vec<String> = results
        .iter()
        .filter(|r| r["ok"] != json!(true))
        .map(|r| js::get_str(r, &["email"]))
        .collect();
    let mut answer = Vec::new();
    if known.is_empty() {
        answer.push("No free/busy information came back.".to_string());
    } else {
        answer.push(format!(
            "{} slot{} when {} free (Eastern, {day_start}-{day_end}).",
            free_for_all.len(),
            plural(free_for_all.len()),
            if known.len() == 1 {
                "they are"
            } else {
                "everyone is"
            }
        ));
    }
    if !unavailable.is_empty() {
        answer.push(format!("No free/busy for {}.", unavailable.join(", ")));
    }
    Ok(json!({
        "ok": !known.is_empty(),
        "status": if known.is_empty() { "freebusy_unavailable" } else { "ok" },
        "account": target.label,
        "timezone": ET_TIMEZONE,
        "start": start,
        "end": end,
        "working_day": format!("{day_start}-{day_end}"),
        "people": results,
        "free_for_all": free_for_all,
        "answer_text": answer.join(" "),
    }))
}
