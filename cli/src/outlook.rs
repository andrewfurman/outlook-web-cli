//! Outlook mail and calendar through the Outlook web session (see session.rs): no app
//! registration and no admin consent, the same access the signed-in web app has.
//! No tokens are ever logged. Commands are read-only, except the draft commands
//! (drafts.rs), which save or discard drafts like Outlook web's Reply / Reply all /
//! New message do. Nothing here ever sends mail.
//!
//! Tests inject a token (OUTLOOK_TEST_TOKEN[_<ACCOUNT>]) and HTTP base URLs.

use crate::accounts::{get_account, list_accounts, resolve_account_alias, Account};
use crate::cdp::{connect_cdp, find_outlook_tab};
use crate::ctx::{now_ms, Ctx};
use crate::html::html_to_text;
use crate::js::{self, normalize_string, plural, trim_to};
use crate::session::{decode_jwt, get_session_token};
use anyhow::{bail, Result};
use base64::Engine;
use chrono::{DateTime, Datelike, Local, NaiveDate, NaiveDateTime, TimeZone, Utc};
use chrono_tz::America::New_York;
use percent_encoding::{utf8_percent_encode, AsciiSet, NON_ALPHANUMERIC};
use serde_json::{json, Map, Value};
use std::time::{Duration, Instant};

const DEFAULT_LIST_LIMIT: i64 = 20;
const MAX_LIST_LIMIT: i64 = 50;
const DEFAULT_READ_MAX_BODY_CHARS: i64 = 20000;
const MAX_READ_MAX_BODY_CHARS: i64 = 60000;
pub(crate) const DEFAULT_TIMEOUT: Duration = Duration::from_secs(15);
pub(crate) const ET_TIMEZONE: &str = "America/New_York";
// `attachments` returns base64 in the response only (nothing is written to disk): whole
// files up to max_data_bytes, or a byte range (offset/length) of one attachment so a
// caller with a small output limit, such as SSM, can pull a big image in pieces.
const DEFAULT_ATTACHMENT_DATA_BYTES: i64 = 250_000;
const MAX_ATTACHMENT_DATA_BYTES: i64 = 1_000_000;
const MAX_INSPECT_IMAGE_BYTES: f64 = 5_000_000.0;
const ATTACHMENT_TIMEOUT: Duration = Duration::from_secs(30);
// A message's attachments stay in memory for a few minutes (never on disk), so pulling
// one image in many byte ranges in one process downloads it from Outlook once.
const ATTACHMENT_CACHE_TTL: Duration = Duration::from_secs(5 * 60);
const ATTACHMENT_CACHE_MAX_ENTRIES: usize = 4;
const DEFAULT_ATTACHMENT_INSPECT_PROMPT: &str = "Describe this screenshot or image in detail, then transcribe ALL visible text verbatim in reading order (menus, labels, table cells, numbers, dates). Put the full transcription in ocr_text.";
// Exchange rejects $orderby on a property that is not also the first $filter clause.
const ALL_TIME_UTC: &str = "1900-01-01T00:00:00Z";

/// `messages` lists one well-known folder, Inbox by default. Listing all of
/// /me/messages mixes in Drafts and Sent Items, so an unsent draft looks like a sent reply.
fn resolve_mail_folder(folder: &str) -> Option<&'static str> {
    let key: String = folder
        .to_lowercase()
        .chars()
        .filter(|c| !c.is_whitespace() && *c != '_' && *c != '-')
        .collect();
    Some(match key.as_str() {
        "inbox" => "Inbox",
        "sent" | "sentitems" => "SentItems",
        "drafts" => "Drafts",
        "deleted" | "deleteditems" => "DeletedItems",
        "archive" => "Archive",
        "junk" | "junkemail" => "JunkEmail",
        _ => return None,
    })
}

// ---------------------------------------------------------------------------
// HTTP
// ---------------------------------------------------------------------------

/// `encodeURIComponent`.
const COMPONENT: &AsciiSet = &NON_ALPHANUMERIC
    .remove(b'-')
    .remove(b'_')
    .remove(b'.')
    .remove(b'!')
    .remove(b'~')
    .remove(b'*')
    .remove(b'\'')
    .remove(b'(')
    .remove(b')');

pub(crate) fn enc(s: &str) -> String {
    utf8_percent_encode(s, COMPONENT).to_string()
}

pub(crate) fn graph_base(ctx: &Ctx) -> String {
    ctx.var("OUTLOOK_GRAPH_BASE_URL")
        .map(str::trim)
        .filter(|s| !s.is_empty())
        .unwrap_or("https://graph.microsoft.com")
        .to_string()
}

pub(crate) fn rest_base(ctx: &Ctx) -> String {
    ctx.var("OUTLOOK_REST_BASE_URL")
        .map(str::trim)
        .filter(|s| !s.is_empty())
        .unwrap_or("https://outlook.office.com")
        .to_string()
}

/// `new URL(path, base + "/")` with `path` absolute, plus query parameters in order.
pub(crate) fn api_url(base: &str, path: &str, params: &[(&str, String)]) -> Result<url::Url> {
    let base = if base.ends_with('/') {
        base.to_string()
    } else {
        format!("{base}/")
    };
    let mut url = url::Url::parse(&base)?.join(path)?;
    if !params.is_empty() {
        url.set_query(None);
        let mut q = url.query_pairs_mut();
        for (k, v) in params {
            q.append_pair(k, v);
        }
    }
    Ok(url)
}

#[derive(Clone, Debug)]
pub struct Resp {
    pub ok: bool,
    pub status: u16,
    pub body: Option<Value>,
    pub error: String,
    pub api: &'static str,
}

fn upstream_error(body: &Option<Value>, text: &str) -> String {
    let m = body
        .as_ref()
        .map(|b| js::get_str(b, &["error", "message"]))
        .unwrap_or_default();
    if m.is_empty() {
        text.chars().take(200).collect()
    } else {
        m
    }
}

pub(crate) fn send(
    req: reqwest::blocking::RequestBuilder,
    timeout: Duration,
    api: &'static str,
) -> Result<Resp> {
    let res = req.timeout(timeout).send()?;
    let status = res.status();
    if status.is_redirection() {
        bail!("fetch failed: unexpected redirect");
    }
    let text = res.text().unwrap_or_default();
    let body = serde_json::from_str::<Value>(&text).ok();
    let ok = status.is_success();
    Ok(Resp {
        ok,
        status: status.as_u16(),
        error: if ok {
            String::new()
        } else {
            upstream_error(&body, &text)
        },
        body,
        api,
    })
}

fn fetch_json(
    ctx: &Ctx,
    url: url::Url,
    headers: &[(&str, String)],
    timeout: Duration,
    api: &'static str,
) -> Result<Resp> {
    let mut req = ctx.http.get(url);
    for (k, v) in headers {
        req = req.header(*k, v);
    }
    send(req, timeout, api)
}

fn bearer_headers(token: &str) -> Vec<(&'static str, String)> {
    vec![
        ("authorization", format!("Bearer {token}")),
        ("accept", "application/json".into()),
    ]
}

fn graph_headers(
    token: &str,
    consistency: bool,
    prefer_tz: Option<&str>,
) -> Vec<(&'static str, String)> {
    let mut h = bearer_headers(token);
    if consistency {
        h.push(("consistency-level", "eventual".into()));
    }
    if let Some(tz) = prefer_tz {
        h.push(("prefer", format!("outlook.timezone=\"{tz}\"")));
    }
    h
}

fn should_fallback(res: &Resp) -> bool {
    !res.ok && matches!(res.status, 401 | 403 | 404)
}

fn other_api(api: &str) -> &'static str {
    if api == "graph" {
        "outlook_rest"
    } else {
        "graph"
    }
}

/// Tries one API, then the other on 401/403/404.
fn with_fallback(first: &str, mut attempt: impl FnMut(&str) -> Result<Resp>) -> Result<Resp> {
    let res = attempt(first)?;
    if should_fallback(&res) {
        return attempt(other_api(first));
    }
    Ok(res)
}

// ---------------------------------------------------------------------------
// Tokens and accounts
// ---------------------------------------------------------------------------

#[derive(Clone, Debug)]
pub(crate) struct Auth {
    pub token: String,
    pub api: String,
    pub source: String,
}

pub fn preferred_api_for_token(claims: &Value) -> &'static str {
    let aud = js::get_str(claims, &["aud"]).to_lowercase();
    if aud.contains("graph.microsoft.com") || aud == "00000003-0000-0000-c000-000000000000" {
        return "graph";
    }
    if aud.contains("outlook.office.com")
        || aud == "00000002-0000-0ff1-ce00-000000000000"
        || aud.contains("substrate.office.com")
    {
        return "outlook_rest";
    }
    "graph"
}

/// The token for an account: a test override, the saved web session, or as a last
/// resort a live token read from the account's Outlook web tab.
pub(crate) fn safe_get_best_token(
    ctx: &Ctx,
    target: &Account,
    purpose: &str,
) -> Result<Option<Auth>> {
    let key: String = target
        .label
        .to_uppercase()
        .chars()
        .map(|c| {
            if c.is_ascii_uppercase() || c.is_ascii_digit() {
                c
            } else {
                '_'
            }
        })
        .collect();
    let env = ctx
        .var(&format!("OUTLOOK_TEST_TOKEN_{key}"))
        .or_else(|| ctx.var("OUTLOOK_TEST_TOKEN"));
    if let Some(t) = env.filter(|t| !t.trim().is_empty()) {
        return Ok(Some(Auth {
            token: t.to_string(),
            api: preferred_api_for_token(&decode_jwt(t)).into(),
            source: String::new(),
        }));
    }
    // The saved web session's token is for https://outlook.office.com with the web app's
    // Mail.ReadWrite and Calendars scopes, so it serves every purpose.
    if let Some(s) = get_session_token(ctx, &target.label, true)? {
        return Ok(Some(Auth {
            token: s.token,
            api: s.api,
            source: s.source,
        }));
    }
    Ok(match target.cdp_port {
        Some(port) => best_token_via_cdp(ctx, port, &target.main_url, purpose),
        None => None,
    })
}

const TOKEN_PROBE: &str = r#"(function() {
  try {
    var out = [];
    for (var i = 0; i < localStorage.length; i++) {
      var k = localStorage.key(i) || "";
      var raw = null;
      try { raw = localStorage.getItem(k) || ""; } catch(e) {}
      if (!raw) continue;
      try {
        var obj = JSON.parse(raw);
        if (k.startsWith("msal.3|") && obj && obj.credentialType === "AccessToken" && typeof obj.secret === "string") {
          out.push({ source: "msal_v3", target: String(obj.target||""), expiresOn: Number(obj.expiresOn||0), token: obj.secret });
          continue;
        }
        if (k.startsWith("accesstoken-") && (obj.secret || obj.accessToken || obj.credential)) {
          out.push({ source: "msal_v2", target: String(obj.target||obj.scopes||""), expiresOn: Number(obj.expiresOn||obj.expires_on||0), token: obj.secret || obj.accessToken || obj.credential });
          continue;
        }
      } catch (e) {
        // ignore parse errors
      }
    }
    return { ok: true, tokens: out };
  } catch (e3) {
    return { ok: false, status: "token_probe_failed" };
  }
})()"#;

fn best_token_via_cdp(ctx: &Ctx, cdp_port: u16, main_url: &str, purpose: &str) -> Option<Auth> {
    let tab = find_outlook_tab(ctx, cdp_port, Some(main_url))?;
    let mut cdp = connect_cdp(&tab.web_socket_debugger_url)?;
    let scan = |cdp: &mut crate::cdp::Cdp| -> Option<Vec<Value>> {
        let v = cdp.evaluate(TOKEN_PROBE).ok()?;
        if v.get("ok") == Some(&json!(true)) {
            v.get("tokens").and_then(Value::as_array).cloned()
        } else {
            Some(Vec::new())
        }
    };
    let mut best = pick_best_token(&scan(&mut cdp)?, purpose);
    if best.is_none() {
        // Reload once so MSAL can refresh silently, then look again.
        let _ = cdp.send("Page.reload", json!({ "ignoreCache": true }));
        std::thread::sleep(Duration::from_millis(1500));
        best = pick_best_token(&scan(&mut cdp)?, purpose);
    }
    best.map(|b| Auth {
        token: b.token,
        api: b.api.into(),
        source: String::new(),
    })
}

#[derive(Clone, Debug, PartialEq)]
pub struct BestToken {
    pub token: String,
    pub api: &'static str,
    pub target: String,
    pub expires_on: f64,
}

/// Picks the longest-lived token from a tab's MSAL cache (entries with `target`,
/// `expiresOn` and `token`) that has the scope `purpose` needs; Outlook REST first.
pub fn pick_best_token(tokens: &[Value], purpose: &str) -> Option<BestToken> {
    let now = (now_ms() / 1000) as f64;
    let margin = 60.0;
    let need = match purpose {
        "calendar" => "Calendars.Read",
        "mail_write" => "Mail.ReadWrite",
        _ => "Mail.Read",
    };
    let valid: Vec<(String, f64, String)> = tokens
        .iter()
        .map(|t| {
            (
                js::str_or_empty(t.get("target")),
                match t.get("expiresOn").filter(|v| js::truthy(v)) {
                    Some(v) => js::number(Some(v)),
                    None => 0.0,
                },
                js::str_or_empty(t.get("token")),
            )
        })
        .filter(|(_, _, tok)| tok.chars().count() > 100)
        .filter(|(_, exp, _)| *exp > now + margin)
        .collect();
    let by_pref = |domain: &str| {
        let mut c: Vec<&(String, f64, String)> = valid
            .iter()
            .filter(|(t, _, _)| t.contains(domain) && t.contains(need))
            .collect();
        c.sort_by(|a, b| b.1.partial_cmp(&a.1).unwrap_or(std::cmp::Ordering::Equal));
        c.first().map(|x| (*x).clone())
    };
    let make = |x: (String, f64, String), api| BestToken {
        token: x.2,
        api,
        target: x.0,
        expires_on: x.1,
    };
    if let Some(o) = by_pref("outlook.office.com") {
        return Some(make(o, "outlook_rest"));
    }
    // Drafts are written through the Outlook REST API only.
    if purpose == "mail_write" {
        return None;
    }
    by_pref("graph.microsoft.com").map(|g| make(g, "graph"))
}

pub(crate) fn account_names(ctx: &Ctx) -> Vec<String> {
    list_accounts(ctx).into_iter().map(|a| a.name).collect()
}

pub(crate) fn resolve_accounts(ctx: &Ctx, scope: Option<&Value>) -> Result<Vec<Account>> {
    let raw = normalize_string(scope);
    let s = if raw.to_lowercase() == "all" {
        "all".to_string()
    } else {
        resolve_account_alias(ctx, scope).unwrap_or(raw)
    };
    let all = list_accounts(ctx);
    if s.is_empty() || s == "all" {
        if all.is_empty() {
            bail!("No Outlook accounts are set up. Run `outlook-web-cli auth chrome --account <name> --cdp-port <port>` first.");
        }
        return Ok(all);
    }
    match get_account(ctx, &s) {
        Some(a) => Ok(vec![a]),
        None => {
            let mut names = account_names(ctx);
            names.push("all".into());
            let shown = scope.map(js::string).unwrap_or_else(|| "undefined".into());
            bail!("Unknown account '{shown}'. Use {}.", names.join(", "))
        }
    }
}

pub(crate) fn account_scope_label(scope: Option<&Value>) -> String {
    let s = normalize_string(scope);
    if s.is_empty() {
        "all".into()
    } else {
        s
    }
}

pub(crate) fn missing_field(field: &str, message: &str) -> Value {
    json!({ "ok": false, "status": "missing_field", "field": field, "message": message })
}

fn first_api(auth: &Auth) -> String {
    if auth.api.is_empty() {
        preferred_api_for_token(&decode_jwt(&auth.token)).into()
    } else {
        auth.api.clone()
    }
}

// ---------------------------------------------------------------------------
// Time
// ---------------------------------------------------------------------------

/// `Date.parse`-like: ISO date (UTC), date-time with zone, or local date-time.
pub(crate) fn parse_js_date(s: &str) -> Option<DateTime<Utc>> {
    let s = s.trim();
    if let Ok(d) = NaiveDate::parse_from_str(s, "%Y-%m-%d") {
        return Some(d.and_hms_opt(0, 0, 0)?.and_utc());
    }
    if let Ok(d) = DateTime::parse_from_rfc3339(s) {
        return Some(d.with_timezone(&Utc));
    }
    for f in ["%Y-%m-%dT%H:%M:%S%.f%:z", "%Y-%m-%dT%H:%M%:z"] {
        if let Ok(d) = DateTime::parse_from_str(&s.replace('Z', "+00:00"), f) {
            return Some(d.with_timezone(&Utc));
        }
    }
    for f in ["%Y-%m-%dT%H:%M:%S%.f", "%Y-%m-%dT%H:%M"] {
        if let Ok(n) = NaiveDateTime::parse_from_str(s, f) {
            return Local
                .from_local_datetime(&n)
                .earliest()
                .map(|d| d.with_timezone(&Utc));
        }
    }
    None
}

fn iso_millis(d: DateTime<Utc>) -> String {
    d.to_rfc3339_opts(chrono::SecondsFormat::Millis, true)
}

/// A `YYYY-MM-DD` date is UTC midnight (as the earlier version did); anything else
/// parseable is converted to UTC. "" when empty or invalid.
fn to_iso_utc_string(input: Option<&Value>) -> String {
    let s = normalize_string(input);
    if s.is_empty() {
        return String::new();
    }
    parse_js_date(&s).map(iso_millis).unwrap_or_default()
}

/// A UTC timestamp as an Eastern wall-clock `YYYY-MM-DDTHH:MM:SS`.
pub(crate) fn to_et_local_string(iso_utc: &str) -> String {
    if iso_utc.is_empty() {
        return String::new();
    }
    match parse_js_date(iso_utc) {
        Some(d) => d
            .with_timezone(&New_York)
            .format("%Y-%m-%dT%H:%M:%S")
            .to_string(),
        None => String::new(),
    }
}

/// The Eastern calendar date `offset_days` days after `now_ms`.
fn et_ymd(now_ms: i64, offset_days: i64) -> String {
    let t = now_ms + offset_days * 24 * 60 * 60 * 1000;
    Utc.timestamp_millis_opt(t)
        .single()
        .unwrap_or_else(Utc::now)
        .with_timezone(&New_York)
        .format("%Y-%m-%d")
        .to_string()
}

fn is_ymd(s: &str) -> bool {
    s.len() == 10 && NaiveDate::parse_from_str(s, "%Y-%m-%d").is_ok() && s.as_bytes()[4] == b'-'
}

/// Eastern wall times `YYYY-MM-DDTHH:MM:SS` for APIs that honor Prefer: outlook.timezone.
pub(crate) fn normalize_et_local_range(
    ctx: &Ctx,
    today: Option<&Value>,
    tomorrow: Option<&Value>,
    start_date: Option<&Value>,
    end_date: Option<&Value>,
) -> (String, String) {
    let now = ctx.now_ms_for_ranges();
    let day = |d: String| (format!("{d}T00:00:00"), format!("{d}T23:59:59"));
    if js::to_boolean(today, false) {
        return day(et_ymd(now, 0));
    }
    if js::to_boolean(tomorrow, false) {
        return day(et_ymd(now, 1));
    }
    let s = normalize_string(start_date);
    let e = normalize_string(end_date);
    let start = if is_ymd(&s) {
        format!("{s}T00:00:00")
    } else {
        s
    };
    let end = if is_ymd(&e) {
        format!("{e}T23:59:59")
    } else {
        e
    };
    if !start.is_empty() && !end.is_empty() {
        return (start, end);
    }
    day(et_ymd(now, 0))
}

// ---------------------------------------------------------------------------
// Compact shapes
// ---------------------------------------------------------------------------

fn cap(pascal: bool, name: &str) -> String {
    if !pascal {
        return name.to_string();
    }
    let mut c = name.chars();
    match c.next() {
        Some(f) => f.to_uppercase().collect::<String>() + c.as_str(),
        None => String::new(),
    }
}

/// Graph (camelCase) or Outlook REST (PascalCase) message -> one compact shape.
fn compact_message(
    msg: &Value,
    account: &str,
    pascal: bool,
    body: Option<i64>,
) -> Map<String, Value> {
    let k = |n: &str| cap(pascal, n);
    let g = |path: &[&str]| {
        let keys: Vec<String> = path.iter().map(|p| k(p)).collect();
        let refs: Vec<&str> = keys.iter().map(String::as_str).collect();
        js::get_str(msg, &refs)
    };
    let received = g(&["receivedDateTime"]);
    let mut out = Map::new();
    out.insert("account".into(), json!(account));
    out.insert("id".into(), json!(g(&["id"])));
    out.insert("subject".into(), json!(trim_to(&g(&["subject"]), 400)));
    out.insert(
        "from".into(),
        json!({ "name": g(&["from", "emailAddress", "name"]), "address": g(&["from", "emailAddress", "address"]) }),
    );
    out.insert(
        "is_unread".into(),
        json!(msg.get(k("isRead")) == Some(&json!(false))),
    );
    out.insert(
        "is_draft".into(),
        json!(msg.get(k("isDraft")) == Some(&json!(true))),
    );
    out.insert("received_datetime".into(), json!(received.clone()));
    out.insert("received_at".into(), json!(to_et_local_string(&received)));
    out.insert("web_link".into(), json!(g(&["webLink"])));
    out.insert("conversation_id".into(), json!(g(&["conversationId"])));
    if let Some(max) = body {
        let kind = g(&["body", "contentType"]).to_lowercase();
        let raw = g(&["body", "content"]);
        let text = if kind == "text" {
            raw
        } else {
            html_to_text(&raw)
        };
        out.insert(
            "body_text".into(),
            json!(trim_to(&text, max.max(0) as usize)),
        );
    }
    out
}

fn compact_event(evt: &Value, account: &str, pascal: bool) -> Value {
    let k = |n: &str| cap(pascal, n);
    let g = |path: &[&str]| {
        let keys: Vec<String> = path.iter().map(|p| k(p)).collect();
        let refs: Vec<&str> = keys.iter().map(String::as_str).collect();
        js::get_str(evt, &refs)
    };
    let when = |which: &str| {
        let tz = g(&[which, "timeZone"]);
        json!({
            "dateTime": g(&[which, "dateTime"]),
            "timeZone": if tz.is_empty() { ET_TIMEZONE.to_string() } else { tz },
        })
    };
    json!({
        "account": account,
        "id": g(&["id"]),
        "subject": trim_to(&g(&["subject"]), 400),
        "is_all_day": evt.get(k("isAllDay")).is_some_and(js::truthy),
        "location": g(&["location", "displayName"]),
        "organizer": { "name": g(&["organizer", "emailAddress", "name"]), "address": g(&["organizer", "emailAddress", "address"]) },
        "start": when("start"),
        "end": when("end"),
        "web_link": g(&["webLink"]),
    })
}

/// Every account's items in one list, newest first.
fn merge_account_lists(lists: &[Value]) -> Vec<Value> {
    let mut merged: Vec<Value> = lists
        .iter()
        .filter(|l| l.get("ok") == Some(&json!(true)))
        .flat_map(|l| {
            l.get("items")
                .and_then(Value::as_array)
                .cloned()
                .unwrap_or_default()
        })
        .collect();
    let t = |v: &Value| {
        let s = match v.get("received_datetime").filter(|x| js::truthy(x)) {
            Some(x) => js::string(x),
            None => js::get_str(v, &["start", "dateTime"]),
        };
        parse_js_date(&s).map(|d| d.timestamp_millis()).unwrap_or(0)
    };
    merged.sort_by_key(|v| std::cmp::Reverse(t(v)));
    merged
}

fn values(body: &Option<Value>) -> Vec<Value> {
    body.as_ref()
        .and_then(|b| b.get("value"))
        .and_then(Value::as_array)
        .cloned()
        .unwrap_or_default()
}

fn needs_sign_in_list(account: &str) -> Value {
    json!({ "ok": false, "account": account, "status": "needs_sign_in", "returned_count": 0, "items": [] })
}

// ---------------------------------------------------------------------------
// Commands
// ---------------------------------------------------------------------------

pub fn outlook_status(ctx: &Ctx, account: Option<&Value>) -> Result<Value> {
    let targets = resolve_accounts(ctx, account)?;
    let mut summaries = Vec::new();
    for t in &targets {
        let auth = safe_get_best_token(ctx, t, "mail").ok().flatten();
        let ok = auth.is_some();
        summaries.push(json!({
            "account": t.label,
            "ok": ok,
            "status": if ok { "signed_in" } else { "needs_sign_in" },
            "token_source": match &auth {
                Some(a) if a.source.is_empty() => "chrome_tab".to_string(),
                Some(a) => a.source.clone(),
                None => String::new(),
            },
        }));
    }
    let signed_in = summaries.iter().filter(|s| s["ok"] == json!(true)).count();
    let answer = if summaries.len() == 1 {
        if signed_in == 1 {
            "Signed in.".to_string()
        } else {
            "Needs sign-in.".to_string()
        }
    } else {
        format!(
            "{signed_in}/{} Outlook sessions are signed in.",
            summaries.len()
        )
    };
    Ok(json!({
        "ok": true,
        "status": "ok",
        "accounts_checked": summaries.len(),
        "accounts": summaries,
        "answer_text": answer,
    }))
}

/// The received-date clause goes first so `$orderby` on it is accepted with other filters.
fn list_filter(
    unread: Option<&Value>,
    from: Option<&Value>,
    since: Option<&Value>,
    received: &str,
    is_read: &str,
    from_address: &str,
) -> String {
    let sender = normalize_string(from);
    let since_iso = to_iso_utc_string(since);
    let mut filters = Vec::new();
    if !since_iso.is_empty() || !sender.is_empty() {
        let at = if since_iso.is_empty() {
            ALL_TIME_UTC.to_string()
        } else {
            since_iso
        };
        filters.push(format!("{received} ge {at}"));
    }
    if js::to_boolean(unread, false) {
        filters.push(format!("{is_read} eq false"));
    }
    if !sender.is_empty() {
        filters.push(format!(
            "{from_address} eq '{}'",
            sender.replace('\'', "''")
        ));
    }
    filters.join(" and ")
}

pub struct MessagesArgs<'a> {
    pub account: Option<&'a Value>,
    pub limit: Option<&'a Value>,
    pub unread: Option<&'a Value>,
    pub from: Option<&'a Value>,
    pub since: Option<&'a Value>,
    pub search: Option<&'a Value>,
    pub folder: Option<&'a Value>,
}

pub fn outlook_messages(ctx: &Ctx, a: MessagesArgs) -> Result<Value> {
    let folder_raw = match a.folder.filter(|v| js::truthy(v)) {
        Some(v) => v.clone(),
        None => match a.folder {
            None => json!("inbox"),
            Some(v) => v.clone(),
        },
    };
    let folder_key = {
        let s = normalize_string(Some(&folder_raw));
        if s.is_empty() {
            "inbox".to_string()
        } else {
            s
        }
    };
    let Some(folder) = resolve_mail_folder(&folder_key) else {
        let shown = js::string(&folder_raw);
        return Ok(json!({
            "ok": false,
            "status": "invalid_folder",
            "folder": shown,
            "answer_text": format!("Unknown Outlook folder \"{shown}\". Use inbox, sent, drafts, deleted, archive or junk."),
        }));
    };
    let targets = resolve_accounts(ctx, a.account)?;
    let limit = js::clamp_integer(a.limit, 1, MAX_LIST_LIMIT, DEFAULT_LIST_LIMIT);
    let search = normalize_string(a.search);
    let mut queries = Vec::new();
    for t in &targets {
        let Some(auth) = safe_get_best_token(ctx, t, "mail")? else {
            queries.push(needs_sign_in_list(&t.label));
            continue;
        };
        let first = first_api(&auth);
        let res = with_fallback(&first, |api| {
            if api == "graph" {
                let mut p = vec![
                    (
                        "$select",
                        "id,subject,from,receivedDateTime,isRead,isDraft,conversationId,webLink"
                            .to_string(),
                    ),
                    ("$top", limit.to_string()),
                ];
                if !search.is_empty() {
                    // Graph returns search hits by relevance and rejects $orderby with $search.
                    p.push(("$search", format!("\"{search}\"")));
                } else {
                    p.push(("$orderby", "receivedDateTime desc".into()));
                    let f = list_filter(
                        a.unread,
                        a.from,
                        a.since,
                        "receivedDateTime",
                        "isRead",
                        "from/emailAddress/address",
                    );
                    if !f.is_empty() {
                        p.push(("$filter", f));
                    }
                }
                let url = api_url(
                    &graph_base(ctx),
                    &format!("/v1.0/me/mailFolders/{folder}/messages"),
                    &p,
                )?;
                fetch_json(
                    ctx,
                    url,
                    &graph_headers(&auth.token, !search.is_empty(), None),
                    DEFAULT_TIMEOUT,
                    "graph",
                )
            } else {
                let mut p = vec![
                    (
                        "$select",
                        "Id,Subject,From,ReceivedDateTime,IsRead,IsDraft,ConversationId,WebLink"
                            .to_string(),
                    ),
                    ("$top", limit.to_string()),
                ];
                let f = list_filter(
                    a.unread,
                    a.from,
                    a.since,
                    "ReceivedDateTime",
                    "IsRead",
                    "From/EmailAddress/Address",
                );
                if !f.is_empty() {
                    p.push(("$filter", f.clone()));
                }
                if f.is_empty() && !search.is_empty() {
                    p.push(("$search", format!("\"{search}\"")));
                } else {
                    p.push(("$orderby", "ReceivedDateTime desc".into()));
                }
                let url = api_url(
                    &rest_base(ctx),
                    &format!("/api/v2.0/me/MailFolders/{folder}/messages"),
                    &p,
                )?;
                fetch_json(
                    ctx,
                    url,
                    &bearer_headers(&auth.token),
                    DEFAULT_TIMEOUT,
                    "outlook_rest",
                )
            }
        })?;
        if !res.ok {
            queries.push(json!({
                "ok": false, "account": t.label, "status": "upstream_error",
                "upstream_http_status": res.status, "upstream_error": res.error,
                "returned_count": 0, "items": [],
            }));
            continue;
        }
        let pascal = res.api != "graph";
        let items: Vec<Value> = values(&res.body)
            .iter()
            .map(|m| {
                let mut c = compact_message(m, &t.label, pascal, None);
                c.insert("folder".into(), json!(folder));
                Value::Object(c)
            })
            .collect();
        queries.push(json!({
            "ok": true, "account": t.label, "status": "ok", "folder": folder,
            "returned_count": items.len(), "items": items,
        }));
    }
    let merged = merge_account_lists(&queries);
    let any_error = queries.iter().any(|q| q["ok"] != json!(true));
    let all_failed = queries.iter().all(|q| q["ok"] != json!(true));
    let answer = if merged.is_empty() && any_error {
        "The Outlook upstream call failed.".to_string()
    } else if merged.is_empty() {
        format!("No messages matched in {folder}.")
    } else {
        format!(
            "Found {} message{} in {folder}.",
            merged.len(),
            plural(merged.len())
        )
    };
    Ok(json!({
        "ok": !all_failed,
        "status": if all_failed { "upstream_error" } else { "ok" },
        "scope": account_scope_label(a.account),
        "folder": folder,
        "returned_count": merged.len(),
        "items": merged,
        "answer_text": answer,
        "accounts": queries,
    }))
}

pub fn outlook_read(
    ctx: &Ctx,
    account: Option<&Value>,
    id: Option<&Value>,
    max_body_chars: Option<&Value>,
) -> Result<Value> {
    let message_id = normalize_string(id);
    if message_id.is_empty() {
        return Ok(missing_field("id", "An Outlook message id is required."));
    }
    let targets = resolve_accounts(ctx, account)?;
    let max = js::clamp_integer(
        max_body_chars,
        2000,
        MAX_READ_MAX_BODY_CHARS,
        DEFAULT_READ_MAX_BODY_CHARS,
    );
    // With "all", the first account that has a token answers.
    for t in &targets {
        let Some(auth) = safe_get_best_token(ctx, t, "mail")? else {
            continue;
        };
        let res = with_fallback(&first_api(&auth), |api| {
            if api == "graph" {
                let url = api_url(
                    &graph_base(ctx),
                    &format!("/v1.0/me/messages/{}", enc(&message_id)),
                    &[(
                        "$select",
                        "id,subject,from,receivedDateTime,isDraft,body,webLink".into(),
                    )],
                )?;
                fetch_json(
                    ctx,
                    url,
                    &graph_headers(&auth.token, false, None),
                    DEFAULT_TIMEOUT,
                    "graph",
                )
            } else {
                let url = api_url(
                    &rest_base(ctx),
                    &format!("/api/v2.0/me/messages/{}", enc(&message_id)),
                    &[(
                        "$select",
                        "Id,Subject,From,ReceivedDateTime,IsDraft,Body,WebLink".into(),
                    )],
                )?;
                fetch_json(
                    ctx,
                    url,
                    &bearer_headers(&auth.token),
                    DEFAULT_TIMEOUT,
                    "outlook_rest",
                )
            }
        })?;
        if !res.ok {
            return Ok(json!({
                "ok": false, "status": "message_read_failed", "account": t.label, "id": message_id,
                "upstream_http_status": res.status, "answer_text": "The Outlook upstream call failed.",
            }));
        }
        let raw = res.body.clone().unwrap_or_else(|| json!({}));
        let c = compact_message(&raw, &t.label, res.api != "graph", Some(max));
        let subject = js::get_str(&Value::Object(c.clone()), &["subject"]);
        let body_text = js::get_str(&Value::Object(c.clone()), &["body_text"]);
        let is_draft = c.get("is_draft") == Some(&json!(true));
        let from = js::get_str(&Value::Object(c.clone()), &["from", "address"]);
        let answer = if is_draft {
            format!(
                "Unsent draft: {}.",
                if subject.is_empty() {
                    "(no subject)"
                } else {
                    &subject
                }
            )
        } else if !body_text.is_empty() {
            format!(
                "{} from {}.",
                if subject.is_empty() {
                    "Email"
                } else {
                    &subject
                },
                if from.is_empty() {
                    "unknown sender"
                } else {
                    &from
                }
            )
        } else {
            "Message has no readable body.".into()
        };
        return Ok(json!({
            "ok": true,
            "status": "ok",
            "account": t.label,
            "id": c["id"],
            "subject": c["subject"],
            "from": c["from"],
            "received_at": c["received_at"],
            "received_datetime": c["received_datetime"],
            "is_draft": is_draft,
            "web_link": c["web_link"],
            "body_text": body_text,
            "answer_text": answer,
        }));
    }
    Ok(json!({
        "ok": false,
        "status": "message_not_found",
        "scope": account_scope_label(account),
        "id": message_id,
        "answer_text": "Message not found.",
    }))
}

pub fn outlook_search(
    ctx: &Ctx,
    account: Option<&Value>,
    query: Option<&Value>,
    limit: Option<&Value>,
) -> Result<Value> {
    let q = normalize_string(query);
    if q.is_empty() {
        return Ok(missing_field("query", "A search query is required."));
    }
    let targets = resolve_accounts(ctx, account)?;
    let limit = js::clamp_integer(limit, 1, MAX_LIST_LIMIT, DEFAULT_LIST_LIMIT);
    let mut lists = Vec::new();
    for t in &targets {
        let Some(auth) = safe_get_best_token(ctx, t, "mail")? else {
            lists.push(needs_sign_in_list(&t.label));
            continue;
        };
        let res = with_fallback(&first_api(&auth), |api| {
            if api == "graph" {
                let url = api_url(
                    &graph_base(ctx),
                    "/v1.0/me/messages",
                    &[
                        (
                            "$select",
                            "id,subject,from,receivedDateTime,isRead,isDraft,conversationId,webLink"
                                .into(),
                        ),
                        ("$search", format!("\"{q}\"")),
                        ("$top", limit.to_string()),
                    ],
                )?;
                fetch_json(
                    ctx,
                    url,
                    &graph_headers(&auth.token, true, None),
                    DEFAULT_TIMEOUT,
                    "graph",
                )
            } else {
                let url = api_url(
                    &rest_base(ctx),
                    "/api/v2.0/me/messages",
                    &[
                        (
                            "$select",
                            "Id,Subject,From,ReceivedDateTime,IsRead,IsDraft,ConversationId,WebLink"
                                .into(),
                        ),
                        ("$top", limit.to_string()),
                        ("$search", format!("\"{q}\"")),
                    ],
                )?;
                fetch_json(
                    ctx,
                    url,
                    &bearer_headers(&auth.token),
                    DEFAULT_TIMEOUT,
                    "outlook_rest",
                )
            }
        })?;
        if !res.ok {
            lists.push(json!({
                "ok": false, "account": t.label, "status": "search_failed",
                "upstream_http_status": res.status, "upstream_error": res.error,
                "returned_count": 0, "items": [],
            }));
            continue;
        }
        let pascal = res.api != "graph";
        let items: Vec<Value> = values(&res.body)
            .iter()
            .map(|m| Value::Object(compact_message(m, &t.label, pascal, None)))
            .collect();
        lists.push(json!({ "ok": true, "account": t.label, "status": "ok", "returned_count": items.len(), "items": items }));
    }
    let merged = merge_account_lists(&lists);
    let any_ok = lists.iter().any(|l| l["ok"] == json!(true));
    Ok(json!({
        "ok": any_ok,
        "status": if any_ok { "ok" } else { "search_failed" },
        "scope": account_scope_label(account),
        "returned_count": merged.len(),
        "items": merged,
        "answer_text": if merged.is_empty() { "No results.".to_string() } else { format!("Found {} result{}.", merged.len(), plural(merged.len())) },
        "accounts": lists,
    }))
}

pub struct RangeArgs<'a> {
    pub today: Option<&'a Value>,
    pub tomorrow: Option<&'a Value>,
    pub start_date: Option<&'a Value>,
    pub end_date: Option<&'a Value>,
}

pub fn outlook_agenda(
    ctx: &Ctx,
    account: Option<&Value>,
    range: RangeArgs,
    limit: Option<&Value>,
) -> Result<Value> {
    let targets = resolve_accounts(ctx, account)?;
    let (start, end) = normalize_et_local_range(
        ctx,
        range.today,
        range.tomorrow,
        range.start_date,
        range.end_date,
    );
    let limit = js::clamp_integer(limit, 1, 200, 50);
    let mut lists = Vec::new();
    for t in &targets {
        let Some(auth) = safe_get_best_token(ctx, t, "calendar")? else {
            lists.push(needs_sign_in_list(&t.label));
            continue;
        };
        let res = with_fallback(&first_api(&auth), |api| {
            if api == "graph" {
                let url = api_url(
                    &graph_base(ctx),
                    "/v1.0/me/calendarView",
                    &[
                        ("startDateTime", start.clone()),
                        ("endDateTime", end.clone()),
                        ("$top", limit.to_string()),
                        ("$orderby", "start/dateTime asc".into()),
                        (
                            "$select",
                            "id,subject,start,end,location,organizer,isAllDay,webLink".into(),
                        ),
                    ],
                )?;
                fetch_json(
                    ctx,
                    url,
                    &graph_headers(&auth.token, false, Some(ET_TIMEZONE)),
                    DEFAULT_TIMEOUT,
                    "graph",
                )
            } else {
                let url = api_url(
                    &rest_base(ctx),
                    "/api/v2.0/me/calendarview",
                    &[
                        ("startDateTime", start.clone()),
                        ("endDateTime", end.clone()),
                        ("$top", limit.to_string()),
                        ("$orderby", "Start/DateTime asc".into()),
                        (
                            "$select",
                            "Id,Subject,Start,End,Location,Organizer,IsAllDay,WebLink".into(),
                        ),
                    ],
                )?;
                let mut h = bearer_headers(&auth.token);
                h.push(("prefer", format!("outlook.timezone=\"{ET_TIMEZONE}\"")));
                fetch_json(ctx, url, &h, DEFAULT_TIMEOUT, "outlook_rest")
            }
        })?;
        if !res.ok {
            lists.push(json!({
                "ok": false, "account": t.label, "status": "agenda_failed",
                "upstream_http_status": res.status, "upstream_error": res.error,
                "returned_count": 0, "items": [],
            }));
            continue;
        }
        let pascal = res.api != "graph";
        let items: Vec<Value> = values(&res.body)
            .iter()
            .map(|e| compact_event(e, &t.label, pascal))
            .collect();
        lists.push(json!({ "ok": true, "account": t.label, "status": "ok", "returned_count": items.len(), "items": items }));
    }
    // A per-account list: calendar events are not merged across accounts.
    let accounts: Vec<Value> = lists
        .iter()
        .map(|l| {
            json!({
                "account": l["account"], "ok": l["ok"], "status": l["status"],
                "returned_count": l["returned_count"], "items": l["items"],
            })
        })
        .collect();
    let total: i64 = accounts
        .iter()
        .filter(|a| a["ok"] == json!(true))
        .map(|a| a["returned_count"].as_i64().unwrap_or(0))
        .sum();
    Ok(json!({
        "ok": true,
        "status": "ok",
        "scope": account_scope_label(account),
        "start_date": start,
        "end_date": end,
        "accounts": accounts,
        "answer_text": if total == 0 { "No events in the selected range.".to_string() } else { format!("{total} event{} in the selected range.", plural(total as usize)) },
    }))
}

// ---------------------------------------------------------------------------
// Attachments
// ---------------------------------------------------------------------------

/// Graph (camelCase) and Outlook REST (PascalCase) attachments -> one compact shape.
/// `data_base64` is kept internally and stripped before output.
fn compact_attachment(raw: &Value, index: usize) -> Map<String, Value> {
    let pick = |name: &str| {
        raw.get(name)
            .filter(|v| !v.is_null())
            .or_else(|| raw.get(cap(true, name)))
    };
    let odata = js::get_str(raw, &["@odata.type"]);
    let ty = odata.rsplit('.').next().unwrap_or("").to_lowercase();
    let kind = if ty.contains("file") {
        "file".to_string()
    } else if ty.contains("item") {
        "item".to_string()
    } else if ty.contains("reference") {
        "reference".to_string()
    } else if ty.is_empty() {
        "file".to_string()
    } else {
        ty
    };
    let content_type = js::str_or_empty(pick("contentType")).to_lowercase();
    let name = js::str_or_empty(pick("name"));
    let cid = js::str_or_empty(pick("contentId"));
    let cid = cid.strip_prefix('<').unwrap_or(&cid);
    let cid = cid.strip_suffix('>').unwrap_or(cid);
    let size = match pick("size").filter(|v| js::truthy(v)) {
        Some(v) => js::num(js::number(Some(v))),
        None => json!(0),
    };
    let mut m = Map::new();
    m.insert("index".into(), json!(index));
    m.insert("attachment_id".into(), json!(js::str_or_empty(pick("id"))));
    m.insert(
        "name".into(),
        json!(if name.is_empty() {
            format!("attachment-{index}")
        } else {
            name
        }),
    );
    m.insert("content_type".into(), json!(content_type));
    m.insert("size".into(), size);
    m.insert(
        "is_inline".into(),
        json!(pick("isInline") == Some(&json!(true))),
    );
    m.insert("content_id".into(), json!(cid));
    m.insert("kind".into(), json!(kind));
    m.insert(
        "is_image".into(),
        json!(kind == "file" && content_type.starts_with("image/")),
    );
    m.insert(
        "data_base64".into(),
        json!(if kind == "file" {
            js::str_or_empty(pick("contentBytes"))
        } else {
            String::new()
        }),
    );
    m
}

fn without_data(att: &Map<String, Value>) -> Map<String, Value> {
    let mut m = att.clone();
    m.shift_remove("data_base64");
    m
}

fn b64_decode(s: &str) -> Vec<u8> {
    let cleaned: String = s
        .chars()
        .filter(|c| !c.is_whitespace() && *c != '=')
        .map(|c| match c {
            '-' => '+',
            '_' => '/',
            c => c,
        })
        .collect();
    base64::engine::general_purpose::STANDARD_NO_PAD
        .decode(cleaned.as_bytes())
        .unwrap_or_default()
}

fn b64_encode(b: &[u8]) -> String {
    base64::engine::general_purpose::STANDARD.encode(b)
}

/// `{ok: true, account, attachments}` or an error result.
fn fetch_attachments(
    ctx: &Ctx,
    account: Option<&Value>,
    message_id: &str,
    attachment_id: &str,
) -> Result<Value> {
    let key = json!([
        normalize_string(account).to_lowercase(),
        message_id,
        attachment_id
    ])
    .to_string();
    {
        let cache = ctx.attachment_cache.borrow();
        if let Some((_, at, result)) = cache.iter().find(|(k, _, _)| *k == key) {
            if at.elapsed() < ATTACHMENT_CACHE_TTL {
                return Ok(result.clone());
            }
        }
    }
    let result = fetch_attachments_uncached(ctx, account, message_id, attachment_id)?;
    if result["ok"] == json!(true) {
        let mut cache = ctx.attachment_cache.borrow_mut();
        cache.retain(|(k, _, _)| *k != key);
        cache.push((key, Instant::now(), result.clone()));
        while cache.len() > ATTACHMENT_CACHE_MAX_ENTRIES {
            cache.remove(0);
        }
    }
    Ok(result)
}

fn fetch_attachments_uncached(
    ctx: &Ctx,
    account: Option<&Value>,
    message_id: &str,
    attachment_id: &str,
) -> Result<Value> {
    let targets = resolve_accounts(ctx, account)?;
    for t in &targets {
        let Some(auth) = safe_get_best_token(ctx, t, "mail")? else {
            continue;
        };
        let one = if attachment_id.is_empty() {
            String::new()
        } else {
            format!("/{}", enc(attachment_id))
        };
        let res = with_fallback(&first_api(&auth), |api| {
            if api == "graph" {
                let url = api_url(
                    &graph_base(ctx),
                    &format!("/v1.0/me/messages/{}/attachments{one}", enc(message_id)),
                    &[],
                )?;
                fetch_json(
                    ctx,
                    url,
                    &graph_headers(&auth.token, false, None),
                    ATTACHMENT_TIMEOUT,
                    "graph",
                )
            } else {
                let url = api_url(
                    &rest_base(ctx),
                    &format!("/api/v2.0/me/messages/{}/attachments{one}", enc(message_id)),
                    &[],
                )?;
                fetch_json(
                    ctx,
                    url,
                    &bearer_headers(&auth.token),
                    ATTACHMENT_TIMEOUT,
                    "outlook_rest",
                )
            }
        })?;
        if !res.ok {
            if targets.len() > 1 && res.status == 404 {
                continue;
            }
            return Ok(json!({
                "ok": false, "status": "attachments_failed", "account": t.label, "id": message_id,
                "upstream_http_status": res.status, "answer_text": "The Outlook upstream call failed.",
            }));
        }
        let raws: Vec<Value> = if attachment_id.is_empty() {
            values(&res.body)
        } else {
            res.body.clone().into_iter().collect()
        };
        let atts: Vec<Value> = raws
            .iter()
            .enumerate()
            .map(|(i, r)| Value::Object(compact_attachment(r, i)))
            .collect();
        return Ok(json!({ "ok": true, "account": t.label, "attachments": atts }));
    }
    Ok(json!({
        "ok": false, "status": "message_not_found", "scope": account_scope_label(account),
        "id": message_id, "answer_text": "Message not found.",
    }))
}

fn strip_angle(s: &str) -> String {
    let s = s.strip_prefix('<').unwrap_or(s);
    s.strip_suffix('>').unwrap_or(s).to_string()
}

fn select_attachment<'a>(
    atts: &'a [Value],
    index: Option<&Value>,
    content_id: Option<&Value>,
    attachment_id: Option<&Value>,
) -> Option<&'a Value> {
    let aid = normalize_string(attachment_id);
    if !aid.is_empty() {
        return atts
            .iter()
            .find(|x| js::get_str(x, &["attachment_id"]) == aid)
            .or_else(|| (atts.len() == 1).then(|| &atts[0]));
    }
    let cid = strip_angle(&normalize_string(content_id)).to_lowercase();
    if !cid.is_empty() {
        return atts
            .iter()
            .find(|x| js::get_str(x, &["content_id"]).to_lowercase() == cid);
    }
    if js::given(index) {
        let n = js::number(index);
        return atts.iter().find(|x| x["index"].as_f64() == Some(n));
    }
    atts.iter().find(|x| x["is_image"] == json!(true))
}

pub struct AttachmentArgs<'a> {
    pub account: Option<&'a Value>,
    pub id: Option<&'a Value>,
    pub include_data: Option<&'a Value>,
    pub max_data_bytes: Option<&'a Value>,
    pub images_only: Option<&'a Value>,
    pub index: Option<&'a Value>,
    pub content_id: Option<&'a Value>,
    pub attachment_id: Option<&'a Value>,
    pub offset: Option<&'a Value>,
    pub length: Option<&'a Value>,
}

/// Lists a message's attachments, including inline images pasted into the body (the
/// cid: images Outlook calls image001.png). Pick one with index, content_id or
/// attachment_id; offset/length return that byte range of it. GET requests only.
pub fn outlook_attachments(ctx: &Ctx, a: AttachmentArgs) -> Result<Value> {
    let message_id = normalize_string(a.id);
    if message_id.is_empty() {
        return Ok(missing_field("id", "An Outlook message id is required."));
    }
    let listed = fetch_attachments(
        ctx,
        a.account,
        &message_id,
        &normalize_string(a.attachment_id),
    )?;
    if listed["ok"] != json!(true) {
        return Ok(listed);
    }
    let account = listed["account"].clone();
    let atts: Vec<Value> = listed["attachments"]
        .as_array()
        .cloned()
        .unwrap_or_default();
    let cap = js::clamp_integer(
        a.max_data_bytes,
        1_000,
        MAX_ATTACHMENT_DATA_BYTES,
        DEFAULT_ATTACHMENT_DATA_BYTES,
    ) as usize;
    let wants_range = js::given(a.offset);
    let pick_one = wants_range
        || !normalize_string(a.attachment_id).is_empty()
        || !normalize_string(a.content_id).is_empty()
        || js::given(a.index);
    if pick_one {
        let Some(att) = select_attachment(&atts, a.index, a.content_id, a.attachment_id) else {
            return Ok(json!({
                "ok": false, "status": "attachment_not_found", "account": account, "id": message_id,
                "answer_text": "No attachment matched that index, content id or attachment id.",
            }));
        };
        let att = att.as_object().cloned().unwrap_or_default();
        let data = js::str_or_empty(att.get("data_base64"));
        let item = without_data(&att);
        let name = js::str_or_empty(item.get("name"));
        let bytes = b64_decode(&data);
        let total = bytes.len();
        if wants_range {
            let start = js::clamp_integer(a.offset, 0, total as i64, 0) as usize;
            let len = js::clamp_integer(
                a.length,
                1_000,
                MAX_ATTACHMENT_DATA_BYTES,
                DEFAULT_ATTACHMENT_DATA_BYTES,
            ) as usize;
            let end = total.min(start + len);
            return Ok(json!({
                "ok": true, "status": "ok", "account": account, "id": message_id, "attachment": item,
                "total_bytes": total, "offset": start, "returned_bytes": end - start,
                "next_offset": if end < total { json!(end) } else { Value::Null },
                "data_base64": b64_encode(&bytes[start..end]),
                "answer_text": format!("Bytes {start}-{end} of {total} of {name}."),
            }));
        }
        let fits = total <= cap;
        let mut out = Map::new();
        for (k, v) in [
            ("ok", json!(true)),
            ("status", json!("ok")),
            ("account", account.clone()),
            ("id", json!(message_id)),
            ("attachment", Value::Object(item)),
            ("total_bytes", json!(total)),
        ] {
            out.insert(k.into(), v);
        }
        if js::to_boolean(a.include_data, true) && fits {
            out.insert("data_base64".into(), json!(data));
        }
        out.insert(
            "answer_text".into(),
            json!(if fits {
                format!("{name}, {total} bytes.")
            } else {
                format!("{name} is {total} bytes; fetch it in pieces with offset and length.")
            }),
        );
        return Ok(Value::Object(out));
    }
    let with_data = js::to_boolean(a.include_data, false);
    let images_only = js::to_boolean(a.images_only, false);
    let mut items = Vec::new();
    for att in &atts {
        if images_only && att["is_image"] != json!(true) {
            continue;
        }
        let att = att.as_object().cloned().unwrap_or_default();
        let data = js::str_or_empty(att.get("data_base64"));
        let mut item = without_data(&att);
        if with_data && !data.is_empty() && b64_decode(&data).len() <= cap {
            item.insert("data_base64".into(), json!(data));
        }
        items.push(Value::Object(item));
    }
    let images = items
        .iter()
        .filter(|x| x["is_image"] == json!(true))
        .count();
    let answer = if items.is_empty() {
        "That email has no attachments.".to_string()
    } else {
        let names: Vec<String> = items
            .iter()
            .take(5)
            .map(|x| js::get_str(x, &["name"]))
            .collect();
        let imgs = if images > 0 {
            format!(", {images} image{}", plural(images))
        } else {
            String::new()
        };
        format!(
            "{} attachment{}{imgs}: {}.",
            items.len(),
            plural(items.len()),
            names.join("; ")
        )
    };
    Ok(json!({
        "ok": true, "status": "ok", "account": account, "id": message_id,
        "returned_count": items.len(), "image_count": images, "attachments": items,
        "answer_text": answer,
    }))
}

/// Sends one image attachment (default: the first image) to the vision model and
/// returns a description plus the full transcribed text.
pub fn outlook_attachment_inspect(
    ctx: &Ctx,
    account: Option<&Value>,
    id: Option<&Value>,
    index: Option<&Value>,
    content_id: Option<&Value>,
    prompt: Option<&Value>,
) -> Result<Value> {
    let message_id = normalize_string(id);
    if message_id.is_empty() {
        return Ok(missing_field("id", "An Outlook message id is required."));
    }
    let listed = fetch_attachments(ctx, account, &message_id, "")?;
    if listed["ok"] != json!(true) {
        return Ok(listed);
    }
    let acct = listed["account"].clone();
    let atts: Vec<Value> = listed["attachments"]
        .as_array()
        .cloned()
        .unwrap_or_default();
    let has_data = |x: &Value| {
        x["is_image"] == json!(true) && js::get(x, &["data_base64"]).is_some_and(js::truthy)
    };
    let images: Vec<Value> = atts
        .iter()
        .filter(|x| has_data(x))
        .map(|x| Value::Object(without_data(x.as_object().unwrap())))
        .collect();
    let image = select_attachment(&atts, index, content_id, None).filter(|x| has_data(x));
    let Some(image) = image else {
        let any = !images.is_empty();
        return Ok(json!({
            "ok": false,
            "status": if any { "image_not_found" } else { "no_images" },
            "account": acct, "id": message_id, "images": images,
            "answer_text": if any { "No image matched that index or content id." } else { "That email has no images." },
        }));
    };
    if js::number(image.get("size")) > MAX_INSPECT_IMAGE_BYTES {
        return Ok(json!({
            "ok": false, "status": "image_too_large", "account": acct, "id": message_id,
            "answer_text": "That image is too large to inspect.",
        }));
    }
    let Some(analyze) = ctx.vision.as_ref() else {
        let message = "No image model is set up. Set OPENROUTER_API_KEY.";
        return Ok(json!({
            "ok": false, "status": "vision_not_configured", "account": acct, "id": message_id,
            "message": message, "answer_text": message,
        }));
    };
    let p = normalize_string(prompt);
    let vision = analyze(&json!({
        "provider": "openrouter",
        "prompt": if p.is_empty() { DEFAULT_ATTACHMENT_INSPECT_PROMPT.to_string() } else { p },
        "images": [{
            "media_type": image["content_type"],
            "data_base64": image["data_base64"],
            "filename": image["name"],
        }],
    }))?;
    let meta = Value::Object(without_data(image.as_object().unwrap()));
    let s = |k: &str| js::str_or_empty(vision.get(k));
    if vision.get("ok") != Some(&json!(true)) {
        let mut out = Map::new();
        out.insert("ok".into(), json!(false));
        out.insert(
            "status".into(),
            json!(if s("status").is_empty() {
                "vision_error".to_string()
            } else {
                s("status")
            }),
        );
        out.insert("account".into(), acct);
        out.insert("id".into(), json!(message_id));
        out.insert("image".into(), meta);
        if let Some(m) = vision.get("message").filter(|m| !m.is_null()) {
            out.insert("message".into(), m.clone());
        }
        out.insert(
            "answer_text".into(),
            json!(if s("message").is_empty() {
                "Vision analysis failed.".to_string()
            } else {
                s("message")
            }),
        );
        return Ok(Value::Object(out));
    }
    let answer = [s("answer_text"), s("description")]
        .into_iter()
        .find(|x| !x.is_empty())
        .unwrap_or_else(|| "I inspected the image.".into());
    Ok(json!({
        "ok": true, "status": "ok", "account": acct, "id": message_id, "image": meta,
        "provider": s("provider"),
        "upstream_provider": s("upstream_provider"),
        "model": s("model"),
        "usage": vision.get("usage").filter(|u| js::truthy(u)).cloned().unwrap_or(Value::Null),
        "description": s("description"),
        "ocr_text": s("ocr_text"),
        "answer_text": answer,
    }))
}

/// Weekday of a date, 0 = Sunday.
pub(crate) fn weekday_sun0(d: NaiveDate) -> u32 {
    d.weekday().num_days_from_sunday()
}
