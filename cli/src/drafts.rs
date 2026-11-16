//! Drafts. Mirrors Outlook web: Reply / Reply all create a draft in the thread with
//! the quoted original below your text; New message saves to Drafts. These only
//! create, update or discard drafts. There is deliberately no send path: every write
//! goes through `rest_write`, which refuses any `/send` path.

use crate::accounts::{get_account, resolve_account_alias, Account};
use crate::ctx::Ctx;
use crate::js::{self, normalize_string};
use crate::outlook::{
    account_names, api_url, enc, missing_field, rest_base, safe_get_best_token, send, Auth, Resp,
    DEFAULT_TIMEOUT,
};
use anyhow::{bail, Result};
use base64::Engine;
use regex::Regex;
use serde_json::{json, Map, Value};
use std::collections::HashSet;
use std::sync::LazyLock;

const MAX_INLINE_IMAGE_BYTES: usize = 3 * 1024 * 1024;

static SEND_RE: LazyLock<Regex> = LazyLock::new(|| Regex::new(r"(?i)/send(\?|$)").unwrap());
static BODY_TAG_RE: LazyLock<Regex> = LazyLock::new(|| Regex::new(r"(?i)<body[^>]*>").unwrap());
static TAG_RE: LazyLock<Regex> = LazyLock::new(|| Regex::new(r"<[^>]+>").unwrap());
static PARAS_RE: LazyLock<Regex> = LazyLock::new(|| Regex::new(r"\n{2,}").unwrap());
static NAMED_ADDR_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"^(.*?)<\s*([^>]+)\s*>$").unwrap());

/// The only way this program writes to a mailbox. Refuses anything that would send.
pub fn rest_write(
    ctx: &Ctx,
    token: &str,
    method: &str,
    path: &str,
    payload: Option<&Value>,
) -> Result<Resp> {
    if SEND_RE.is_match(path) {
        bail!("Refusing to send mail: drafts only.");
    }
    let url = api_url(&rest_base(ctx), &format!("/api/v2.0/{path}"), &[])?;
    let method = reqwest::Method::from_bytes(method.as_bytes())?;
    let mut req = ctx
        .http
        .request(method, url)
        .header("authorization", format!("Bearer {token}"))
        .header("accept", "application/json")
        .header("content-type", "application/json");
    if let Some(p) = payload {
        req = req.body(p.to_string());
    }
    send(req, DEFAULT_TIMEOUT, "outlook_rest")
}

fn single_work_account(ctx: &Ctx, account: Option<&Value>) -> Option<Account> {
    resolve_account_alias(ctx, account).and_then(|k| get_account(ctx, &k))
}

fn pick_one_account(ctx: &Ctx) -> Value {
    let names = account_names(ctx);
    let list = if names.is_empty() {
        "none configured".to_string()
    } else {
        names.join(", ")
    };
    missing_field("account", &format!("Pick one account: {list}."))
}

fn draft_token(ctx: &Ctx, target: &Account) -> Option<Auth> {
    safe_get_best_token(ctx, target, "mail_write")
        .ok()
        .flatten()
}

fn confirmation_required(message: String) -> Value {
    json!({ "ok": false, "status": "confirmation_required", "message": message, "answer_text": message })
}

fn no_draft_token(target: &Account) -> Value {
    json!({
        "ok": false,
        "status": "needs_sign_in",
        "account": target.label,
        "answer_text": format!("The {} Outlook session has no token that can save drafts. Run outlook-web-cli auth status, then sign in again (outlook-web-cli auth chrome or auth cookies).", target.label),
    })
}

fn upstream_error(target: &Account, res: &Resp, status: &str) -> Value {
    json!({
        "ok": false,
        "status": status,
        "account": target.label,
        "upstream_http_status": res.status,
        "upstream_error": res.error.chars().take(300).collect::<String>(),
        "answer_text": "Outlook didn't save the draft.",
    })
}

fn recipients_out(list: Option<&Value>) -> Vec<Value> {
    list.and_then(Value::as_array)
        .into_iter()
        .flatten()
        .map(|r| {
            json!({
                "name": js::get_str(r, &["EmailAddress", "Name"]),
                "address": js::get_str(r, &["EmailAddress", "Address"]),
            })
        })
        .collect()
}

fn draft_result(target: &Account, msg: &Value, answer: &str) -> Value {
    json!({
        "ok": true,
        "status": "draft_saved",
        "account": target.label,
        "id": js::get_str(msg, &["Id"]),
        "subject": js::get_str(msg, &["Subject"]),
        "to": recipients_out(msg.get("ToRecipients")),
        "cc": recipients_out(msg.get("CcRecipients")),
        "is_draft": msg.get("IsDraft") != Some(&json!(false)),
        "web_link": js::get_str(msg, &["WebLink"]),
        "answer_text": format!("{answer} It's in {} Outlook Drafts and has not been sent.", target.label),
    })
}

fn draft_content(body: Option<&Value>, body_html: Option<&Value>) -> Option<Value> {
    let html = normalize_string(body_html);
    if !html.is_empty() {
        return Some(json!({ "ContentType": "HTML", "Content": html }));
    }
    let text = normalize_string(body);
    if text.is_empty() {
        return None;
    }
    Some(json!({ "ContentType": "HTML", "Content": text_to_html(&text) }))
}

fn text_to_html(text: &str) -> String {
    let esc = text
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;");
    PARAS_RE
        .split(&esc)
        .map(|p| format!("<p>{}</p>", p.replace('\n', "<br>")))
        .collect()
}

/// Puts the new text above the quoted original, where Outlook web's cursor starts.
fn prepend_to_body(existing: &Value, content: &Value) -> Value {
    let current = js::get_str(existing, &["Content"]);
    let new = js::get_str(content, &["Content"]);
    if js::get_str(existing, &["ContentType"]).to_lowercase() == "text" {
        let plain = new.replace("<br>", "\n").replace("</p><p>", "\n\n");
        let plain = TAG_RE
            .replace_all(&plain, "")
            .replace("&lt;", "<")
            .replace("&gt;", ">")
            .replace("&amp;", "&");
        return json!({ "ContentType": "Text", "Content": format!("{plain}\n\n{current}") });
    }
    let block = format!("<div>{new}</div><br>");
    match BODY_TAG_RE.find(&current) {
        None => json!({ "ContentType": "HTML", "Content": format!("{block}{current}") }),
        Some(m) => json!({
            "ContentType": "HTML",
            "Content": format!("{}{block}{}", &current[..m.end()], &current[m.end()..]),
        }),
    }
}

/// Addresses from "a@x.com, Name <b@x.com>; ..." or a list of strings or
/// `{address|email, name}` objects, as Outlook REST recipients.
pub fn parse_recipients(value: Option<&Value>) -> Vec<Value> {
    let items: Vec<Value> = match value {
        Some(Value::Array(a)) => a.clone(),
        other => {
            let s = match other {
                Some(v) if js::truthy(v) => js::string(v),
                _ => String::new(),
            };
            s.split([',', ';'])
                .map(|x| Value::String(x.to_string()))
                .collect()
        }
    };
    items
        .iter()
        .filter_map(|v| match v {
            Value::Object(o) => {
                let address = normalize_string(
                    o.get("address")
                        .filter(|x| js::truthy(x))
                        .or_else(|| o.get("email")),
                );
                if address.is_empty() {
                    return None;
                }
                let name = normalize_string(o.get("name"));
                let mut ea = Map::new();
                ea.insert("Address".into(), json!(address));
                if !name.is_empty() {
                    ea.insert("Name".into(), json!(name));
                }
                Some(json!({ "EmailAddress": ea }))
            }
            other => {
                let s = js::string(other).trim().to_string();
                if s.is_empty() {
                    return None;
                }
                let (address, name) = match NAMED_ADDR_RE.captures(&s) {
                    Some(c) => (
                        c[2].trim().to_string(),
                        c[1].trim()
                            .trim_start_matches('"')
                            .trim_end_matches('"')
                            .trim()
                            .to_string(),
                    ),
                    None => (s.clone(), String::new()),
                };
                if !address.contains('@') {
                    return None;
                }
                let mut ea = Map::new();
                ea.insert("Address".into(), json!(address));
                if !name.is_empty() {
                    ea.insert("Name".into(), json!(name));
                }
                Some(json!({ "EmailAddress": ea }))
            }
        })
        .collect()
}

fn inline_image_hosts(ctx: &Ctx) -> Vec<String> {
    let mut hosts: Vec<String> = Vec::new();
    for h in ctx
        .var("OUTLOOK_INLINE_IMAGE_HOSTS")
        .unwrap_or("")
        .split(',')
    {
        let h = h.trim();
        if !h.is_empty() && !hosts.iter().any(|x| x == h) {
            hosts.push(h.to_string());
        }
    }
    hosts
}

/// Inline images for a draft body that references them as `<img src="cid:...">`.
/// Each image is fetched from an allowed host (OUTLOOK_INLINE_IMAGE_HOSTS) and attached
/// with IsInline, so it shows without "download pictures". No host is allowed until one
/// is set, so a draft can't be made to fetch arbitrary URLs.
fn attach_inline_images(
    ctx: &Ctx,
    token: &str,
    draft_id: &str,
    images: Option<&Value>,
) -> Result<std::result::Result<usize, String>> {
    let list = images
        .and_then(Value::as_array)
        .cloned()
        .unwrap_or_default();
    let allowed = inline_image_hosts(ctx);
    for img in &list {
        let content_id = normalize_string(
            img.get("content_id")
                .filter(|v| js::truthy(v))
                .or_else(|| img.get("contentId")),
        );
        let Ok(url) = url::Url::parse(&js::str_or_empty(img.get("url"))) else {
            return Ok(Err("Each inline image needs a valid url.".into()));
        };
        let host = url.host_str().unwrap_or("").to_string();
        let local = host == "127.0.0.1" || host == "localhost";
        if content_id.is_empty() || !allowed.contains(&host) || (url.scheme() != "https" && !local)
        {
            return Ok(Err(if allowed.is_empty() {
                "Inline images are off: set OUTLOOK_INLINE_IMAGE_HOSTS to the image hosts to allow."
                    .into()
            } else {
                format!(
                    "Inline image must have a content_id and an https url on {}.",
                    allowed.join(", ")
                )
            }));
        }
        let path = url.path().to_string();
        let fetched = ctx.http.get(url.clone()).timeout(DEFAULT_TIMEOUT).send();
        let res = match fetched {
            Ok(r) if r.status().is_success() => r,
            _ => return Ok(Err(format!("Couldn't fetch image {path}."))),
        };
        let content_type = res
            .headers()
            .get("content-type")
            .and_then(|v| v.to_str().ok())
            .unwrap_or("")
            .to_string();
        if !content_type.starts_with("image/") {
            return Ok(Err(format!("Couldn't fetch image {path}.")));
        }
        let Ok(bytes) = res.bytes() else {
            return Ok(Err(format!("Couldn't fetch image {path}.")));
        };
        if bytes.len() > MAX_INLINE_IMAGE_BYTES {
            return Ok(Err(format!("Image {path} is over 3 MB.")));
        }
        let name = normalize_string(img.get("name"));
        let name = if name.is_empty() {
            path.rsplit('/').next().unwrap_or("").to_string()
        } else {
            name
        };
        let saved = rest_write(
            ctx,
            token,
            "POST",
            &format!("me/messages/{}/attachments", enc(draft_id)),
            Some(&json!({
                "@odata.type": "#Microsoft.OutlookServices.FileAttachment",
                "Name": name,
                "ContentType": content_type,
                "ContentBytes": base64::engine::general_purpose::STANDARD.encode(&bytes),
                "ContentId": content_id,
                "IsInline": true,
            })),
        )?;
        if !saved.ok {
            return Ok(Err(if saved.error.is_empty() {
                "Outlook rejected the image.".into()
            } else {
                saved.error
            }));
        }
    }
    Ok(Ok(list.len()))
}

fn with_image_failure(mut result: Value, error: String) -> Value {
    result["ok"] = json!(false);
    result["status"] = json!("inline_image_failed");
    result["error"] = json!(error);
    result
}

pub struct ReplyArgs<'a> {
    pub account: Option<&'a Value>,
    pub id: Option<&'a Value>,
    pub body: Option<&'a Value>,
    pub body_html: Option<&'a Value>,
    pub reply_all: Option<&'a Value>,
    pub add_to: Option<&'a Value>,
    pub add_cc: Option<&'a Value>,
    pub inline_images: Option<&'a Value>,
    pub confirmed: bool,
}

pub fn outlook_draft_reply(ctx: &Ctx, a: ReplyArgs) -> Result<Value> {
    let message_id = normalize_string(a.id);
    if message_id.is_empty() {
        return Ok(missing_field("id", "An Outlook message id is required."));
    }
    let Some(content) = draft_content(a.body, a.body_html) else {
        return Ok(missing_field("body", "Draft text is required."));
    };
    let Some(target) = single_work_account(ctx, a.account) else {
        return Ok(pick_one_account(ctx));
    };
    let all = js::to_boolean(a.reply_all, false);
    if !a.confirmed {
        return Ok(confirmation_required(format!(
            "Confirm that Andrew wants to save a {} draft in {} Outlook.",
            if all { "reply-all" } else { "reply" },
            target.label
        )));
    }
    let Some(auth) = draft_token(ctx, &target) else {
        return Ok(no_draft_token(&target));
    };
    let created = rest_write(
        ctx,
        &auth.token,
        "POST",
        &format!(
            "me/messages/{}/{}",
            enc(&message_id),
            if all { "createreplyall" } else { "createreply" }
        ),
        Some(&json!({})),
    )?;
    if !created.ok {
        return Ok(upstream_error(&target, &created, "reply_draft_failed"));
    }
    let draft = created.body.clone().unwrap_or_else(|| json!({}));
    let merged = prepend_to_body(draft.get("Body").unwrap_or(&json!({})), &content);
    // Optional extra recipients (e.g. the rest of the team on a reply-all); anyone
    // already on the draft is kept once.
    let mut patch = Map::new();
    patch.insert("Body".into(), merged);
    let extra_to = parse_recipients(a.add_to);
    let extra_cc = parse_recipients(a.add_cc);
    if !extra_to.is_empty() || !extra_cc.is_empty() {
        let mut seen = HashSet::new();
        let mut keep = |list: Vec<Value>| -> Vec<Value> {
            list.into_iter()
                .filter(|r| {
                    let key = js::get_str(r, &["EmailAddress", "Address"]).to_lowercase();
                    !key.is_empty() && seen.insert(key)
                })
                .collect()
        };
        let existing = |k: &str| {
            draft
                .get(k)
                .and_then(Value::as_array)
                .cloned()
                .unwrap_or_default()
        };
        let mut to = existing("ToRecipients");
        to.extend(extra_to);
        let mut cc = existing("CcRecipients");
        cc.extend(extra_cc);
        patch.insert("ToRecipients".into(), Value::Array(keep(to)));
        patch.insert("CcRecipients".into(), Value::Array(keep(cc)));
    }
    let draft_id = js::str_or_empty(draft.get("Id"));
    let draft_id = if draft_id.is_empty() {
        "undefined".to_string()
    } else {
        draft_id
    };
    let patched = rest_write(
        ctx,
        &auth.token,
        "PATCH",
        &format!("me/messages/{}", enc(&draft_id)),
        Some(&Value::Object(patch)),
    )?;
    if !patched.ok {
        return Ok(upstream_error(&target, &patched, "reply_draft_failed"));
    }
    let saved = patched
        .body
        .clone()
        .filter(|b| !b.is_null())
        .unwrap_or_else(|| draft.clone());
    let label = if all {
        "Reply-all draft saved"
    } else {
        "Reply draft saved"
    };
    match attach_inline_images(ctx, &auth.token, &draft_id, a.inline_images)? {
        Err(e) => Ok(with_image_failure(
            draft_result(
                &target,
                &saved,
                &format!("{label}, but an image couldn't be attached."),
            ),
            e,
        )),
        Ok(0) => Ok(draft_result(&target, &saved, &format!("{label}."))),
        Ok(n) => Ok(draft_result(
            &target,
            &saved,
            &format!("{label} with {n} inline image(s)."),
        )),
    }
}

pub struct NewArgs<'a> {
    pub account: Option<&'a Value>,
    pub to: Option<&'a Value>,
    pub cc: Option<&'a Value>,
    pub subject: Option<&'a Value>,
    pub body: Option<&'a Value>,
    pub body_html: Option<&'a Value>,
    pub inline_images: Option<&'a Value>,
    pub confirmed: bool,
}

pub fn outlook_draft_new(ctx: &Ctx, a: NewArgs) -> Result<Value> {
    let to = parse_recipients(a.to);
    if to.is_empty() {
        return Ok(missing_field("to", "At least one recipient is required."));
    }
    let Some(content) = draft_content(a.body, a.body_html) else {
        return Ok(missing_field("body", "Draft text is required."));
    };
    let Some(target) = single_work_account(ctx, a.account) else {
        return Ok(pick_one_account(ctx));
    };
    if !a.confirmed {
        return Ok(confirmation_required(format!(
            "Confirm that Andrew wants to save a new draft in {} Outlook.",
            target.label
        )));
    }
    let Some(auth) = draft_token(ctx, &target) else {
        return Ok(no_draft_token(&target));
    };
    // POST /me/messages saves the message to Drafts without sending it.
    let created = rest_write(
        ctx,
        &auth.token,
        "POST",
        "me/messages",
        Some(&json!({
            "Subject": normalize_string(a.subject),
            "Body": content,
            "ToRecipients": to,
            "CcRecipients": parse_recipients(a.cc),
        })),
    )?;
    if !created.ok {
        return Ok(upstream_error(&target, &created, "new_draft_failed"));
    }
    let draft = created
        .body
        .clone()
        .filter(|b| !b.is_null())
        .unwrap_or_else(|| json!({}));
    let id = js::str_or_empty(draft.get("Id"));
    let id = if id.is_empty() {
        "undefined".to_string()
    } else {
        id
    };
    match attach_inline_images(ctx, &auth.token, &id, a.inline_images)? {
        Err(e) => Ok(with_image_failure(
            draft_result(
                &target,
                &draft,
                "Draft saved, but an image couldn't be attached.",
            ),
            e,
        )),
        Ok(0) => Ok(draft_result(&target, &draft, "Draft saved.")),
        Ok(n) => Ok(draft_result(
            &target,
            &draft,
            &format!("Draft saved with {n} inline image(s)."),
        )),
    }
}

/// Moves a draft to Deleted Items (recoverable). Anything that isn't a draft is refused.
pub fn outlook_delete_draft(
    ctx: &Ctx,
    account: Option<&Value>,
    id: Option<&Value>,
    confirmed: bool,
) -> Result<Value> {
    let message_id = normalize_string(id);
    if message_id.is_empty() {
        return Ok(missing_field("id", "An Outlook draft id is required."));
    }
    let Some(target) = single_work_account(ctx, account) else {
        return Ok(pick_one_account(ctx));
    };
    if !confirmed {
        return Ok(confirmation_required(format!(
            "Confirm that Andrew wants to delete a draft in {} Outlook.",
            target.label
        )));
    }
    let Some(auth) = draft_token(ctx, &target) else {
        return Ok(no_draft_token(&target));
    };
    let path = format!("me/messages/{}", enc(&message_id));
    let found = rest_write(
        ctx,
        &auth.token,
        "GET",
        &format!("{path}?$select=Subject,IsDraft"),
        None,
    )?;
    if !found.ok {
        return Ok(upstream_error(&target, &found, "delete_draft_failed"));
    }
    let body = found.body.clone().unwrap_or(Value::Null);
    if body.get("IsDraft") != Some(&json!(true)) {
        return Ok(json!({
            "ok": false,
            "status": "not_a_draft",
            "account": target.label,
            "answer_text": "That message isn't a draft, so it was left alone.",
        }));
    }
    let moved = rest_write(
        ctx,
        &auth.token,
        "POST",
        &format!("{path}/move"),
        Some(&json!({ "DestinationId": "DeletedItems" })),
    )?;
    if !moved.ok {
        return Ok(upstream_error(&target, &moved, "delete_draft_failed"));
    }
    Ok(json!({
        "ok": true,
        "status": "draft_deleted",
        "account": target.label,
        "subject": js::get_str(&body, &["Subject"]),
        "answer_text": format!("Draft moved to {} Outlook Deleted Items.", target.label),
    }))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn send_paths_are_refused() {
        let ctx = Ctx::with_vars([("OUTLOOK_REST_BASE_URL", "http://127.0.0.1:9")]);
        for p in [
            "me/messages/X/send",
            "me/messages/X/SEND",
            "me/messages/X/send?y=1",
        ] {
            let r = rest_write(&ctx, "t", "POST", p, Some(&json!({})));
            assert!(
                r.unwrap_err().to_string().contains("Refusing to send"),
                "{p}"
            );
        }
    }

    #[test]
    fn recipients_parse() {
        let r = parse_recipients(Some(&json!(
            "Oscar Buitrago <oscar@contoso.com>, dia@contoso.com; nope"
        )));
        assert_eq!(
            r,
            vec![
                json!({"EmailAddress": {"Address": "oscar@contoso.com", "Name": "Oscar Buitrago"}}),
                json!({"EmailAddress": {"Address": "dia@contoso.com"}}),
            ]
        );
    }
}
