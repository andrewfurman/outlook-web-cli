//! Accounts are named mailboxes, e.g. "work". Each one may list aliases (other names
//! or email domains that pick it), the signed-in username, and, optionally, a Chrome
//! DevTools port where Outlook web is open, used to import or re-import the web session.
//!
//! They come from `$OUTLOOK_WEB_CONFIG_DIR/config.json`
//! (default `~/.config/outlook-web-cli/config.json`):
//!
//! ```json
//! { "accounts": [ { "name": "work", "aliases": ["contoso.com"], "username": "me@contoso.com",
//!                   "cdp_port": 9230, "mailbox_url": "https://outlook.cloud.microsoft/mail/" } ] }
//! ```
//!
//! Saved sessions alone also count as accounts, so `outlook-web-cli auth chrome --account work`
//! is enough to start without writing a config file.

use crate::ctx::Ctx;
use crate::js;
use regex::Regex;
use serde_json::Value;
use std::sync::LazyLock;

pub const DEFAULT_MAILBOX_URL: &str = "https://outlook.cloud.microsoft/mail/";

pub static NAME_RE: LazyLock<Regex> =
    LazyLock::new(|| Regex::new(r"^[a-z0-9][a-z0-9_-]*$").unwrap());

#[derive(Clone, Debug, PartialEq)]
pub struct Account {
    pub name: String,
    pub label: String,
    pub aliases: Vec<String>,
    pub username: String,
    pub email_domain: String,
    pub access: String,
    pub cdp_port: Option<u16>,
    pub main_url: String,
}

pub fn normalize_account(raw: &Value) -> Option<Account> {
    let obj = raw.as_object()?;
    let name = js::str_or_empty(obj.get("name")).trim().to_lowercase();
    if !NAME_RE.is_match(&name) {
        return None;
    }
    let first = |a: &str, b: &str| {
        obj.get(a)
            .filter(|v| !v.is_null())
            .or_else(|| obj.get(b).filter(|v| !v.is_null()))
    };
    let port = match first("cdp_port", "cdpPort") {
        Some(v) => js::number(Some(v)),
        None => 0.0,
    };
    let username = js::str_or_empty(obj.get("username"));
    let email_domain = match first("email_domain", "emailDomain") {
        Some(v) => js::string(v),
        None => username.split('@').nth(1).unwrap_or("").to_string(),
    };
    Some(Account {
        label: name.clone(),
        name,
        aliases: match obj.get("aliases") {
            Some(Value::Array(a)) => a.iter().map(js::string).collect(),
            _ => Vec::new(),
        },
        username,
        email_domain,
        access: match obj.get("access") {
            Some(v) if js::truthy(v) => js::string(v),
            _ => "read + drafts".into(),
        },
        cdp_port: (port.fract() == 0.0 && port > 0.0 && port <= 65535.0).then_some(port as u16),
        main_url: first("mailbox_url", "mainUrl")
            .map(js::string)
            .unwrap_or_else(|| DEFAULT_MAILBOX_URL.into()),
    })
}

fn read_config_file(ctx: &Ctx) -> Vec<Value> {
    let path = ctx.config_dir().join("config.json");
    std::fs::read_to_string(path)
        .ok()
        .and_then(|t| serde_json::from_str::<Value>(&t).ok())
        .and_then(|v| v.get("accounts").and_then(Value::as_array).cloned())
        .unwrap_or_default()
}

/// Accounts from config.json, then any saved session not named there.
pub fn list_accounts(ctx: &Ctx) -> Vec<Account> {
    let from_file: Vec<Account> = read_config_file(ctx)
        .iter()
        .filter_map(normalize_account)
        .collect();
    let mut saved: Vec<String> = std::fs::read_dir(ctx.sessions_dir())
        .map(|rd| {
            rd.filter_map(|e| e.ok())
                .filter_map(|e| e.file_name().into_string().ok())
                .filter_map(|f| f.strip_suffix(".json").map(str::to_string))
                .filter(|n| !from_file.iter().any(|a| &a.name == n))
                .collect()
        })
        .unwrap_or_default();
    saved.sort();
    let mut all = from_file;
    all.extend(
        saved
            .iter()
            .filter_map(|n| normalize_account(&serde_json::json!({ "name": n }))),
    );
    all
}

fn squash(s: &str) -> String {
    s.to_lowercase()
        .chars()
        .filter(|c| c.is_ascii_lowercase() || c.is_ascii_digit())
        .collect()
}

fn after_at(s: &str) -> &str {
    s.rfind('@').map(|i| &s[i + 1..]).unwrap_or(s)
}

/// Maps a typed or spoken name, alias, email address or domain to an account name.
/// Returns None for empty input or anything that is not a configured account.
pub fn resolve_account_alias(ctx: &Ctx, value: Option<&Value>) -> Option<String> {
    let raw = match value {
        Some(Value::String(s)) => s.trim().to_lowercase(),
        _ => return None,
    };
    if raw.is_empty() {
        return None;
    }
    let candidates = [squash(&raw), squash(after_at(&raw))];
    for account in list_accounts(ctx) {
        let mut keys = vec![account.name.clone()];
        keys.extend(account.aliases.iter().cloned());
        keys.push(account.username.clone());
        keys.push(after_at(&account.username).to_string());
        let keys: Vec<String> = keys
            .iter()
            .map(|k| squash(k))
            .filter(|k| !k.is_empty())
            .collect();
        if candidates.iter().any(|c| keys.contains(c)) {
            return Some(account.name);
        }
    }
    None
}

pub fn resolve_alias_str(ctx: &Ctx, value: &str) -> Option<String> {
    resolve_account_alias(ctx, Some(&Value::String(value.to_string())))
}

pub fn get_account(ctx: &Ctx, name: &str) -> Option<Account> {
    list_accounts(ctx).into_iter().find(|a| a.name == name)
}
