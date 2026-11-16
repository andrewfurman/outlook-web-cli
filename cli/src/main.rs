use anyhow::{bail, Context, Result};
use clap::parser::ValueSource;
use clap::{ArgMatches, Args, CommandFactory, FromArgMatches, Parser, Subcommand};
use outlook_web_cli::accounts::{get_account, list_accounts, resolve_alias_str, NAME_RE};
use outlook_web_cli::commands;
use outlook_web_cli::js::{self, Body};
use outlook_web_cli::session::{
    self, delete_session, import_cookies, import_from_chrome, refresh_accounts, refresh_session,
    session_status,
};
use outlook_web_cli::Ctx;
use serde_json::{json, Map, Value};
use std::io::Read;

const OVERVIEW: &str = "SIGN IN (once per account; the session is saved to ~/.config/outlook-web-cli):
  outlook-web-cli auth chrome  --account work --cdp-port 9230 [--mailbox-url URL]
      import from a Chrome where Outlook web is open (started with --remote-debugging-port)
  outlook-web-cli auth cookies --account work --file cookies.json [--username me@contoso.com]
      import login.microsoftonline.com cookies exported from any browser (JSON or cookies.txt; - = stdin)
  outlook-web-cli auth status  [--account work]
  outlook-web-cli auth refresh [--account work | --all] [--rotate]   mint a fresh token; --rotate also renews the sign-in cookies
  outlook-web-cli auth logout  --account work

MAIL AND CALENDAR (read-only, plus drafts; nothing ever sends mail):
  outlook-web-cli accounts
  outlook-web-cli status      [--account work|all]
  outlook-web-cli messages    [--account all] [--limit 20] [--unread] [--from addr] [--since 2026-11-01] [--folder inbox|sent|drafts|archive]
  outlook-web-cli search      --query \"invoice\" [--account all] [--limit 20]
  outlook-web-cli read        --account work --id MESSAGE_ID [--max-body-chars 20000]
  outlook-web-cli agenda      [--account all] [--today | --tomorrow | --start-date D --end-date D]
  outlook-web-cli freebusy    --emails a@contoso.com,b@contoso.com [--account work] [--today | --tomorrow | --start-date D --end-date D]
                              [--day-start 09:00 --day-end 17:00] [--min-minutes 30] [--weekends]
      when each person is busy, and the slots when everyone is free (Eastern time; no calendar sharing needed)
  outlook-web-cli attachments --account work --id MESSAGE_ID [--include-data] [--images-only] [--index N] [--offset N --length N]
  outlook-web-cli attachment-inspect --account work --id MESSAGE_ID [--index N] [--prompt TEXT]   (needs OPENROUTER_API_KEY)
  outlook-web-cli draft-new   --account work --to a@b.com --subject S --body TEXT --yes
  outlook-web-cli draft-reply --account work --id MESSAGE_ID --body TEXT [--reply-all] [--add-to x@y.com] --yes
  outlook-web-cli draft-delete --account work --id DRAFT_ID --yes

Any command also takes --json '{\"account\":\"work\",...}' (snake_case fields, merged over the flags).
Output is JSON, and the exit code is 1 when ok is false. Draft commands do nothing without --yes.";

#[derive(Parser)]
#[command(
    name = "outlook-web-cli",
    version,
    about = "Outlook mail and calendar from the command line, through your Outlook web session",
    long_about = "Unofficial Outlook (Microsoft 365) mail and calendar CLI. Uses your Outlook web session instead of an Azure app registration: no admin consent and no client secret. It never sends mail.",
    after_help = OVERVIEW
)]
struct Cli {
    #[command(subcommand)]
    command: Option<Command>,
}

#[derive(Args)]
struct Common {
    /// JSON body merged over the flags, e.g. '{"account":"work","limit":5}'
    #[arg(long, value_name = "JSON")]
    json: Option<String>,
}

#[derive(Subcommand)]
enum Command {
    /// Sign in, check, renew or remove saved sessions
    Auth {
        #[command(subcommand)]
        command: AuthCommand,
    },
    /// Configured accounts and whether a session is saved
    Accounts,
    /// Whether each account can get a token, and from where
    Status {
        /// Account name, alias, email domain or address, or "all"
        #[arg(long)]
        account: Option<String>,
        #[command(flatten)]
        common: Common,
    },
    /// Inbox (or another folder), newest first
    Messages {
        #[arg(long)]
        account: Option<String>,
        /// 1-50 (default 20)
        #[arg(long, visible_alias = "max-results")]
        limit: Option<String>,
        /// Only unread messages
        #[arg(long, num_args = 0..=1, default_missing_value = "true", value_name = "BOOL")]
        unread: Option<String>,
        /// Sender address
        #[arg(long, alias = "sender")]
        from: Option<String>,
        /// YYYY-MM-DD
        #[arg(long)]
        since: Option<String>,
        /// Search words (relevance order)
        #[arg(long)]
        search: Option<String>,
        /// inbox, sent, drafts, deleted, archive or junk
        #[arg(long)]
        folder: Option<String>,
        #[command(flatten)]
        common: Common,
    },
    /// Full-mailbox search
    Search {
        #[arg(long)]
        query: Option<String>,
        #[arg(long)]
        account: Option<String>,
        #[arg(long, visible_alias = "max-results")]
        limit: Option<String>,
        #[command(flatten)]
        common: Common,
    },
    /// One message, with its body as text
    Read {
        #[arg(long)]
        account: Option<String>,
        #[arg(long, alias = "message-id")]
        id: Option<String>,
        /// 2000-60000 (default 20000)
        #[arg(long)]
        max_body_chars: Option<String>,
        #[command(flatten)]
        common: Common,
    },
    /// Calendar for a day or range (America/New_York)
    Agenda {
        #[arg(long)]
        account: Option<String>,
        #[arg(long, num_args = 0..=1, default_missing_value = "true", value_name = "BOOL")]
        today: Option<String>,
        #[arg(long, num_args = 0..=1, default_missing_value = "true", value_name = "BOOL")]
        tomorrow: Option<String>,
        #[arg(long)]
        start_date: Option<String>,
        #[arg(long)]
        end_date: Option<String>,
        #[arg(long, visible_alias = "max-results")]
        limit: Option<String>,
        #[command(flatten)]
        common: Common,
    },
    /// When each person is busy, and the slots when everyone is free
    Freebusy {
        /// Comma-separated addresses (up to 20)
        #[arg(long, alias = "email", alias = "people", alias = "attendees")]
        emails: Option<String>,
        #[arg(long)]
        account: Option<String>,
        #[arg(long, num_args = 0..=1, default_missing_value = "true", value_name = "BOOL")]
        today: Option<String>,
        #[arg(long, num_args = 0..=1, default_missing_value = "true", value_name = "BOOL")]
        tomorrow: Option<String>,
        /// YYYY-MM-DD or YYYY-MM-DDTHH:MM
        #[arg(long, alias = "start")]
        start_date: Option<String>,
        /// YYYY-MM-DD or YYYY-MM-DDTHH:MM
        #[arg(long, alias = "end")]
        end_date: Option<String>,
        /// Working day start (default 09:00)
        #[arg(long)]
        day_start: Option<String>,
        /// Working day end (default 17:00)
        #[arg(long)]
        day_end: Option<String>,
        /// Shortest slot to report, 5-480 (default 30)
        #[arg(long)]
        min_minutes: Option<String>,
        /// Include Saturdays and Sundays
        #[arg(long, num_args = 0..=1, default_missing_value = "true", value_name = "BOOL")]
        weekends: Option<String>,
        #[command(flatten)]
        common: Common,
    },
    /// List attachments, including inline screenshots; optionally their base64 data
    Attachments {
        #[arg(long)]
        account: Option<String>,
        #[arg(long, alias = "message-id")]
        id: Option<String>,
        #[arg(long, num_args = 0..=1, default_missing_value = "true", value_name = "BOOL")]
        include_data: Option<String>,
        #[arg(long)]
        max_data_bytes: Option<String>,
        #[arg(long, num_args = 0..=1, default_missing_value = "true", value_name = "BOOL")]
        images_only: Option<String>,
        #[arg(long, alias = "attachment-index")]
        index: Option<String>,
        #[arg(long, alias = "cid")]
        content_id: Option<String>,
        #[arg(long)]
        attachment_id: Option<String>,
        /// Byte offset of a range of one attachment
        #[arg(long)]
        offset: Option<String>,
        /// Range length, 1000-1000000 (default 250000)
        #[arg(long)]
        length: Option<String>,
        #[command(flatten)]
        common: Common,
    },
    /// Describe and transcribe an image attachment (needs OPENROUTER_API_KEY)
    #[command(name = "attachment-inspect")]
    AttachmentInspect {
        #[arg(long)]
        account: Option<String>,
        #[arg(long, alias = "message-id")]
        id: Option<String>,
        #[arg(long, alias = "attachment-index")]
        index: Option<String>,
        #[arg(long, alias = "cid")]
        content_id: Option<String>,
        #[arg(long)]
        prompt: Option<String>,
        #[command(flatten)]
        common: Common,
    },
    /// Save a new draft (requires --yes)
    #[command(name = "draft-new")]
    DraftNew {
        #[arg(long)]
        account: Option<String>,
        #[arg(long)]
        to: Option<String>,
        #[arg(long)]
        cc: Option<String>,
        #[arg(long)]
        subject: Option<String>,
        /// Plain text body
        #[arg(long, alias = "message", alias = "text")]
        body: Option<String>,
        /// HTML body (used instead of --body)
        #[arg(long, alias = "html")]
        body_html: Option<String>,
        /// Actually save the draft
        #[arg(long, num_args = 0..=1, default_missing_value = "true", value_name = "BOOL")]
        yes: Option<String>,
        #[command(flatten)]
        common: Common,
    },
    /// Save a reply or reply-all draft above the quoted original (requires --yes)
    #[command(name = "draft-reply")]
    DraftReply {
        #[arg(long)]
        account: Option<String>,
        #[arg(long, alias = "message-id")]
        id: Option<String>,
        #[arg(long, alias = "message", alias = "text")]
        body: Option<String>,
        #[arg(long, alias = "html")]
        body_html: Option<String>,
        #[arg(long, num_args = 0..=1, default_missing_value = "true", value_name = "BOOL")]
        reply_all: Option<String>,
        /// Extra To recipients
        #[arg(long)]
        add_to: Option<String>,
        /// Extra Cc recipients
        #[arg(long)]
        add_cc: Option<String>,
        #[arg(long, num_args = 0..=1, default_missing_value = "true", value_name = "BOOL")]
        yes: Option<String>,
        #[command(flatten)]
        common: Common,
    },
    /// Move a draft to Deleted Items; anything that isn't a draft is refused (requires --yes)
    #[command(name = "draft-delete")]
    DraftDelete {
        #[arg(long)]
        account: Option<String>,
        #[arg(long, alias = "message-id")]
        id: Option<String>,
        #[arg(long, num_args = 0..=1, default_missing_value = "true", value_name = "BOOL")]
        yes: Option<String>,
        #[command(flatten)]
        common: Common,
    },
}

#[derive(Subcommand)]
enum AuthCommand {
    /// Import the session from a Chrome where Outlook web is open
    Chrome {
        #[arg(long)]
        account: Option<String>,
        /// The Chrome --remote-debugging-port (remembered in config.json)
        #[arg(long)]
        cdp_port: Option<String>,
        /// Pick the tab for this mailbox URL (for a second mailbox in one Chrome)
        #[arg(long)]
        mailbox_url: Option<String>,
    },
    /// Import login.microsoftonline.com cookies (Cookie-Editor JSON or cookies.txt)
    Cookies {
        #[arg(long)]
        account: Option<String>,
        /// Export file, or - for stdin
        #[arg(long)]
        file: Option<String>,
        #[arg(long)]
        username: Option<String>,
        #[arg(long)]
        tenant: Option<String>,
    },
    /// Token minutes left, refresh-token hours left and cookie expiry (no secrets)
    Status {
        #[arg(long)]
        account: Option<String>,
    },
    /// Mint fresh tokens now; --rotate also renews the sign-in cookies (cron keepalive)
    Refresh {
        #[arg(long)]
        account: Option<String>,
        #[arg(long)]
        all: bool,
        #[arg(long)]
        rotate: bool,
    },
    /// Remove an account's saved session
    Logout {
        #[arg(long)]
        account: Option<String>,
    },
}

fn out(result: &Value) {
    println!(
        "{}",
        serde_json::to_string_pretty(result).unwrap_or_else(|_| "{}".into())
    );
    if result.get("ok") == Some(&json!(false)) {
        std::process::exit(1);
    }
}

/// The flags given on the command line, as a snake_case body (values as strings).
fn body_from_matches(m: &ArgMatches) -> Body {
    let mut body = Map::new();
    for id in m.ids() {
        let id = id.as_str();
        if matches!(id, "json" | "yes") || m.value_source(id) != Some(ValueSource::CommandLine) {
            continue;
        }
        if let Ok(Some(v)) = m.try_get_one::<String>(id) {
            body.insert(id.to_string(), json!(v));
        }
    }
    body
}

fn account_name(ctx: &Ctx, account: Option<&str>, required: bool) -> Result<String> {
    let raw = account.unwrap_or("").trim();
    let name = resolve_alias_str(ctx, raw).unwrap_or_else(|| raw.to_lowercase());
    if required && !NAME_RE.is_match(&name) {
        bail!("--account NAME is required (letters, digits, - or _).");
    }
    Ok(name)
}

/// Remembers the Chrome port and username so a dead session can be re-imported.
fn remember_account(ctx: &Ctx, name: &str, fields: &[(&str, Value)]) -> Result<()> {
    let file = ctx.config_dir().join("config.json");
    let mut config = std::fs::read_to_string(&file)
        .ok()
        .and_then(|t| serde_json::from_str::<Value>(&t).ok())
        .filter(Value::is_object)
        .unwrap_or_else(|| json!({ "accounts": [] }));
    if !config["accounts"].is_array() {
        config["accounts"] = json!([]);
    }
    let accounts = config["accounts"].as_array_mut().unwrap();
    let set: Vec<&(&str, Value)> = fields.iter().filter(|(_, v)| js::truthy(v)).collect();
    match accounts
        .iter_mut()
        .find(|a| a.get("name") == Some(&json!(name)))
    {
        Some(existing) if existing.is_object() => {
            for (k, v) in set {
                existing[*k] = v.clone();
            }
        }
        _ => {
            let mut entry = Map::new();
            entry.insert("name".into(), json!(name));
            for (k, v) in set {
                entry.insert((*k).into(), v.clone());
            }
            accounts.push(Value::Object(entry));
        }
    }
    session::create_private_dir(&ctx.config_dir())?;
    session::write_private(
        &file,
        format!("{}\n", serde_json::to_string_pretty(&config)?).as_bytes(),
    )?;
    Ok(())
}

fn verified(ctx: &Ctx, name: &str, check: &session::Refresh) -> Value {
    let mut o = session_status(ctx, name)
        .as_object()
        .cloned()
        .unwrap_or_default();
    o.insert("ok".into(), json!(check.ok));
    o.insert(
        "verified".into(),
        if check.ok {
            json!(format!("token minted via {}", check.via))
        } else {
            json!(check.errors)
        },
    );
    Value::Object(o)
}

fn auth(ctx: &Ctx, cmd: AuthCommand) -> Result<Value> {
    match cmd {
        AuthCommand::Chrome {
            account,
            cdp_port,
            mailbox_url,
        } => {
            let name = account_name(ctx, account.as_deref(), true)?;
            let known = get_account(ctx, &name);
            let port = match cdp_port.as_deref().filter(|p| !p.is_empty()) {
                Some(p) => js::number(Some(&json!(p))),
                None => known
                    .as_ref()
                    .and_then(|k| k.cdp_port)
                    .map(f64::from)
                    .unwrap_or(0.0),
            };
            if !(port > 0.0 && port <= 65535.0 && port.fract() == 0.0) {
                bail!("--cdp-port is required (the Chrome --remote-debugging-port where Outlook web is open).");
            }
            let port = port as u16;
            let main_url = mailbox_url
                .clone()
                .filter(|u| !u.is_empty())
                .or_else(|| known.map(|k| k.main_url));
            let imported = import_from_chrome(ctx, &name, port, main_url.as_deref())?;
            if imported["ok"] != json!(true) {
                return Ok(imported);
            }
            remember_account(
                ctx,
                &name,
                &[
                    ("cdp_port", json!(port)),
                    ("mailbox_url", json!(mailbox_url.unwrap_or_default())),
                    ("username", imported["username"].clone()),
                ],
            )?;
            let check = refresh_session(ctx, &name, None, false)?;
            Ok(verified(ctx, &name, &check))
        }
        AuthCommand::Cookies {
            account,
            file,
            username,
            tenant,
        } => {
            let name = account_name(ctx, account.as_deref(), true)?;
            let Some(file) = file.filter(|f| !f.is_empty()) else {
                bail!("--file PATH (or - for stdin) is required.");
            };
            let text = if file == "-" {
                let mut s = String::new();
                std::io::stdin().read_to_string(&mut s)?;
                s
            } else {
                std::fs::read_to_string(&file).with_context(|| format!("Can't read {file}"))?
            };
            let username = username.unwrap_or_default();
            let imported =
                import_cookies(ctx, &name, &text, &username, &tenant.unwrap_or_default())?;
            if imported["ok"] != json!(true) {
                return Ok(imported);
            }
            remember_account(ctx, &name, &[("username", json!(username))])?;
            let check = refresh_session(ctx, &name, None, true)?;
            Ok(verified(ctx, &name, &check))
        }
        AuthCommand::Status { account } => {
            let names = match account.as_deref().filter(|a| !a.is_empty()) {
                Some(a) => vec![account_name(ctx, Some(a), true)?],
                None => list_accounts(ctx).into_iter().map(|a| a.name).collect(),
            };
            let accounts: Vec<Value> = names
                .iter()
                .map(|n| {
                    let mut o = session_status(ctx, n)
                        .as_object()
                        .cloned()
                        .unwrap_or_default();
                    o.insert(
                        "cdp_port".into(),
                        json!(get_account(ctx, n).and_then(|a| a.cdp_port)),
                    );
                    Value::Object(o)
                })
                .collect();
            Ok(json!({
                "ok": true,
                "config_dir": ctx.config_dir().display().to_string(),
                "accounts": accounts,
            }))
        }
        AuthCommand::Refresh {
            account,
            all,
            rotate,
        } => {
            let accounts = match account.as_deref().filter(|a| !a.is_empty()) {
                Some(a) if !all => Some(vec![account_name(ctx, Some(a), true)?]),
                _ => None,
            };
            refresh_accounts(ctx, accounts, rotate)
        }
        AuthCommand::Logout { account } => {
            let name = account_name(ctx, account.as_deref(), true)?;
            Ok(json!({ "ok": true, "account": name, "removed": delete_session(ctx, &name) }))
        }
    }
}

fn run(ctx: &Ctx, matches: &ArgMatches, cli: Cli) -> Result<Value> {
    let Some(command) = cli.command else {
        Cli::command().print_long_help()?;
        std::process::exit(0);
    };
    match command {
        Command::Auth { command } => auth(ctx, command),
        Command::Accounts => {
            let accounts: Vec<Value> = list_accounts(ctx)
                .into_iter()
                .map(|a| {
                    json!({
                        "account": a.name,
                        "username": a.username,
                        "aliases": a.aliases,
                        "cdp_port": a.cdp_port,
                        "saved_session": session_status(ctx, &a.name)["saved"],
                    })
                })
                .collect();
            Ok(json!({ "ok": true, "accounts": accounts }))
        }
        _ => {
            let (name, sub) = matches.subcommand().expect("subcommand");
            let mut body = body_from_matches(sub);
            if let Ok(Some(text)) = sub.try_get_one::<String>("json") {
                match serde_json::from_str::<Value>(text).context("--json is not valid JSON")? {
                    Value::Object(extra) => {
                        for (k, v) in extra {
                            body.insert(k, v);
                        }
                    }
                    _ => bail!("--json must be a JSON object."),
                }
            }
            let confirmed = matches!(
                sub.try_get_one::<String>("yes"),
                Ok(Some(v)) if v == "true"
            );
            commands::run(ctx, name, &body, confirmed)
        }
    }
}

fn main() {
    let matches = Cli::command().get_matches();
    let cli = match Cli::from_arg_matches(&matches) {
        Ok(c) => c,
        Err(e) => e.exit(),
    };
    let ctx = Ctx::from_env();
    match run(&ctx, &matches, cli) {
        Ok(v) => out(&v),
        Err(e) => out(&json!({ "ok": false, "status": "error", "message": format!("{e:#}") })),
    }
}
