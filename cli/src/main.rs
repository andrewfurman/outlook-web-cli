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

const TOP_HELP: &str = "{name} {version} - {about}

Usage: {name} <COMMAND> [OPTIONS]

FIRST-TIME SETUP (once per account)
  Open Outlook web in Chrome started with --remote-debugging-port=9230, then run
  `{name} auth chrome --account work --cdp-port 9230`. The session is saved in
  ~/.config/{name} (override with OUTLOOK_WEB_CONFIG_DIR).

COMMANDS
 Sign in
  auth chrome         Import the session from a Chrome where Outlook web is open
  auth cookies        Import login.microsoftonline.com cookies exported from a browser
  auth status         Token and cookie lifetimes per account (no secrets)
  auth refresh        Mint fresh tokens; --rotate also renews the sign-in cookies
  auth logout         Remove an account's saved session
  accounts            Configured accounts and whether a session is saved
  status              Whether each account can get a token, and from where
 Read (nothing is changed)
  messages            Inbox (or another folder), newest first
  search              Full-mailbox search
  read                One message, with its body as text
  agenda              Calendar for a day or range (America/New_York)
  freebusy            When people are busy, and the slots when everyone is free
  attachments         List attachments, or fetch their base64 data
  attachment-inspect  Describe and transcribe an image attachment
 Change (need --yes; mail is never sent)
  draft-new           Save a new draft
  draft-reply         Save a reply or reply-all draft
  draft-delete        Move a draft to Deleted Items

EXAMPLES
  {name} accounts
  {name} messages --account work --unread --limit 10
  {name} search --account work --query \"invoice\"
  {name} read --account work --id AAMkAGExample
  {name} agenda --tomorrow
  {name} freebusy --emails jane@contoso.com,sam@fabrikam.com --tomorrow
  {name} draft-reply --account work --id AAMkAGExample --body \"Thanks!\" --yes
  {name} auth refresh --all --rotate

More: `{name} <COMMAND> --help` (or `{name} help <COMMAND>`) shows every flag with examples.
Output: JSON on stdout; exit code 1 when \"ok\" is false. Any command also takes --json '{...}'.
Safety: drafts only, nothing is ever sent; draft commands do nothing without --yes.";

#[derive(Parser)]
#[command(
    name = "outlook-web-cli",
    version,
    about = "Outlook mail and calendar from the command line, through your Outlook web session",
    help_template = TOP_HELP,
    override_usage = "outlook-web-cli <COMMAND> [OPTIONS]"
)]
struct Cli {
    #[command(subcommand)]
    command: Option<Command>,
}

#[derive(Args)]
struct Common {
    /// JSON object merged over the flags, snake_case fields, e.g. '{"account":"work","limit":5}'
    #[arg(long, value_name = "JSON")]
    json: Option<String>,
}

#[derive(Subcommand)]
enum Command {
    /// Sign in, check, renew or remove saved sessions
    #[command(after_help = "Examples:
  outlook-web-cli auth chrome --account work --cdp-port 9230
  outlook-web-cli auth status
  outlook-web-cli auth refresh --all --rotate

Run `outlook-web-cli auth <SUBCOMMAND> --help` for the flags of each.")]
    Auth {
        #[command(subcommand)]
        command: AuthCommand,
    },
    /// Configured accounts and whether a session is saved
    #[command(after_help = "Examples:
  outlook-web-cli accounts

Takes no flags except --json. Add accounts with `outlook-web-cli auth chrome --help`.")]
    Accounts,
    /// Whether each account can get a token, and from where
    #[command(after_help = "Examples:
  outlook-web-cli status
  outlook-web-cli status --account work
  outlook-web-cli status --account all")]
    Status {
        /// Account name, alias, email domain or address, or `all` (default: every account)
        #[arg(long, value_name = "ACCOUNT")]
        account: Option<String>,
        #[command(flatten)]
        common: Common,
    },
    /// Inbox (or another folder), newest first
    #[command(after_help = "Examples:
  outlook-web-cli messages --unread --limit 10
  outlook-web-cli messages --account work --from jane@contoso.com --since 2026-11-01
  outlook-web-cli messages --account work --folder sent --limit 5
  outlook-web-cli messages --account work --search \"quarterly report\"

Without --account, every account is read and merged newest first.")]
    Messages {
        /// Account name, alias, email domain or address, or `all` (default: every account)
        #[arg(long, value_name = "ACCOUNT")]
        account: Option<String>,
        /// How many messages, 1-50 (default 20)
        #[arg(long, visible_alias = "max-results", value_name = "N")]
        limit: Option<String>,
        /// Only unread messages; bare --unread means true (optional value: true or false)
        #[arg(long, num_args = 0..=1, default_missing_value = "true", value_name = "BOOL")]
        unread: Option<String>,
        /// Only messages from this sender address
        #[arg(long, alias = "sender", value_name = "ADDRESS")]
        from: Option<String>,
        /// Only messages received on or after this date, YYYY-MM-DD
        #[arg(long, value_name = "DATE")]
        since: Option<String>,
        /// Search words within the folder (results in relevance order)
        #[arg(long, value_name = "WORDS")]
        search: Option<String>,
        /// Folder: inbox (default), sent, drafts, deleted, archive or junk
        #[arg(long, value_name = "FOLDER")]
        folder: Option<String>,
        #[command(flatten)]
        common: Common,
    },
    /// Full-mailbox search
    #[command(after_help = "Examples:
  outlook-web-cli search --query \"invoice\"
  outlook-web-cli search --account work --query \"budget review\" --limit 5

--query is required. Searches every folder; use `messages --folder` to list one folder.")]
    Search {
        /// Words to search for (required)
        #[arg(long, value_name = "WORDS")]
        query: Option<String>,
        /// Account name, alias, email domain or address, or `all` (default: every account)
        #[arg(long, value_name = "ACCOUNT")]
        account: Option<String>,
        /// How many results, 1-50 (default 20)
        #[arg(long, visible_alias = "max-results", value_name = "N")]
        limit: Option<String>,
        #[command(flatten)]
        common: Common,
    },
    /// One message, with its body as text
    #[command(after_help = "Examples:
  outlook-web-cli read --account work --id AAMkAGExample
  outlook-web-cli read --account work --id AAMkAGExample --max-body-chars 60000

Get message ids from `messages` or `search` (the \"id\" field).")]
    Read {
        /// Account that holds the message; see `accounts` for names
        #[arg(long, value_name = "ACCOUNT")]
        account: Option<String>,
        /// Message id from `messages` or `search` (required)
        #[arg(long, alias = "message-id", value_name = "MESSAGE_ID")]
        id: Option<String>,
        /// Longest body to return, 2000-60000 characters (default 20000)
        #[arg(long, value_name = "N")]
        max_body_chars: Option<String>,
        #[command(flatten)]
        common: Common,
    },
    /// Calendar for a day or range (America/New_York)
    #[command(after_help = "Examples:
  outlook-web-cli agenda
  outlook-web-cli agenda --tomorrow --account work
  outlook-web-cli agenda --start-date 2026-11-16 --end-date 2026-11-20

With no day flags, shows today. Times are America/New_York.")]
    Agenda {
        /// Account name, alias, email domain or address, or `all` (default: every account)
        #[arg(long, value_name = "ACCOUNT")]
        account: Option<String>,
        /// Show today (the default); bare flag means true (optional value: true or false)
        #[arg(long, num_args = 0..=1, default_missing_value = "true", value_name = "BOOL")]
        today: Option<String>,
        /// Show tomorrow instead; bare flag means true (optional value: true or false)
        #[arg(long, num_args = 0..=1, default_missing_value = "true", value_name = "BOOL")]
        tomorrow: Option<String>,
        /// First day of a range, YYYY-MM-DD or YYYY-MM-DDTHH:MM (use with --end-date)
        #[arg(long, value_name = "DATE")]
        start_date: Option<String>,
        /// Last day of a range, YYYY-MM-DD or YYYY-MM-DDTHH:MM (use with --start-date)
        #[arg(long, value_name = "DATE")]
        end_date: Option<String>,
        /// Most events to return, 1-200 (default 50)
        #[arg(long, visible_alias = "max-results", value_name = "N")]
        limit: Option<String>,
        #[command(flatten)]
        common: Common,
    },
    /// When each person is busy, and the slots when everyone is free
    #[command(after_help = "Examples:
  outlook-web-cli freebusy --emails jane@contoso.com,sam@contoso.com --tomorrow
  outlook-web-cli freebusy --emails jane@contoso.com --start-date 2026-11-17 --end-date 2026-11-19
  outlook-web-cli freebusy --emails jane@contoso.com,sam@fabrikam.com --today --day-start 08:00 --day-end 18:00 --min-minutes 60
  outlook-web-cli freebusy --emails jane@contoso.com --start-date 2026-11-17T11:00 --end-date 2026-11-17T15:00

--emails is required. Works without calendar sharing; times are America/New_York.
With no day flags, checks today. Without --account, the account matching the email domain is used.")]
    Freebusy {
        /// Comma-separated email addresses, up to 20 (required)
        #[arg(
            long,
            alias = "email",
            alias = "people",
            alias = "attendees",
            value_name = "ADDRESSES"
        )]
        emails: Option<String>,
        /// Account to ask from (default: the one matching the email domain, else the only account)
        #[arg(long, value_name = "ACCOUNT")]
        account: Option<String>,
        /// Check today (the default); bare flag means true (optional value: true or false)
        #[arg(long, num_args = 0..=1, default_missing_value = "true", value_name = "BOOL")]
        today: Option<String>,
        /// Check tomorrow; bare flag means true (optional value: true or false)
        #[arg(long, num_args = 0..=1, default_missing_value = "true", value_name = "BOOL")]
        tomorrow: Option<String>,
        /// Start of the range, YYYY-MM-DD or YYYY-MM-DDTHH:MM (use with --end-date)
        #[arg(long, alias = "start", value_name = "DATE")]
        start_date: Option<String>,
        /// End of the range, YYYY-MM-DD or YYYY-MM-DDTHH:MM (use with --start-date)
        #[arg(long, alias = "end", value_name = "DATE")]
        end_date: Option<String>,
        /// Working day start, HH:MM (default 09:00)
        #[arg(long, value_name = "HH:MM")]
        day_start: Option<String>,
        /// Working day end, HH:MM (default 17:00)
        #[arg(long, value_name = "HH:MM")]
        day_end: Option<String>,
        /// Shortest free slot to report, 5-480 minutes (default 30)
        #[arg(long, value_name = "MINUTES")]
        min_minutes: Option<String>,
        /// Include Saturdays and Sundays; bare flag means true (optional value: true or false)
        #[arg(long, num_args = 0..=1, default_missing_value = "true", value_name = "BOOL")]
        weekends: Option<String>,
        #[command(flatten)]
        common: Common,
    },
    /// List attachments, including inline screenshots; optionally their base64 data
    #[command(after_help = "Examples:
  outlook-web-cli attachments --account work --id AAMkAGExample
  outlook-web-cli attachments --account work --id AAMkAGExample --images-only --include-data
  outlook-web-cli attachments --account work --id AAMkAGExample --index 0 --include-data
  outlook-web-cli attachments --account work --id AAMkAGExample --index 0 --offset 250000 --length 250000

Without --include-data only names, types and sizes are listed. Pick one attachment with
--index, --content-id or --attachment-id. --offset/--length pull a big file in pieces.")]
    Attachments {
        /// Account that holds the message; see `accounts` for names
        #[arg(long, value_name = "ACCOUNT")]
        account: Option<String>,
        /// Message id from `messages` or `search` (required)
        #[arg(long, alias = "message-id", value_name = "MESSAGE_ID")]
        id: Option<String>,
        /// Also return base64 file data (files up to --max-data-bytes); optional value: true or false
        #[arg(long, num_args = 0..=1, default_missing_value = "true", value_name = "BOOL")]
        include_data: Option<String>,
        /// Largest file to include data for, 1000-1000000 bytes (default 250000)
        #[arg(long, value_name = "BYTES")]
        max_data_bytes: Option<String>,
        /// Only list image attachments; optional value: true or false
        #[arg(long, num_args = 0..=1, default_missing_value = "true", value_name = "BOOL")]
        images_only: Option<String>,
        /// Pick one attachment by its index in the listing (starts at 0)
        #[arg(long, alias = "attachment-index", value_name = "N")]
        index: Option<String>,
        /// Pick one attachment by its content id (inline images, the cid in the HTML)
        #[arg(long, alias = "cid", value_name = "CID")]
        content_id: Option<String>,
        /// Pick one attachment by its attachment id from the listing
        #[arg(long, value_name = "ATTACHMENT_ID")]
        attachment_id: Option<String>,
        /// Start byte of a range of the picked attachment (returns --length bytes)
        #[arg(long, value_name = "BYTES")]
        offset: Option<String>,
        /// Range length with --offset, 1000-1000000 bytes (default 250000)
        #[arg(long, value_name = "BYTES")]
        length: Option<String>,
        #[command(flatten)]
        common: Common,
    },
    /// Describe and transcribe an image attachment (needs OPENROUTER_API_KEY)
    #[command(
        name = "attachment-inspect",
        after_help = "Examples:
  outlook-web-cli attachment-inspect --account work --id AAMkAGExample
  outlook-web-cli attachment-inspect --account work --id AAMkAGExample --index 1
  outlook-web-cli attachment-inspect --account work --id AAMkAGExample --content-id shot1 --prompt \"List every error message\"

Sends the image to a vision model through OpenRouter, so OPENROUTER_API_KEY must be set
(OUTLOOK_WEB_VISION_MODEL picks the model). With no picker flag, the first image is used."
    )]
    AttachmentInspect {
        /// Account that holds the message; see `accounts` for names
        #[arg(long, value_name = "ACCOUNT")]
        account: Option<String>,
        /// Message id from `messages` or `search` (required)
        #[arg(long, alias = "message-id", value_name = "MESSAGE_ID")]
        id: Option<String>,
        /// Pick the image by its index in `attachments` (default: the first image)
        #[arg(long, alias = "attachment-index", value_name = "N")]
        index: Option<String>,
        /// Pick the image by its content id (inline images)
        #[arg(long, alias = "cid", value_name = "CID")]
        content_id: Option<String>,
        /// What to ask the model about the image (default: describe it and transcribe all text)
        #[arg(long, value_name = "TEXT")]
        prompt: Option<String>,
        #[command(flatten)]
        common: Common,
    },
    /// Save a new draft (requires --yes)
    #[command(
        name = "draft-new",
        after_help = "Examples:
  outlook-web-cli draft-new --account work --to jane@contoso.com --subject \"Agenda\" --body \"Hi Jane, here is the agenda.\" --yes
  outlook-web-cli draft-new --account work --to jane@contoso.com,sam@fabrikam.com --cc pat@contoso.com --subject \"Notes\" --body \"See below.\" --yes
  outlook-web-cli draft-new --account work --to jane@contoso.com --subject \"Report\" --body-html \"<p>Hello <b>Jane</b></p>\" --yes

The draft is saved in Drafts, never sent. Without --yes nothing is saved (the command only
explains what it would do). --to is required; give --body or --body-html.
Inline images need --json, e.g. --json '{\"inline_images\":[{\"url\":\"https://example.com/a.png\",\"content_id\":\"shot1\"}]}'."
    )]
    DraftNew {
        /// Account to save the draft in; see `accounts` for names
        #[arg(long, value_name = "ACCOUNT")]
        account: Option<String>,
        /// Recipient address(es), comma-separated (required)
        #[arg(long, value_name = "ADDRESSES")]
        to: Option<String>,
        /// Cc address(es), comma-separated
        #[arg(long, value_name = "ADDRESSES")]
        cc: Option<String>,
        /// Subject line
        #[arg(long, value_name = "TEXT")]
        subject: Option<String>,
        /// Plain-text body
        #[arg(long, alias = "message", alias = "text", value_name = "TEXT")]
        body: Option<String>,
        /// HTML body (used instead of --body)
        #[arg(long, alias = "html", value_name = "HTML")]
        body_html: Option<String>,
        /// Confirm: actually save the draft (without it nothing is saved)
        #[arg(long, num_args = 0..=1, default_missing_value = "true", value_name = "BOOL")]
        yes: Option<String>,
        #[command(flatten)]
        common: Common,
    },
    /// Save a reply or reply-all draft above the quoted original (requires --yes)
    #[command(
        name = "draft-reply",
        after_help = "Examples:
  outlook-web-cli draft-reply --account work --id AAMkAGExample --body \"Thanks, done.\" --yes
  outlook-web-cli draft-reply --account work --id AAMkAGExample --reply-all --body \"Agreed.\" --yes
  outlook-web-cli draft-reply --account work --id AAMkAGExample --body \"Adding Sam.\" --add-to sam@fabrikam.com --add-cc pat@contoso.com --yes

The reply is saved in Drafts, never sent. Without --yes nothing is saved (the command only
explains what it would do). --id and --body (or --body-html) are required."
    )]
    DraftReply {
        /// Account that holds the message; see `accounts` for names
        #[arg(long, value_name = "ACCOUNT")]
        account: Option<String>,
        /// Id of the message to reply to, from `messages` or `search` (required)
        #[arg(long, alias = "message-id", value_name = "MESSAGE_ID")]
        id: Option<String>,
        /// Plain-text reply, placed above the quoted original
        #[arg(long, alias = "message", alias = "text", value_name = "TEXT")]
        body: Option<String>,
        /// HTML reply (used instead of --body)
        #[arg(long, alias = "html", value_name = "HTML")]
        body_html: Option<String>,
        /// Reply to everyone on the original message; optional value: true or false
        #[arg(long, num_args = 0..=1, default_missing_value = "true", value_name = "BOOL")]
        reply_all: Option<String>,
        /// Extra To recipient(s), comma-separated
        #[arg(long, value_name = "ADDRESSES")]
        add_to: Option<String>,
        /// Extra Cc recipient(s), comma-separated
        #[arg(long, value_name = "ADDRESSES")]
        add_cc: Option<String>,
        /// Confirm: actually save the draft (without it nothing is saved)
        #[arg(long, num_args = 0..=1, default_missing_value = "true", value_name = "BOOL")]
        yes: Option<String>,
        #[command(flatten)]
        common: Common,
    },
    /// Move a draft to Deleted Items; anything that isn't a draft is refused (requires --yes)
    #[command(
        name = "draft-delete",
        after_help = "Examples:
  outlook-web-cli draft-delete --account work --id AAMkAGDraftExample --yes
  outlook-web-cli messages --account work --folder drafts

Find draft ids with `messages --folder drafts`. The draft moves to Deleted Items; received
or sent messages are refused. Without --yes nothing is deleted."
    )]
    DraftDelete {
        /// Account that holds the draft; see `accounts` for names
        #[arg(long, value_name = "ACCOUNT")]
        account: Option<String>,
        /// Draft id from `messages --folder drafts` (required)
        #[arg(long, alias = "message-id", value_name = "DRAFT_ID")]
        id: Option<String>,
        /// Confirm: actually delete the draft (without it nothing is deleted)
        #[arg(long, num_args = 0..=1, default_missing_value = "true", value_name = "BOOL")]
        yes: Option<String>,
        #[command(flatten)]
        common: Common,
    },
}

#[derive(Subcommand)]
enum AuthCommand {
    /// Import the session from a Chrome where Outlook web is open
    #[command(after_help = "Examples:
  outlook-web-cli auth chrome --account work --cdp-port 9230
  outlook-web-cli auth chrome --account fabrikam --cdp-port 9231 --mailbox-url https://outlook.cloud.microsoft/mail/jane@fabrikam.com/
  outlook-web-cli auth chrome --account work

Start Chrome first: google-chrome --remote-debugging-port=9230 --user-data-dir=$HOME/.chrome-outlook
and sign in to Outlook web there (tick \"Stay signed in\"). The port is remembered per account,
so later runs can omit --cdp-port. --account and a port are required the first time.")]
    Chrome {
        /// Name to save the account under: letters, digits, - or _ (required)
        #[arg(long, value_name = "NAME")]
        account: Option<String>,
        /// The Chrome --remote-debugging-port where Outlook web is open (required the first time; remembered)
        #[arg(long, value_name = "PORT")]
        cdp_port: Option<String>,
        /// Pick the tab for this mailbox URL (for a second mailbox in the same Chrome)
        #[arg(long, value_name = "URL")]
        mailbox_url: Option<String>,
    },
    /// Import login.microsoftonline.com cookies (Cookie-Editor JSON or cookies.txt)
    #[command(after_help = "Examples:
  outlook-web-cli auth cookies --account work --file ~/Downloads/cookies.json --username jane@contoso.com
  outlook-web-cli auth cookies --account work --file ~/Downloads/cookies.txt
  cat cookies.json | outlook-web-cli auth cookies --account work --file -

Export the cookies for login.microsoftonline.com from your browser first, then delete the
export file: it is a credential. Only that domain's cookies are kept.")]
    Cookies {
        /// Name to save the account under: letters, digits, - or _ (required)
        #[arg(long, value_name = "NAME")]
        account: Option<String>,
        /// Cookie export, Cookie-Editor JSON or Netscape cookies.txt; `-` reads stdin (required)
        #[arg(long, value_name = "PATH")]
        file: Option<String>,
        /// Sign-in address, remembered for the account (e.g. jane@contoso.com)
        #[arg(long, value_name = "EMAIL")]
        username: Option<String>,
        /// Azure tenant id if known (default: organizations)
        #[arg(long, value_name = "TENANT_ID")]
        tenant: Option<String>,
    },
    /// Token minutes left, refresh-token hours left and cookie expiry (no secrets)
    #[command(after_help = "Examples:
  outlook-web-cli auth status
  outlook-web-cli auth status --account work")]
    Status {
        /// Only this account (default: every account)
        #[arg(long, value_name = "ACCOUNT")]
        account: Option<String>,
    },
    /// Mint fresh tokens now; --rotate also renews the sign-in cookies (cron keepalive)
    #[command(after_help = "Examples:
  outlook-web-cli auth refresh --account work
  outlook-web-cli auth refresh --all
  outlook-web-cli auth refresh --all --rotate

Run `auth refresh --all --rotate` from cron every few hours to keep every account signed in.
With neither --account nor --all, every account is refreshed.")]
    Refresh {
        /// Only this account (default: every account)
        #[arg(long, value_name = "ACCOUNT")]
        account: Option<String>,
        /// Refresh every account (the default when --account is not given)
        #[arg(long)]
        all: bool,
        /// Also renew the sign-in cookies with a silent sign-in (keeps the session alive for months)
        #[arg(long)]
        rotate: bool,
    },
    /// Remove an account's saved session
    #[command(after_help = "Examples:
  outlook-web-cli auth logout --account work

Deletes the saved session file for that account (you must sign in again). It does not
sign you out of Outlook web in the browser.")]
    Logout {
        /// Account whose saved session to delete (required)
        #[arg(long, value_name = "ACCOUNT")]
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

/// The subcommand path as typed, e.g. "auth chrome" or "read".
fn command_path(m: &ArgMatches) -> String {
    let mut parts = Vec::new();
    let mut cur = m;
    while let Some((name, sub)) = cur.subcommand() {
        parts.push(name);
        cur = sub;
    }
    parts.join(" ")
}

/// A failed result about a missing or invalid argument also says where the examples are.
fn with_hint(mut result: Value, path: &str) -> Value {
    let status = result["status"].as_str().unwrap_or("");
    let message = result["message"].as_str().unwrap_or("");
    let usage_problem = status.starts_with("missing_")
        || status == "invalid_folder"
        || message.contains("is required");
    if result.get("ok") == Some(&json!(false)) && usage_problem && !path.is_empty() {
        if let Some(o) = result.as_object_mut() {
            o.insert(
                "hint".into(),
                json!(format!(
                    "Run `outlook-web-cli {path} --help` for the flags and examples."
                )),
            );
        }
    }
    result
}

fn main() {
    let matches = Cli::command().get_matches();
    let path = command_path(&matches);
    let cli = match Cli::from_arg_matches(&matches) {
        Ok(c) => c,
        Err(e) => e.exit(),
    };
    let ctx = Ctx::from_env();
    match run(&ctx, &matches, cli) {
        Ok(v) => out(&with_hint(v, &path)),
        Err(e) => out(&with_hint(
            json!({ "ok": false, "status": "error", "message": format!("{e:#}") }),
            &path,
        )),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Every command and subcommand has a description and an Examples section, and every
    /// flag has help text (so a new flag can't be added undocumented).
    #[test]
    fn every_command_and_flag_is_documented() {
        fn walk(cmd: &clap::Command, path: &str, count: &mut usize) {
            for sub in cmd.get_subcommands() {
                if sub.get_name() == "help" {
                    continue;
                }
                let p = format!("{path} {}", sub.get_name());
                assert!(sub.get_about().is_some(), "{p}: no description");
                let after = sub.get_after_help().map(|h| h.to_string());
                assert!(
                    after.is_some_and(|h| h.contains("Examples:")),
                    "{p}: no Examples: section in after_help"
                );
                for arg in sub.get_arguments() {
                    assert!(
                        arg.get_help()
                            .is_some_and(|h| !h.to_string().trim().is_empty()),
                        "{p}: --{} has no help text",
                        arg.get_id()
                    );
                    if arg.get_action().takes_values() {
                        assert!(
                            arg.get_value_names().is_some(),
                            "{p}: --{} has no value name",
                            arg.get_id()
                        );
                    }
                }
                *count += 1;
                walk(sub, &p, count);
            }
        }
        let mut count = 0;
        walk(&Cli::command(), "outlook-web-cli", &mut count);
        assert_eq!(count, 18, "command count changed: update the help too");
    }

    #[test]
    fn top_level_help_lists_every_command_once() {
        let help = Cli::command().render_long_help().to_string();
        assert!(help.contains("EXAMPLES"));
        for sub in Cli::command().get_subcommands() {
            let name = sub.get_name();
            if name != "help" {
                assert!(
                    help.contains(&format!("  {name} ")),
                    "{name} not in top help"
                );
            }
        }
    }
}
