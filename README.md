# outlook-web-cli

An unofficial **Rust CLI** for Outlook (Microsoft 365) mail and calendar, using your **Outlook web session** instead of an Azure app registration. If you can open Outlook in a browser, you can use it from the terminal, from scripts, and from AI agents. No admin consent and no client secret.

It reads mail, searches, reads the calendar and attachments, checks when coworkers are free or busy, and saves drafts. **It never sends mail**: there is no send code path, and draft commands do nothing without `--yes`.

```bash
outlook-web-cli auth chrome --account work --cdp-port 9230
outlook-web-cli messages --account work --unread --limit 10
outlook-web-cli read --account work --id AAMkAG...
outlook-web-cli agenda --today
outlook-web-cli freebusy --emails alice@contoso.com,bob@contoso.com --tomorrow
outlook-web-cli draft-reply --account work --id AAMkAG... --reply-all --body "Thanks, done." --yes
```

## Install

Install a current stable Rust toolchain from [rustup.rs](https://rustup.rs/) (1.88 or newer), then:

```bash
cargo install --locked --git https://github.com/andrewfurman/outlook-web-cli
outlook-web-cli --help
```

Cargo places `outlook-web-cli` in `~/.cargo/bin`; add that directory to your PATH if needed. From a checkout:

```bash
cargo install --locked --path cli
```

To build a release binary without installing it (Linux x86_64 shown; the binary lands in `cli/target/x86_64-unknown-linux-gnu/release/outlook-web-cli`):

```bash
cargo build --locked --release --manifest-path cli/Cargo.toml --target x86_64-unknown-linux-gnu
```

It runs on Linux and macOS. Sessions saved by the earlier Node.js version (v0.3.0) keep working unchanged: the config directory, `config.json` and the session file format are the same.

## Multiple accounts

Each mailbox is a named account with its own saved session, so work mailboxes in different organizations live side by side:

```bash
outlook-web-cli auth chrome --account contoso  --cdp-port 9230
outlook-web-cli auth chrome --account fabrikam --cdp-port 9231 --mailbox-url https://outlook.cloud.microsoft/mail/me@fabrikam.com/
outlook-web-cli accounts                                   # both, with their aliases and session state
outlook-web-cli messages --unread                          # every account, newest first across both
outlook-web-cli agenda --tomorrow --account fabrikam       # one account; aliases, email domains and addresses also work
outlook-web-cli freebusy --emails pat@fabrikam.com         # picks fabrikam from the email domain
outlook-web-cli auth refresh --all --rotate                # keep every account signed in (cron)
```

Read commands (`status`, `messages`, `search`, `agenda`) cover every account when `--account` is left out or set to `all`. Commands that act on one message or draft need `--account`.

## Signing in: how the web session works

Outlook web is a single-page app that signs in with Microsoft's MSAL library. Two things in your browser keep you signed in, and `outlook-web-cli` saves and uses both, the same way the web app does:

| What | Lifetime | How outlook-web-cli uses it |
| --- | --- | --- |
| Outlook web's **refresh token** (in the page's local storage) | 24 hours from sign-in (Microsoft's fixed limit for single-page apps) | Redeemed at `login.microsoftonline.com` with Outlook's `Origin`, for a 60–90 minute access token |
| Microsoft's **sign-in cookies** (`ESTSAUTHPERSISTENT` on `login.microsoftonline.com`) | Months, when "Stay signed in" was chosen | A silent sign-in (`prompt=none` + PKCE), exactly like Outlook web's hidden renewal iframe, gives a new refresh token and renewed cookies |

The access token is for `https://outlook.office.com` with the web app's own `Mail.ReadWrite` and `Calendars` scopes, and calls go to the Outlook REST API. Each call uses the saved access token while it's valid, then the refresh token, then the cookies. Run `outlook-web-cli auth refresh --all --rotate` from cron every few hours to keep the cookies renewed.

### Import from Chrome

Start Chrome with a DevTools port, sign in to Outlook web there (tick "Stay signed in"), then:

```bash
google-chrome --remote-debugging-port=9230 --user-data-dir="$HOME/.chrome-outlook" https://outlook.cloud.microsoft/mail/
outlook-web-cli auth chrome --account work --cdp-port 9230
```

The port is remembered. If the saved session ever stops working, `outlook-web-cli` re-imports from that Chrome automatically. If that Chrome isn't running, it goes straight to "needs sign-in".

For a second mailbox open in the same Chrome profile (`https://outlook.cloud.microsoft/mail/you@other.com/`), pass `--mailbox-url` so the right tab is picked.

### Import cookies from any browser

Export the cookies for `login.microsoftonline.com` with a cookie-export extension, either as Cookie-Editor JSON or as a Netscape `cookies.txt`:

```bash
outlook-web-cli auth cookies --account work --file ~/Downloads/cookies.json --username me@contoso.com
rm ~/Downloads/cookies.json
```

Only the `login.microsoftonline.com` cookies are kept.

### Check and renew

```bash
outlook-web-cli auth status          # per account: token minutes left, refresh-token hours left, cookie expiry
outlook-web-cli auth refresh --all   # mint fresh tokens now; --rotate also renews the sign-in cookies
outlook-web-cli auth logout --account work
```

## Where things are stored

`~/.config/outlook-web-cli/` (or `$OUTLOOK_WEB_CONFIG_DIR`):

- `config.json` holds the account list: name, aliases, username, and Chrome port.
- `sessions/<account>.json` holds the tokens and cookies, written atomically with mode `0600`. Fields this version doesn't know about are kept when the file is rewritten.

They are credentials: anyone with the file can read that mailbox until the session ends. Keep them out of git and backups you share. Signing out of Outlook web everywhere, or changing your password, ends the session.

`config.json` example with aliases (so `--account contoso.com` or `--account me@contoso.com` also work):

```json
{
  "accounts": [
    { "name": "work", "aliases": ["contoso", "contoso.com"], "username": "me@contoso.com", "cdp_port": 9230 }
  ]
}
```

## Commands

| Command | What it does |
| --- | --- |
| `accounts` | Configured accounts and whether a session is saved |
| `status [--account A\|all]` | Whether each account can get a token, and from where |
| `messages` | Inbox (or `--folder sent\|drafts\|archive\|deleted\|junk`), newest first; `--unread`, `--from`, `--since YYYY-MM-DD`, `--search`, `--limit` |
| `search --query Q` | Full-mailbox search |
| `read --id ID` | One message, with its body as text |
| `agenda` | Calendar for `--today`, `--tomorrow`, or `--start-date`/`--end-date` (America/New_York) |
| `freebusy --emails a@x.com,b@x.com` | When each person is busy and the slots when everyone is free, like Outlook's Scheduling Assistant. Works for coworkers who haven't shared their calendar (you see busy/tentative/out-of-office, plus subjects only if they share more). Same range flags as `agenda` (`--start-date`/`--end-date` also take times, e.g. `2026-11-17T11:00`); `--day-start 09:00 --day-end 17:00`, `--min-minutes 30`, `--weekends`. Up to 20 people; times are America/New_York. The account defaults to the one matching their email domain |
| `attachments --id ID` | List attachments, including inline screenshots; `--include-data` returns base64 (`--offset`/`--length` for big files) |
| `attachment-inspect --id ID` | Describe and transcribe an image attachment with a vision model (`OPENROUTER_API_KEY`) |
| `draft-new`, `draft-reply`, `draft-delete` | Save, reply to (`--reply-all`, `--add-to`, `--add-cc`) or discard drafts. Requires `--yes` |

### Getting help

`outlook-web-cli` (or `--help`, `-h`, `help`) prints one screen: first-time setup, every command grouped by purpose, copy-pasteable examples and the safety rules. Each command has its own page with every flag, accepted values, defaults and examples:

```bash
outlook-web-cli --help               # overview and examples
outlook-web-cli freebusy --help      # one command: every flag, plus examples (same as: outlook-web-cli help freebusy)
outlook-web-cli auth chrome --help   # subcommands too
```

Common tasks:

```bash
outlook-web-cli accounts
outlook-web-cli messages --account work --unread --limit 10
outlook-web-cli search --account work --query "invoice"
outlook-web-cli read --account work --id AAMkAGExample
outlook-web-cli agenda --tomorrow
outlook-web-cli freebusy --emails jane@contoso.com,sam@fabrikam.com --tomorrow
outlook-web-cli draft-reply --account work --id AAMkAGExample --body "Thanks!" --yes
outlook-web-cli auth refresh --all --rotate
```

A missing or invalid argument returns `"ok": false` with a `hint` that names the command's `--help`. Draft commands never send mail and do nothing without `--yes`.

Omit `--account` on read commands to cover every account. Every command takes `--json '{...}'` with the same fields in snake_case, merged over the flags (for example `--json '{"inline_images":[{"url":"https://...","content_id":"shot1"}]}'` on `draft-new`). Output is JSON, and the exit code is 1 when `ok` is false. `outlook-web-cli help COMMAND` shows each command's flags.

Environment settings:

| Variable | What it does |
| --- | --- |
| `OUTLOOK_WEB_CONFIG_DIR` | Where `config.json` and `sessions/` live (default `~/.config/outlook-web-cli`) |
| `OPENROUTER_API_KEY`, `OUTLOOK_WEB_VISION_MODEL` | Image model for `attachment-inspect` (default model `google/gemini-2.5-flash`) |
| `OUTLOOK_INLINE_IMAGE_HOSTS` | Comma-separated hosts that draft inline images may be fetched from (none by default) |

## Tests

```bash
cargo test --manifest-path cli/Cargo.toml
cargo clippy --manifest-path cli/Cargo.toml --all-targets -- -D warnings
cargo fmt --manifest-path cli/Cargo.toml --check
```

The tests use fake Outlook, Microsoft Graph, Microsoft sign-in and Chrome DevTools servers and need no account. `OUTLOOK_TEST_TOKEN` / `OUTLOOK_TEST_TOKEN_<ACCOUNT>`, `OUTLOOK_REST_BASE_URL`, `OUTLOOK_GRAPH_BASE_URL`, `OUTLOOK_WEB_LOGIN_BASE_URL`, `OPENROUTER_BASE_URL` and `OUTLOOK_TEST_NOW_MS` point the code at test servers and a fixed clock. CI runs all three commands on Linux and macOS.

## Contributing

All changes to this repository are made through pull requests, for tracking: create a branch, make a focused change, and open a PR describing the resulting behavior and how it was validated. Nothing is pushed directly to `main`. CI must pass before a PR is merged. Use only placeholder data (contoso.com, fabrikam.com) in tests and examples, never real mailboxes, tokens or cookies.

## History

Rewritten in Rust in v1.0.0 to match the owner's other command-line tools, such as [google-photos-web-cli](https://github.com/andrewfurman/google-photos-web-cli). The commands, flags, JSON output and saved-session format are the same. The Node.js version, including its importable JavaScript library (`configureAccounts`, `commands`, `setVisionAnalyzer`, ...), remains available at tag [`v0.3.0`](https://github.com/andrewfurman/outlook-web-cli/tree/v0.3.0):

```bash
npm install https://github.com/andrewfurman/outlook-web-cli/archive/refs/tags/v0.3.0.tar.gz
```

## Caveats

This is not a Microsoft product. It relies on how Outlook web signs in today (its public client ID, its redirect URI, and the Outlook REST API), which Microsoft can change. Organizations can block it with Conditional Access, such as device compliance or token protection. Use it only with mailboxes you're allowed to access, and within your organization's policies.
