# outlook-web-cli

Unofficial command-line access to Outlook (Microsoft 365) mail and calendar, using your **Outlook web session** instead of an Azure app registration. If you can open Outlook in a browser, you can use it from the terminal, from scripts, and from AI agents. No admin consent and no client secret.

It reads mail, searches, reads the calendar and attachments, and saves drafts. **It never sends mail**: there is no send code path, and draft commands do nothing without `--yes`.

```bash
outlook-web auth chrome --account work --cdp-port 9230
outlook-web messages --account work --unread --limit 10
outlook-web read --account work --id AAMkAG...
outlook-web agenda --today
outlook-web draft-reply --account work --id AAMkAG... --reply-all --body "Thanks, done." --yes
```

## Install

Needs Node.js 22 or newer.

```bash
npm install -g https://github.com/andrewfurman/outlook-web-cli/archive/refs/tags/v0.1.0.tar.gz
outlook-web help
```

Or from a checkout: `npm install && npm link`.

## Signing in: how the web session works

Outlook web is a single-page app that signs in with Microsoft's MSAL library. Two things in your browser keep you signed in, and `outlook-web` saves and uses both, the same way the web app does:

| What | Lifetime | How outlook-web uses it |
| --- | --- | --- |
| Outlook web's **refresh token** (in the page's local storage) | 24 hours from sign-in (Microsoft's fixed limit for single-page apps) | Redeemed at `login.microsoftonline.com` with Outlook's `Origin`, for a 60–90 minute access token |
| Microsoft's **sign-in cookies** (`ESTSAUTHPERSISTENT` on `login.microsoftonline.com`) | Months, when "Stay signed in" was chosen | A silent sign-in (`prompt=none` + PKCE), exactly like Outlook web's hidden renewal iframe, gives a new refresh token and renewed cookies |

The access token is for `https://outlook.office.com` with the web app's own `Mail.ReadWrite` and `Calendars` scopes, and calls go to the Outlook REST API. Each call uses the saved access token while it's valid, then the refresh token, then the cookies. Run `outlook-web auth refresh --all --rotate` from cron every few hours to keep the cookies renewed.

### Import from Chrome

Start Chrome with a DevTools port, sign in to Outlook web there (tick "Stay signed in"), then:

```bash
google-chrome --remote-debugging-port=9230 --user-data-dir="$HOME/.chrome-outlook" https://outlook.cloud.microsoft/mail/
outlook-web auth chrome --account work --cdp-port 9230
```

The port is remembered. If the saved session ever stops working, `outlook-web` re-imports from that Chrome automatically. If that Chrome isn't running, it goes straight to "needs sign-in".

For a second mailbox open in the same Chrome profile (`https://outlook.cloud.microsoft/mail/you@other.com/`), pass `--mailbox-url` so the right tab is picked.

### Import cookies from any browser

Export the cookies for `login.microsoftonline.com` with a cookie-export extension, either as Cookie-Editor JSON or as a Netscape `cookies.txt`:

```bash
outlook-web auth cookies --account work --file ~/Downloads/cookies.json --username me@contoso.com
rm ~/Downloads/cookies.json
```

Only the `login.microsoftonline.com` cookies are kept.

### Check and renew

```bash
outlook-web auth status          # per account: token minutes left, refresh-token hours left, cookie expiry
outlook-web auth refresh --all   # mint fresh tokens now; --rotate also renews the sign-in cookies
outlook-web auth logout --account work
```

## Where things are stored

`~/.config/outlook-web-cli/` (or `$OUTLOOK_WEB_CONFIG_DIR`):

- `config.json` holds the account list: name, aliases, username, and Chrome port.
- `sessions/<account>.json` holds the tokens and cookies, written atomically with mode `0600`.

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
| `attachments --id ID` | List attachments, including inline screenshots; `--include-data` returns base64 (`--offset`/`--length` for big files) |
| `attachment-inspect --id ID` | Describe and transcribe an image attachment with a vision model (`OPENROUTER_API_KEY`) |
| `draft-new`, `draft-reply`, `draft-delete` | Save, reply to (`--reply-all`, `--add-to`, `--add-cc`) or discard drafts. Requires `--yes` |

Omit `--account` on read commands to cover every account. Every command takes `--json '{...}'` with the same fields in snake_case. Output is JSON, and the exit code is 1 when `ok` is false.

## Use as a library

```js
import { configureAccounts, commands, outlookMessages, setVisionAnalyzer } from "outlook-web-cli";

configureAccounts([{ name: "work", aliases: ["contoso.com"], cdp_port: 9230 }]);
const inbox = await outlookMessages({ account: "work", unread: true, limit: 5 });
const draft = await commands["draft-new"]({ account: "work", to: "a@b.com", subject: "Hi", body: "..." }, { confirmed: true });
```

[PhoneClaw](https://github.com/andrewfurman/phone-claw) uses it this way for voice access to work mail.

## Tests

```bash
npm test
```

The tests use fake Outlook and Microsoft sign-in servers and need no account. `OUTLOOK_TEST_TOKEN` / `OUTLOOK_TEST_TOKEN_<ACCOUNT>`, `OUTLOOK_REST_BASE_URL`, `OUTLOOK_GRAPH_BASE_URL` and `OUTLOOK_WEB_LOGIN_BASE_URL` point the code at test servers.

## Caveats

This is not a Microsoft product. It relies on how Outlook web signs in today (its public client ID, its redirect URI, and the Outlook REST API), which Microsoft can change. Organizations can block it with Conditional Access, such as device compliance or token protection. Use it only with mailboxes you're allowed to access, and within your organization's policies.
