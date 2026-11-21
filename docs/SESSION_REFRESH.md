# Keeping a session alive

`outlook-web-cli` signs in with what Outlook web itself keeps in a browser: an MSAL refresh
token and Microsoft's sign-in cookies. Both are saved per account in
`~/.config/outlook-web-cli/sessions/<account>.json` (or under `$OUTLOOK_WEB_CONFIG_DIR`). This
page explains how long each part lasts, how to keep a session alive on a server or virtual
machine with `auth refresh --all --rotate`, the one deadline that no refresh can move, and how
to sign in again on a headless machine.

Never commit, paste, or print a session file, a cookie export, or a token. The CLI prints
metadata only (`auth status`, `auth refresh`); keep it that way in your scripts and logs.

## 1. What a session is made of

| Part | Where it comes from | Lifetime |
| --- | --- | --- |
| **Access token** for `https://outlook.office.com` | Minted from the refresh token or the cookies | About 60-90 minutes. Reused until 2 minutes before it expires |
| **Refresh token** of the Outlook web single-page app | Outlook web's local storage (import), or minted by a silent sign-in | **24 hours**, fixed, from the sign-in that issued it. Redeeming it returns a new one but does **not** move that 24-hour deadline. The CLI stops using it after 23.5 hours |
| **Sign-in cookies** on `login.microsoftonline.com` (`ESTSAUTHPERSISTENT`, `ESTSAUTH`, ...) | The browser you signed in with ("Stay signed in" ticked) | `ESTSAUTHPERSISTENT` has a **fixed expiry**, typically months out, set when you signed in interactively (see section 3) |

Every command gets a token in this order:

1. the saved access token, if it has more than 2 minutes left;
2. the refresh token, if it is less than 23.5 hours old;
3. a **silent sign-in** with the cookies: the same `prompt=none` authorization-code + PKCE request
   that Outlook web sends from its hidden renewal iframe. It returns a new access token, a new
   refresh token (its 24-hour clock starts again), and `Set-Cookie` headers that the CLI merges
   into the saved cookies;
4. if all of that fails and the account has a remembered Chrome DevTools port
   (`auth chrome --cdp-port N`), a re-import from that Chrome (section 4).

When all four fail, the account reports `needs_sign_in`.

## 2. Keep-alive: `auth refresh --all --rotate`

```bash
outlook-web-cli auth refresh --all --rotate
```

For every configured account, `--rotate` **skips the refresh token and always does the silent
sign-in** (step 3). That:

- renews the sign-in cookies (merging whatever Microsoft sends back), so the saved cookie jar
  stays as current as a browser's;
- issues a fresh refresh token and restarts its 24-hour window, so ordinary commands for the
  next day are served from the refresh token without touching the cookies;
- keeps the session in regular use, like a browser tab that stays open.

Without `--rotate`, `auth refresh` uses the refresh token while it is under 23.5 hours old, and
that does **not** extend its 24-hour deadline. Use `--rotate` for the keep-alive.

If an account fails and has a remembered Chrome port, `auth refresh` tries a re-import from that
Chrome and reports `"reimported_from_chrome": true`. The output is JSON with each account's
`ok`, `via`, and the `auth status` fields (no secrets). The exit code is `1` when any account
failed, so a scheduler can alert on it.

### How often

**Every 6 hours** is a good default:

- It keeps the refresh token far inside its 24-hour window: three more runs would have to fail
  before it lapses. Even then the cookies still work, so a missed run only means the next
  command does a silent sign-in itself.
- It is about as often as an open Outlook web tab renews, without hammering
  `login.microsoftonline.com`. There is no benefit to running it more than about once an hour.
- A random delay of a few minutes avoids every machine renewing at the same moment.

### crontab

```cron
# crontab -e   (user crontab of the account that owns ~/.config/outlook-web-cli)
17 */6 * * * timeout 5m $HOME/.cargo/bin/outlook-web-cli auth refresh --all --rotate >> $HOME/.local/state/outlook-web-cli/keepalive.log 2>&1 || echo "outlook-web-cli keep-alive failed; see docs/SESSION_REFRESH.md section 4" | mail -s "Outlook session needs sign-in" you@example.com
```

Create the log directory first (`mkdir -p ~/.local/state/outlook-web-cli`). The log holds JSON
status lines only (account names, usernames, expiry dates), never tokens or cookies; keep it
private anyway.

### systemd timer

A system unit running as a dedicated user:

```ini
# /etc/systemd/system/outlook-keepalive.service
[Unit]
Description=Renew saved Outlook web sessions (tokens and sign-in cookies)
Wants=network-online.target
After=network-online.target

[Service]
Type=oneshot
User=outlook
Environment=HOME=/home/outlook
ExecStart=/usr/local/bin/outlook-web-cli auth refresh --all --rotate
TimeoutStartSec=300

# /etc/systemd/system/outlook-keepalive.timer
[Unit]
Description=Renew saved Outlook web sessions every 6 hours

[Timer]
OnBootSec=5min
OnUnitActiveSec=6h
RandomizedDelaySec=10min
Persistent=true

[Install]
WantedBy=timers.target
```

```bash
sudo systemctl daemon-reload
sudo systemctl enable --now outlook-keepalive.timer
systemctl list-timers outlook-keepalive.timer
journalctl -u outlook-keepalive.service -n 20
```

For a user unit instead, put the two files in `~/.config/systemd/user/` (drop `User=` and
`Environment=`, use `%h/.cargo/bin/outlook-web-cli`), run
`systemctl --user enable --now outlook-keepalive.timer`, and `loginctl enable-linger $USER` so it
runs without a login session. To be told about failures, add `OnFailure=` pointing at a unit that
sends you a message.

## 3. The deadline no refresh can move

`ESTSAUTHPERSISTENT`, the "Stay signed in" cookie, has a **hard expiry**. A silent sign-in
renews its *value* and the other cookies, but Microsoft keeps the expiry tied to the last
**interactive** sign-in: rotating does not push it out. In practice it is a few months after the
sign-in, and your organization's policies (for example Conditional Access sign-in frequency or
persistent-browser-session settings) can make it shorter or force an earlier sign-in.

So the keep-alive keeps a session working **until** that date, not forever. See it with:

```bash
outlook-web-cli auth status
# per account: "sign_in_cookie_expires": "YYYY-MM-DD", "refresh_token_hours_left", "access_token_valid_min"
```

Put the date in your calendar, or let a check warn you (section 5). Shortly before it, sign in
again interactively (section 4): a fresh interactive sign-in gives a new expiry months out.

The session also ends early, whatever the date, if you change your password, use "Sign out
everywhere", an admin revokes your sessions, or a security policy asks for a new sign-in.

## 4. Signing in again on a headless VM

You need a real browser sign-in (password, MFA, "Stay signed in" ticked). On a machine without a
screen, run Chrome on a virtual display and reach it through noVNC in your own browser.

The pieces (everything bound to localhost except the noVNC web port):

```bash
# packages: xvfb openbox x11vnc websockify novnc, plus Google Chrome
Xvfb :56 -screen 0 1440x900x24 &
DISPLAY=:56 openbox &
DISPLAY=:56 google-chrome \
  --user-data-dir="$HOME/.chrome-outlook" \
  --remote-debugging-address=127.0.0.1 --remote-debugging-port=9230 \
  --no-first-run https://outlook.cloud.microsoft/mail/ &
x11vnc -localhost -nopw -display :56 -rfbport 5956 -forever -shared &
websockify --web /usr/share/novnc 6083 localhost:5956 &
```

1. Open `http://<vm>:6083/vnc.html` through something only you can reach: an SSH tunnel
   (`ssh -L 6083:localhost:6083 <vm>`, then `http://localhost:6083/vnc.html`), a private VPN, or
   your VM provider's authenticated HTTPS proxy. Never expose the noVNC port to the internet: it
   gives control of a browser signed in to your mailbox, and VNC has no password here.
2. Sign in to Outlook web in that Chrome and tick **Stay signed in**. Wait until the inbox loads.
3. Import the session over the DevTools port (stays on localhost):

   ```bash
   outlook-web-cli auth chrome --account work --cdp-port 9230
   outlook-web-cli auth refresh --account work --rotate
   outlook-web-cli auth status --account work     # new sign_in_cookie_expires date
   ```

   For a second mailbox in the same profile, open it in a tab and add
   `--mailbox-url https://outlook.cloud.microsoft/mail/you@other.com/`.

4. Then either:
   - **stop the desktop** (Chrome, x11vnc, websockify, openbox, Xvfb). The CLI's saved session
     is independent of Chrome; the keep-alive keeps it renewed. Stop Chrome by ending the
     process. Do **not** use "Sign out" in Outlook web: that ends the session server-side, for
     the CLI too. Or
   - **leave Chrome running** (the desktop and VNC can still be stopped). The port is
     remembered with the account, so if the saved session ever dies before the cookie's hard
     expiry, `auth refresh` and every other command re-import from that Chrome automatically.
     Outlook web in an open tab renews itself, so it is a working backup until its own sign-in
     expires. It costs a few hundred MB of memory.

### Without a desktop: import cookies from another machine

Sign in to Outlook web in any browser (tick "Stay signed in"), export the cookies for
`login.microsoftonline.com` with a cookie-export extension (Cookie-Editor JSON or a Netscape
`cookies.txt`), copy the file to the VM over SSH, and:

```bash
outlook-web-cli auth cookies --account work --file cookies.json --username you@example.com
shred -u cookies.json 2>/dev/null || rm cookies.json
outlook-web-cli auth refresh --account work --rotate
```

Only the `login.microsoftonline.com` cookies are kept. The first silent sign-in mints the
refresh token. Afterwards, sign that browser out only by closing it, not with "Sign out", or the
imported session ends too.

## 5. Health checks and alerts

- **The keep-alive is the health check.** `auth refresh --all --rotate` exits `1` when any
  account cannot get a token; alert on that (the `|| mail ...` in the crontab above, or
  `OnFailure=` in systemd).
- **A live check from any script:** `outlook-web-cli status` tries to get a token for every
  account. It always exits `0`, so test the JSON:

  ```bash
  outlook-web-cli status | jq -e '.accounts | all(.ok)' >/dev/null || echo "an Outlook account needs sign-in"
  ```

- **Warn before the hard expiry** (section 3), for example daily:

  ```bash
  outlook-web-cli auth status | jq -r --arg soon "$(date -u -d '+14 days' +%F)" \
    '.accounts[] | select(.sign_in_cookie_expires != "" and .sign_in_cookie_expires != "session" and .sign_in_cookie_expires <= $soon)
     | "\(.account): sign-in cookie expires \(.sign_in_cookie_expires); sign in again"'
  ```

  (`date -d` is GNU; on macOS use `date -u -v+14d +%F`.) Any output means it is time to redo
  section 4. `auth status` reads only the saved files and makes no network call.

## 6. Secrets

- `sessions/*.json` holds a refresh token and sign-in cookies: anyone with the file can read the
  mailbox until the session ends. It is written with mode `0600`; keep the config directory out
  of git, shared backups, container images, and support bundles.
- Never print a session file, `cookies.json`/`cookies.txt`, or a token into a terminal you
  share, an issue, a chat, or a log. The CLI's own output contains metadata only.
- Delete cookie exports right after importing them.
- Keep the DevTools port (`--remote-debugging-port`) and VNC on localhost, and stop the noVNC
  desktop when the sign-in is done.
- `outlook-web-cli auth logout --account NAME` deletes the saved session file. It does not sign
  you out of Microsoft; to end the session everywhere, sign out in a browser or change your
  password.
