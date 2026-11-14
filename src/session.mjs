import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import { getAccount, listAccounts, sessionsDir } from "./accounts.mjs";
import { connectCdp, findOutlookTab } from "./cdp.mjs";

// Saved Outlook web sessions, used like a browser would: no Chrome needed after import.
//
// A session file (sessions/<account>.json, mode 0600) holds what Outlook web itself keeps:
// - the MSAL refresh token of the Outlook web app (a single-page app, so the token is
//   redeemed with Outlook's Origin header and lasts 24 hours from sign-in),
// - the login.microsoftonline.com sign-in cookies (ESTSAUTHPERSISTENT and friends, which
//   last months when "Stay signed in" was chosen),
// - the current access token for https://outlook.office.com and its expiry.
//
// Getting a token tries, in order: the saved access token, the refresh token, then a
// silent sign-in with the cookies (prompt=none + PKCE, as Outlook web's own hidden
// iframe does), which also yields a fresh refresh token and rotated cookies. When the
// account has a Chrome DevTools port, a dead session is re-imported from that Chrome.
// Token and cookie values are never logged or returned by status calls.

export const OUTLOOK_WEB_CLIENT_ID = "9199bf20-a13f-4107-85dc-02114787ef48";
export const OUTLOOK_WEB_ORIGIN = "https://outlook.cloud.microsoft";
export const OUTLOOK_WEB_REDIRECT_URI = "https://outlook.cloud.microsoft/mail/";
export const OUTLOOK_SCOPE = "https://outlook.office.com/.default openid profile offline_access";

const ACCESS_TOKEN_MARGIN_SEC = 120;
// Single-page-app refresh tokens expire 24 hours after the sign-in that issued them.
const REFRESH_TOKEN_MAX_AGE_MS = 23.5 * 60 * 60 * 1000;
const LOGIN_TIMEOUT_MS = 15000;
const BROWSER_UA =
  "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/141.0.0.0 Safari/537.36";
const inFlight = new Map();

function loginBase() {
  return (process.env.OUTLOOK_WEB_LOGIN_BASE_URL || "https://login.microsoftonline.com").replace(/\/+$/, "");
}

function sessionPath(name) {
  if (!/^[a-z0-9][a-z0-9_-]*$/.test(String(name))) throw new Error(`Invalid account name '${name}'.`);
  return path.join(sessionsDir(), `${name}.json`);
}

export function loadSession(name) {
  try {
    return JSON.parse(fs.readFileSync(sessionPath(name), "utf8"));
  } catch {
    return null;
  }
}

export function saveSession(name, session) {
  const file = sessionPath(name);
  fs.mkdirSync(path.dirname(file), { recursive: true, mode: 0o700 });
  const tmp = `${file}.${process.pid}.${Date.now()}.tmp`;
  fs.writeFileSync(tmp, JSON.stringify({ ...session, version: 1, account: name, updated_at: new Date().toISOString() }, null, 2), { mode: 0o600 });
  fs.renameSync(tmp, file);
}

export function deleteSession(name) {
  try {
    fs.unlinkSync(sessionPath(name));
    return true;
  } catch {
    return false;
  }
}

// Metadata only: safe to print.
export function sessionStatus(name) {
  const s = loadSession(name);
  if (!s) return { account: name, saved: false };
  const now = Date.now();
  const persistent = (s.cookies || []).find((c) => c.name === "ESTSAUTHPERSISTENT");
  const rtAge = s.refresh_token_obtained_at ? now - Date.parse(s.refresh_token_obtained_at) : null;
  return {
    account: name,
    saved: true,
    username: s.username || "",
    tenant_id: s.tenant_id || "",
    source: s.source || "",
    access_token_valid_min: s.access_token?.expires_on ? Math.max(0, Math.round((s.access_token.expires_on * 1000 - now) / 60000)) : 0,
    refresh_token: Boolean(s.refresh_token),
    refresh_token_hours_left: rtAge == null || !s.refresh_token ? 0 : Math.max(0, Math.round((REFRESH_TOKEN_MAX_AGE_MS - rtAge) / 360000) / 10),
    sign_in_cookie: Boolean(persistent),
    sign_in_cookie_expires: persistent?.expires ? new Date(persistent.expires * 1000).toISOString().slice(0, 10) : persistent ? "session" : "",
    updated_at: s.updated_at || "",
  };
}

// Returns { token, api: "outlook_rest", source } or null. One refresh per account at a time.
export async function getSessionToken(name, { allowChromeImport = true } = {}) {
  const key = `${name}:${allowChromeImport}`;
  if (inFlight.has(key)) return inFlight.get(key);
  const p = getSessionTokenUnlocked(name, { allowChromeImport }).finally(() => inFlight.delete(key));
  inFlight.set(key, p);
  return p;
}

async function getSessionTokenUnlocked(name, { allowChromeImport }) {
  const session = loadSession(name);
  if (session) {
    const nowSec = Date.now() / 1000;
    if (session.access_token?.token && session.access_token.expires_on > nowSec + ACCESS_TOKEN_MARGIN_SEC) {
      return { token: session.access_token.token, api: "outlook_rest", source: "saved" };
    }
    const fresh = await refreshSession(name, { session });
    if (fresh.ok) return { token: fresh.token, api: "outlook_rest", source: fresh.via };
  }
  const account = getAccount(name);
  if (allowChromeImport && account?.cdpPort) {
    const imported = await importFromChrome(name, { cdpPort: account.cdpPort, mainUrl: account.mainUrl });
    if (imported.ok) {
      const fresh = await refreshSession(name);
      if (fresh.ok) return { token: fresh.token, api: "outlook_rest", source: `chrome+${fresh.via}` };
    }
  }
  return null;
}

// Mints a new access token from the refresh token, else from the sign-in cookies.
// force: "cookies" skips the refresh token (the keepalive uses it to rotate cookies).
export async function refreshSession(name, { session = loadSession(name), force = "" } = {}) {
  if (!session) return { ok: false, account: name, status: "no_session" };
  const errors = [];
  const rtFresh = session.refresh_token && session.refresh_token_obtained_at
    && Date.now() - Date.parse(session.refresh_token_obtained_at) < REFRESH_TOKEN_MAX_AGE_MS;
  if (rtFresh && force !== "cookies") {
    const r = await redeem(session, { grant_type: "refresh_token", refresh_token: session.refresh_token });
    if (r.ok) return finish(name, session, r.json, "refresh_token", { keepRefreshTokenAge: true });
    errors.push(`refresh_token: ${r.error}`);
  }
  if ((session.cookies || []).some((c) => c.name === "ESTSAUTHPERSISTENT" || c.name === "ESTSAUTH")) {
    const r = await silentSignIn(session);
    if (r.ok) return finish(name, { ...session, cookies: r.cookies }, r.json, "cookies", { keepRefreshTokenAge: false });
    if (r.cookies) saveSession(name, { ...session, cookies: r.cookies });
    errors.push(`cookies: ${r.error}`);
  }
  if (!errors.length) errors.push("nothing usable saved");
  return { ok: false, account: name, status: "needs_sign_in", errors };
}

function finish(name, session, json, via, { keepRefreshTokenAge }) {
  const claims = decodeJwt(json.access_token);
  const id = decodeJwt(json.id_token);
  const next = {
    ...session,
    tenant_id: id.tid || claims.tid || session.tenant_id,
    home_account_id: id.oid && id.tid ? `${id.oid}.${id.tid}` : session.home_account_id,
    username: id.preferred_username || claims.upn || claims.unique_name || session.username || "",
    refresh_token: json.refresh_token || session.refresh_token,
    refresh_token_obtained_at: json.refresh_token
      ? (keepRefreshTokenAge && session.refresh_token_obtained_at ? session.refresh_token_obtained_at : new Date().toISOString())
      : session.refresh_token_obtained_at,
    access_token: {
      token: json.access_token,
      expires_on: Number(claims.exp) || Math.floor(Date.now() / 1000) + Number(json.expires_in || 3600),
      audience: String(claims.aud || ""),
    },
  };
  saveSession(name, next);
  return { ok: true, account: name, token: json.access_token, via };
}

async function redeem(session, params) {
  const tenant = session.tenant_id || "organizations";
  try {
    const res = await fetch(`${loginBase()}/${tenant}/oauth2/v2.0/token`, {
      method: "POST",
      headers: { origin: session.origin || OUTLOOK_WEB_ORIGIN, "content-type": "application/x-www-form-urlencoded", "user-agent": BROWSER_UA },
      body: new URLSearchParams({ client_id: session.client_id || OUTLOOK_WEB_CLIENT_ID, scope: OUTLOOK_SCOPE, ...params }),
      redirect: "error",
      signal: AbortSignal.timeout(LOGIN_TIMEOUT_MS),
    });
    const json = await res.json().catch(() => ({}));
    if (res.ok && json.access_token) return { ok: true, json };
    return { ok: false, error: `${json.error || `http_${res.status}`} ${firstLine(json.error_description)}`.trim() };
  } catch (err) {
    return { ok: false, error: String(err?.message || err) };
  }
}

// prompt=none authorization-code + PKCE flow with the saved sign-in cookies, exactly as
// Outlook web renews in a hidden iframe. Returns rotated cookies either way.
async function silentSignIn(session) {
  const tenant = session.tenant_id || "organizations";
  const redirectUri = session.redirect_uri || OUTLOOK_WEB_REDIRECT_URI;
  const verifier = crypto.randomBytes(32).toString("base64url");
  const query = new URLSearchParams({
    client_id: session.client_id || OUTLOOK_WEB_CLIENT_ID,
    response_type: "code",
    redirect_uri: redirectUri,
    scope: OUTLOOK_SCOPE,
    prompt: "none",
    response_mode: "fragment",
    code_challenge: crypto.createHash("sha256").update(verifier).digest("base64url"),
    code_challenge_method: "S256",
    state: crypto.randomBytes(8).toString("hex"),
    nonce: crypto.randomBytes(8).toString("hex"),
  });
  if (session.username) query.set("login_hint", session.username);
  let res;
  try {
    res = await fetch(`${loginBase()}/${tenant}/oauth2/v2.0/authorize?${query}`, {
      headers: { cookie: cookieHeader(session.cookies), "user-agent": BROWSER_UA },
      redirect: "manual",
      signal: AbortSignal.timeout(LOGIN_TIMEOUT_MS),
    });
  } catch (err) {
    return { ok: false, error: String(err?.message || err) };
  }
  const cookies = mergeSetCookies(session.cookies, res.headers.getSetCookie?.() || []);
  const location = res.headers.get("location") || "";
  const params = new URLSearchParams(location.includes("#") ? location.split("#")[1] : location.split("?")[1] || "");
  const code = params.get("code");
  if (!code) {
    const err = params.get("error") || `http_${res.status}`;
    return { ok: false, cookies, error: `${err} ${firstLine(params.get("error_description"))}`.trim() };
  }
  const r = await redeem(session, { grant_type: "authorization_code", code, redirect_uri: redirectUri, code_verifier: verifier });
  return r.ok ? { ok: true, json: r.json, cookies } : { ok: false, cookies, error: r.error };
}

// Reads the refresh token and sign-in cookies from a signed-in Outlook web tab.
export async function importFromChrome(name, { cdpPort, mainUrl } = {}) {
  const tab = await findOutlookTab(cdpPort, mainUrl);
  if (!tab) return { ok: false, account: name, status: "no_outlook_tab", message: `No Outlook web tab on Chrome DevTools port ${cdpPort}.` };
  const cdp = await connectCdp(tab.webSocketDebuggerUrl);
  if (!cdp) return { ok: false, account: name, status: "cdp_unreachable" };
  try {
    const msal = await cdp.evaluate(`(() => {
      const out = { refresh: [], accounts: [] };
      for (let i = 0; i < localStorage.length; i++) {
        const k = localStorage.key(i) || "";
        if (!k.startsWith("msal.")) continue;
        let o; try { o = JSON.parse(localStorage.getItem(k) || ""); } catch { continue; }
        if (!o || typeof o !== "object") continue;
        if (o.credentialType === "RefreshToken" && typeof o.secret === "string") out.refresh.push({ secret: o.secret, clientId: o.clientId, homeAccountId: o.homeAccountId, environment: o.environment });
        else if (o.homeAccountId && o.username && o.realm) out.accounts.push({ homeAccountId: o.homeAccountId, username: o.username, realm: o.realm });
      }
      return out;
    })()`);
    const all = await cdp.send("Network.getAllCookies");
    const cookies = (all.cookies || [])
      .filter((c) => /(^|\.)login\.microsoftonline\.com$/.test(c.domain))
      .map((c) => ({ name: c.name, value: c.value, domain: c.domain, path: c.path || "/", expires: c.expires > 0 ? Math.floor(c.expires) : null }));
    const refresh = (msal?.refresh || []).find((r) => r.clientId === OUTLOOK_WEB_CLIENT_ID) || (msal?.refresh || [])[0];
    const acct = (msal?.accounts || []).find((a) => a.homeAccountId === refresh?.homeAccountId) || (msal?.accounts || [])[0];
    if (!refresh && !cookies.some((c) => c.name.startsWith("ESTSAUTH"))) {
      return { ok: false, account: name, status: "not_signed_in", message: "The Outlook tab has no saved sign-in. Sign in to Outlook web in that Chrome first." };
    }
    const previous = loadSession(name) || {};
    saveSession(name, {
      ...previous,
      source: `chrome:${cdpPort}`,
      client_id: refresh?.clientId || OUTLOOK_WEB_CLIENT_ID,
      origin: originOf(tab.url),
      redirect_uri: OUTLOOK_WEB_REDIRECT_URI,
      tenant_id: acct?.realm || String(refresh?.homeAccountId || "").split(".")[1] || previous.tenant_id || "",
      home_account_id: refresh?.homeAccountId || previous.home_account_id || "",
      username: acct?.username || previous.username || "",
      refresh_token: refresh?.secret || "",
      // Outlook web renews its refresh token often, so a freshly read one is at most a day old;
      // count it as new and let the cookie sign-in take over if Microsoft rejects it.
      refresh_token_obtained_at: refresh?.secret ? new Date().toISOString() : null,
      cookies,
      access_token: null,
    });
    return { ok: true, account: name, ...sessionStatus(name) };
  } finally {
    cdp.close();
  }
}

// Imports cookies exported from a browser (Cookie-Editor / EditThisCookie JSON, or a
// Netscape cookies.txt). Only login.microsoftonline.com cookies are kept.
export function importCookies(name, text, { username = "", tenant = "" } = {}) {
  const cookies = parseCookieExport(text).filter((c) => /(^|\.)login\.microsoftonline\.com$/.test(c.domain));
  if (!cookies.some((c) => c.name === "ESTSAUTHPERSISTENT" || c.name === "ESTSAUTH")) {
    return { ok: false, account: name, status: "no_sign_in_cookie", message: "No ESTSAUTHPERSISTENT or ESTSAUTH cookie for login.microsoftonline.com in that export." };
  }
  const previous = loadSession(name) || {};
  saveSession(name, {
    ...previous,
    source: "cookies",
    client_id: OUTLOOK_WEB_CLIENT_ID,
    origin: OUTLOOK_WEB_ORIGIN,
    redirect_uri: OUTLOOK_WEB_REDIRECT_URI,
    tenant_id: tenant || previous.tenant_id || "",
    username: username || previous.username || "",
    refresh_token: "",
    refresh_token_obtained_at: null,
    cookies,
    access_token: null,
  });
  return { ok: true, account: name, imported_cookies: cookies.length };
}

export function parseCookieExport(text) {
  const s = String(text || "").trim();
  if (s.startsWith("[") || s.startsWith("{")) {
    const parsed = JSON.parse(s);
    const list = Array.isArray(parsed) ? parsed : parsed.cookies || [];
    return list.map((c) => ({
      name: String(c.name || ""),
      value: String(c.value || ""),
      domain: String(c.domain || "").toLowerCase(),
      path: String(c.path || "/"),
      expires: Number(c.expirationDate || c.expires || 0) > 0 ? Math.floor(Number(c.expirationDate || c.expires)) : null,
    })).filter((c) => c.name);
  }
  return s.split(/\r?\n/)
    .map((line) => line.replace(/^#HttpOnly_/, ""))
    .filter((line) => line && !line.startsWith("#"))
    .map((line) => line.split("\t"))
    .filter((f) => f.length >= 7)
    .map(([domain, , cpath, , expires, cname, value]) => ({
      name: cname,
      value,
      domain: domain.toLowerCase(),
      path: cpath || "/",
      expires: Number(expires) > 0 ? Number(expires) : null,
    }));
}

function cookieHeader(cookies) {
  const nowSec = Date.now() / 1000;
  return (cookies || [])
    .filter((c) => c.value && (!c.expires || c.expires > nowSec))
    .map((c) => `${c.name}=${c.value}`)
    .join("; ");
}

export function mergeSetCookies(existing, setCookieHeaders) {
  const jar = new Map((existing || []).map((c) => [`${c.name}|${c.domain}`, { ...c }]));
  for (const header of setCookieHeaders) {
    const [pair, ...attrs] = String(header).split(";");
    const eq = pair.indexOf("=");
    if (eq <= 0) continue;
    const name = pair.slice(0, eq).trim();
    const value = pair.slice(eq + 1).trim();
    const attr = Object.fromEntries(attrs.map((a) => {
      const i = a.indexOf("=");
      return i < 0 ? [a.trim().toLowerCase(), true] : [a.slice(0, i).trim().toLowerCase(), a.slice(i + 1).trim()];
    }));
    const domain = String(attr.domain || "login.microsoftonline.com").toLowerCase();
    let expires = null;
    if (attr["max-age"] !== undefined) expires = Math.floor(Date.now() / 1000) + Number(attr["max-age"]);
    else if (attr.expires) expires = Math.floor(Date.parse(attr.expires) / 1000) || null;
    // Same name on the bare or dotted domain is the same cookie for our purposes.
    for (const k of [...jar.keys()]) {
      if (k.startsWith(`${name}|`) && k.split("|")[1].replace(/^\./, "") === domain.replace(/^\./, "")) jar.delete(k);
    }
    if (!value || (expires !== null && expires <= Date.now() / 1000)) continue;
    jar.set(`${name}|${domain}`, { name, value, domain, path: String(attr.path || "/"), expires });
  }
  return [...jar.values()];
}

export function decodeJwt(token) {
  try {
    return JSON.parse(Buffer.from(String(token).split(".")[1], "base64url").toString("utf8"));
  } catch {
    return {};
  }
}

function originOf(url) {
  try {
    return new URL(url).origin;
  } catch {
    return OUTLOOK_WEB_ORIGIN;
  }
}

function firstLine(s) {
  return String(s || "").split(/\r?\n/)[0].slice(0, 200);
}

// Renews every account (or the named ones): a fresh token, and with rotate also renewed
// sign-in cookies. An account whose session is dead is re-imported from its Chrome port.
// Safe to print: no tokens or cookies in the result. Run from cron as a keepalive.
export async function refreshAccounts({ accounts, rotate = false } = {}) {
  const names = accounts?.length ? accounts : listAccounts().map((a) => a.name);
  const results = [];
  for (const name of names) {
    let r = await refreshSession(name, { force: rotate ? "cookies" : "" });
    const account = getAccount(name);
    if (!r.ok && account?.cdpPort) {
      const imported = await importFromChrome(name, { cdpPort: account.cdpPort, mainUrl: account.mainUrl });
      if (imported.ok) r = { ...(await refreshSession(name)), reimported_from_chrome: true };
    }
    const { token, ...safe } = r;
    results.push({ ...safe, ...sessionStatus(name) });
  }
  return { ok: results.length > 0 && results.every((r) => r.ok), status: "ok", accounts: results };
}
