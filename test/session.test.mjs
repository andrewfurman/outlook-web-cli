import assert from "node:assert/strict";
import crypto from "node:crypto";
import fs from "node:fs";
import http from "node:http";
import os from "node:os";
import path from "node:path";

// A fake login.microsoftonline.com: the silent authorize endpoint and the token endpoint.
const b64 = (o) => Buffer.from(JSON.stringify(o)).toString("base64url");
const jwt = (claims) => `${b64({ alg: "none" })}.${b64(claims)}.sig`;
const TENANT = "11111111-2222-3333-4444-555555555555";
const calls = [];
const codes = new Map();
let authorizeMode = "ok";
let tokenCounter = 0;

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, "http://x");
  let body = "";
  for await (const chunk of req) body += chunk;
  const form = new URLSearchParams(body);
  calls.push({ method: req.method, path: url.pathname, query: url.searchParams, form, headers: req.headers });
  if (url.pathname.endsWith("/oauth2/v2.0/authorize")) {
    const redirect = url.searchParams.get("redirect_uri");
    if (authorizeMode === "login_required" || !/ESTSAUTHPERSISTENT=cookie-v\d/.test(req.headers.cookie || "")) {
      res.writeHead(302, { location: `${redirect}#error=login_required&error_description=AADSTS50058%3A+no+session` });
      return res.end();
    }
    const code = `code-${codes.size + 1}`;
    codes.set(code, url.searchParams.get("code_challenge"));
    const n = Number((req.headers.cookie.match(/ESTSAUTHPERSISTENT=cookie-v(\d+)/) || [])[1]);
    res.writeHead(302, {
      location: `${redirect}#code=${code}&state=${url.searchParams.get("state")}`,
      "set-cookie": [
        `ESTSAUTHPERSISTENT=cookie-v${n + 1}; domain=.login.microsoftonline.com; path=/; expires=Wed, 10 Feb 2027 00:00:00 GMT; secure; HttpOnly`,
        "buid=; domain=login.microsoftonline.com; path=/; max-age=0",
      ],
    });
    return res.end();
  }
  if (url.pathname.endsWith("/oauth2/v2.0/token")) {
    res.setHeader("content-type", "application/json");
    if (req.headers.origin !== "https://outlook.cloud.microsoft") {
      res.writeHead(400);
      return res.end(JSON.stringify({ error: "invalid_request", error_description: "AADSTS9002327: Tokens issued for the 'Single-Page Application' client-type may only be redeemed via cross-origin requests." }));
    }
    const grant = form.get("grant_type");
    if (grant === "authorization_code") {
      const challenge = codes.get(form.get("code"));
      const expected = crypto.createHash("sha256").update(form.get("code_verifier") || "").digest("base64url");
      if (!challenge || challenge !== expected) {
        res.writeHead(400);
        return res.end(JSON.stringify({ error: "invalid_grant", error_description: "PKCE mismatch" }));
      }
    } else if (grant === "refresh_token" && !String(form.get("refresh_token")).startsWith("rt-")) {
      res.writeHead(400);
      return res.end(JSON.stringify({ error: "invalid_grant", error_description: "AADSTS700084: The refresh token was issued to a single page app (SPA), and therefore has a fixed, limited lifetime of 1.00:00:00" }));
    }
    tokenCounter += 1;
    res.writeHead(200);
    return res.end(JSON.stringify({
      token_type: "Bearer",
      expires_in: 3600,
      access_token: jwt({ aud: "https://outlook.office.com", tid: TENANT, scp: "Mail.ReadWrite Calendars.ReadWrite", exp: Math.floor(Date.now() / 1000) + 3600, n: tokenCounter }),
      refresh_token: `rt-${tokenCounter}`,
      id_token: jwt({ tid: TENANT, oid: "oid-1", preferred_username: "me@contoso.com" }),
    }));
  }
  res.writeHead(404);
  res.end("{}");
});
await new Promise((r) => server.listen(0, "127.0.0.1", r));
process.env.OUTLOOK_WEB_LOGIN_BASE_URL = `http://127.0.0.1:${server.address().port}`;
process.env.OUTLOOK_WEB_CONFIG_DIR = fs.mkdtempSync(path.join(os.tmpdir(), "outlook-web-session-"));

const session = await import("../src/session.mjs");
const { configureAccounts } = await import("../src/accounts.mjs");
const { outlookStatus } = await import("../src/outlook.mjs");
configureAccounts([{ name: "contoso", username: "me@contoso.com" }]);

try {
  // 1. Import cookies exported from a browser (Cookie-Editor JSON); only Microsoft sign-in cookies are kept.
  const exported = JSON.stringify([
    { name: "ESTSAUTHPERSISTENT", value: "cookie-v1", domain: ".login.microsoftonline.com", path: "/", expirationDate: 1802000000 },
    { name: "buid", value: "b", domain: "login.microsoftonline.com", path: "/" },
    { name: "SID", value: "google", domain: ".google.com", path: "/" },
  ]);
  const imported = session.importCookies("contoso", exported, { username: "me@contoso.com" });
  assert.equal(imported.ok, true);
  assert.equal(imported.imported_cookies, 2);
  const file = path.join(process.env.OUTLOOK_WEB_CONFIG_DIR, "sessions", "contoso.json");
  assert.equal(fs.statSync(file).mode & 0o777, 0o600, "session file is private");

  // 2. Cookies -> silent sign-in (prompt=none + PKCE) -> token; cookies rotate, tenant and RT are saved.
  const first = await session.refreshSession("contoso");
  assert.equal(first.ok, true, JSON.stringify(first));
  assert.equal(first.via, "cookies");
  const authorize = calls.find((c) => c.path.endsWith("/authorize"));
  assert.equal(authorize.query.get("prompt"), "none");
  assert.equal(authorize.query.get("code_challenge_method"), "S256");
  assert.equal(authorize.query.get("login_hint"), "me@contoso.com");
  assert.match(authorize.path, /^\/organizations\//, "first sign-in uses the organizations endpoint");
  let saved = session.loadSession("contoso");
  assert.equal(saved.tenant_id, TENANT);
  assert.equal(saved.refresh_token, "rt-1");
  assert.equal(saved.cookies.find((c) => c.name === "ESTSAUTHPERSISTENT").value, "cookie-v2", "rotated cookie saved");
  assert.equal(saved.cookies.find((c) => c.name === "buid"), undefined, "expired cookie dropped");

  // 3. A valid saved access token is reused without calling Microsoft.
  calls.length = 0;
  const reused = await session.getSessionToken("contoso");
  assert.equal(reused.source, "saved");
  assert.equal(reused.api, "outlook_rest");
  assert.equal(calls.length, 0);

  // 4. An expired access token is renewed with the refresh token (Origin header required).
  saved = session.loadSession("contoso");
  session.saveSession("contoso", { ...saved, access_token: { ...saved.access_token, expires_on: 1 } });
  const renewed = await session.getSessionToken("contoso");
  assert.equal(renewed.source, "refresh_token");
  assert.equal(calls.at(-1).form.get("grant_type"), "refresh_token");
  assert.equal(calls.at(-1).form.get("refresh_token"), "rt-1");
  assert.match(calls.at(-1).path, new RegExp(`^/${TENANT}/`));
  assert.equal(session.loadSession("contoso").refresh_token, "rt-2");

  // 5. A rejected refresh token falls back to the cookies.
  saved = session.loadSession("contoso");
  session.saveSession("contoso", { ...saved, refresh_token: "expired", access_token: null });
  const fallback = await session.refreshSession("contoso");
  assert.equal(fallback.ok, true);
  assert.equal(fallback.via, "cookies");

  // 6. A day-old refresh token is not tried at all.
  saved = session.loadSession("contoso");
  session.saveSession("contoso", { ...saved, refresh_token_obtained_at: new Date(Date.now() - 25 * 3600e3).toISOString() });
  calls.length = 0;
  assert.equal((await session.refreshSession("contoso")).via, "cookies");
  assert.ok(!calls.some((c) => c.form.get("grant_type") === "refresh_token"));

  // 7. --rotate (force: "cookies") always renews the cookies.
  const before = session.loadSession("contoso").cookies.find((c) => c.name === "ESTSAUTHPERSISTENT").value;
  assert.equal((await session.refreshSession("contoso", { force: "cookies" })).via, "cookies");
  assert.notEqual(session.loadSession("contoso").cookies.find((c) => c.name === "ESTSAUTHPERSISTENT").value, before);

  // 8. Status shows lifetimes, never secrets; the outlook layer reports where the token came from.
  const status = session.sessionStatus("contoso");
  assert.equal(status.saved, true);
  assert.equal(status.sign_in_cookie, true);
  assert.equal(status.sign_in_cookie_expires, "2027-02-10");
  assert.ok(status.access_token_valid_min > 50);
  assert.ok(!/rt-|cookie-v|eyJ/.test(JSON.stringify(status)), "no secrets in status");
  const live = await outlookStatus({ account: "contoso" });
  assert.equal(live.accounts[0].status, "signed_in");
  assert.equal(live.accounts[0].token_source, "saved");

  // 9. A signed-out session reports needs_sign_in with the reason.
  authorizeMode = "login_required";
  saved = session.loadSession("contoso");
  session.saveSession("contoso", { ...saved, refresh_token: "", access_token: null });
  const dead = await session.refreshSession("contoso");
  assert.equal(dead.ok, false);
  assert.equal(dead.status, "needs_sign_in");
  assert.match(dead.errors.join(" "), /login_required/);
  assert.equal(await session.getSessionToken("contoso"), null);

  // 10. Netscape cookies.txt exports work too; logout removes the session.
  const txt = "# Netscape HTTP Cookie File\n#HttpOnly_.login.microsoftonline.com\tTRUE\t/\tTRUE\t1802000000\tESTSAUTHPERSISTENT\tcookie-v9\n";
  assert.deepEqual(session.parseCookieExport(txt).map((c) => [c.name, c.value, c.domain]), [["ESTSAUTHPERSISTENT", "cookie-v9", ".login.microsoftonline.com"]]);
  assert.equal(session.importCookies("contoso", "[]").status, "no_sign_in_cookie");
  assert.equal(session.deleteSession("contoso"), true);
  assert.equal(session.sessionStatus("contoso").saved, false);

  console.log("session tests: ok");
} finally {
  server.close();
}
