import fs from "node:fs";
import os from "node:os";
import path from "node:path";

// Accounts are named mailboxes, e.g. "work". Each one may list aliases (other names or
// email domains that pick it), the signed-in username, and, optionally, a Chrome DevTools
// port where Outlook web is open, used to import or re-import the web session.
//
// They come from configureAccounts() (for programs that embed this package), or else
// from $OUTLOOK_WEB_CONFIG_DIR/config.json (default ~/.config/outlook-web-cli/config.json):
//
//   { "accounts": [ { "name": "work", "aliases": ["contoso.com"], "username": "me@contoso.com",
//                     "cdp_port": 9230, "mailbox_url": "https://outlook.cloud.microsoft/mail/" } ] }
//
// Saved sessions alone also count as accounts, so `outlook-web auth chrome --account work`
// is enough to start without writing a config file.

export const DEFAULT_MAILBOX_URL = "https://outlook.cloud.microsoft/mail/";

let configured = null;

export function configDir() {
  return process.env.OUTLOOK_WEB_CONFIG_DIR || path.join(os.homedir(), ".config", "outlook-web-cli");
}

export function sessionsDir() {
  return path.join(configDir(), "sessions");
}

// Replaces the account list (programs that embed this package call this once at start-up).
export function configureAccounts(list) {
  configured = (list || []).map(normalizeAccount).filter(Boolean);
  return configured;
}

export function listAccounts() {
  if (configured) return configured;
  const fromFile = readConfigFile().map(normalizeAccount).filter(Boolean);
  const names = new Set(fromFile.map((a) => a.name));
  let saved = [];
  try {
    saved = fs.readdirSync(sessionsDir())
      .filter((f) => f.endsWith(".json"))
      .map((f) => f.slice(0, -5))
      .filter((n) => !names.has(n))
      .map((name) => normalizeAccount({ name }));
  } catch {
    // no sessions yet
  }
  return [...fromFile, ...saved];
}

// Maps a typed or spoken name, alias, email address or domain to an account name.
// Returns null for empty input, "all", or anything that is not a configured account.
export function resolveAccountAlias(value) {
  const raw = typeof value === "string" ? value.trim().toLowerCase() : "";
  if (!raw) return null;
  const squash = (s) => String(s || "").toLowerCase().replace(/[^a-z0-9]/g, "");
  const candidates = [squash(raw), squash(raw.replace(/^.*@/, ""))];
  for (const account of listAccounts()) {
    const keys = [account.name, ...account.aliases, account.username, account.username.replace(/^.*@/, "")]
      .map(squash)
      .filter(Boolean);
    if (candidates.some((c) => keys.includes(c))) return account.name;
  }
  return null;
}

export function getAccount(name) {
  return listAccounts().find((a) => a.name === name) || null;
}

function normalizeAccount(raw) {
  if (!raw || typeof raw !== "object") return null;
  const name = String(raw.name || "").trim().toLowerCase();
  if (!/^[a-z0-9][a-z0-9_-]*$/.test(name)) return null;
  const cdpPort = Number(raw.cdp_port ?? raw.cdpPort ?? 0);
  return {
    name,
    label: name,
    aliases: Array.isArray(raw.aliases) ? raw.aliases.map(String) : [],
    username: String(raw.username || ""),
    emailDomain: String(raw.email_domain ?? raw.emailDomain ?? (String(raw.username || "").split("@")[1] || "")),
    access: String(raw.access || "read + drafts"),
    cdpPort: Number.isInteger(cdpPort) && cdpPort > 0 ? cdpPort : null,
    mainUrl: String(raw.mailbox_url ?? raw.mainUrl ?? DEFAULT_MAILBOX_URL),
  };
}

function readConfigFile() {
  try {
    const parsed = JSON.parse(fs.readFileSync(path.join(configDir(), "config.json"), "utf8"));
    return Array.isArray(parsed?.accounts) ? parsed.accounts : [];
  } catch {
    return [];
  }
}
