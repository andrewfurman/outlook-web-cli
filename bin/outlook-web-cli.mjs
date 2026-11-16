#!/usr/bin/env node
import fs from "node:fs";
import path from "node:path";
import {
  commands,
  configDir,
  deleteSession,
  getAccount,
  importCookies,
  importFromChrome,
  listAccounts,
  refreshAccounts,
  refreshSession,
  resolveAccountAlias,
  sessionStatus,
  setVisionAnalyzer,
} from "../src/index.mjs";
import { openRouterVisionAnalyzer } from "../src/vision.mjs";

const HELP = `outlook-web-cli: Outlook mail and calendar from the command line, through your Outlook web session.

Sign in (once per account; the session is saved to ${path.join("~", ".config", "outlook-web-cli")}):
  outlook-web-cli auth chrome  --account work --cdp-port 9230 [--mailbox-url URL]
                           import from a Chrome where Outlook web is open (started with --remote-debugging-port)
  outlook-web-cli auth cookies --account work --file cookies.json [--username me@contoso.com]
                           import login.microsoftonline.com cookies exported from any browser (JSON or cookies.txt; - = stdin)
  outlook-web-cli auth status  [--account work]
  outlook-web-cli auth refresh [--account work | --all] [--rotate]   mint a fresh token; --rotate also renews the sign-in cookies
  outlook-web-cli auth logout  --account work

Mail and calendar (read-only, plus drafts; nothing ever sends mail):
  outlook-web-cli accounts
  outlook-web-cli status      [--account work|all]
  outlook-web-cli messages    [--account all] [--limit 20] [--unread] [--from addr] [--since 2026-11-01] [--folder inbox|sent|drafts|archive]
  outlook-web-cli search      --query "invoice" [--account all] [--limit 20]
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

Any command also takes --json '{"account":"work",...}'. Output is JSON. Draft commands do nothing without --yes.
`;

function parseArgs(argv) {
  const positional = [];
  const flags = {};
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (!a.startsWith("--")) {
      positional.push(a);
      continue;
    }
    const eq = a.indexOf("=");
    const key = (eq > 0 ? a.slice(2, eq) : a.slice(2)).replace(/-/g, "_");
    if (eq > 0) flags[key] = a.slice(eq + 1);
    else if (i + 1 < argv.length && !argv[i + 1].startsWith("--")) flags[key] = argv[++i];
    else flags[key] = true;
  }
  return { positional, flags };
}

function out(result) {
  process.stdout.write(`${JSON.stringify(result, null, 2)}\n`);
  if (result && result.ok === false) process.exitCode = 1;
}

function readInput(file) {
  return file === "-" ? fs.readFileSync(0, "utf8") : fs.readFileSync(file, "utf8");
}

// Remembers the Chrome port so a dead session can be re-imported automatically.
function rememberAccount(name, fields) {
  const file = path.join(configDir(), "config.json");
  let config = { accounts: [] };
  try { config = JSON.parse(fs.readFileSync(file, "utf8")); } catch {}
  if (!Array.isArray(config.accounts)) config.accounts = [];
  const existing = config.accounts.find((a) => a.name === name);
  if (existing) Object.assign(existing, Object.fromEntries(Object.entries(fields).filter(([, v]) => v)));
  else config.accounts.push({ name, ...Object.fromEntries(Object.entries(fields).filter(([, v]) => v)) });
  fs.mkdirSync(configDir(), { recursive: true, mode: 0o700 });
  fs.writeFileSync(file, `${JSON.stringify(config, null, 2)}\n`, { mode: 0o600 });
}

function accountName(flags, { required = true } = {}) {
  const raw = String(flags.account || "").trim();
  const name = resolveAccountAlias(raw) || raw.toLowerCase();
  if (required && !/^[a-z0-9][a-z0-9_-]*$/.test(name)) throw new Error("--account NAME is required (letters, digits, - or _).");
  return name;
}

async function auth(sub, flags) {
  if (sub === "chrome") {
    const name = accountName(flags);
    const known = getAccount(name);
    const cdpPort = Number(flags.cdp_port || known?.cdpPort || 0);
    if (!cdpPort) throw new Error("--cdp-port is required (the Chrome --remote-debugging-port where Outlook web is open).");
    const mainUrl = flags.mailbox_url || known?.mainUrl;
    const imported = await importFromChrome(name, { cdpPort, mainUrl });
    if (!imported.ok) return out(imported);
    rememberAccount(name, { cdp_port: cdpPort, mailbox_url: flags.mailbox_url, username: imported.username });
    const check = await refreshSession(name);
    return out({ ...sessionStatus(name), ok: check.ok, verified: check.ok ? `token minted via ${check.via}` : check.errors });
  }
  if (sub === "cookies") {
    const name = accountName(flags);
    if (!flags.file) throw new Error("--file PATH (or - for stdin) is required.");
    const imported = importCookies(name, readInput(flags.file), { username: flags.username, tenant: flags.tenant });
    if (!imported.ok) return out(imported);
    rememberAccount(name, { username: flags.username });
    const check = await refreshSession(name, { force: "cookies" });
    return out({ ...sessionStatus(name), ok: check.ok, verified: check.ok ? `token minted via ${check.via}` : check.errors });
  }
  if (sub === "status") {
    const names = flags.account ? [accountName(flags)] : listAccounts().map((a) => a.name);
    return out({ ok: true, config_dir: configDir(), accounts: names.map((n) => ({ ...sessionStatus(n), cdp_port: getAccount(n)?.cdpPort || null })) });
  }
  if (sub === "refresh") {
    const accounts = flags.all || !flags.account ? undefined : [accountName(flags)];
    return out(await refreshAccounts({ accounts, rotate: Boolean(flags.rotate) }));
  }
  if (sub === "logout") {
    const name = accountName(flags);
    return out({ ok: true, account: name, removed: deleteSession(name) });
  }
  throw new Error(`Unknown auth command '${sub || ""}'. Use chrome, cookies, status, refresh or logout.`);
}

async function main() {
  const { positional, flags } = parseArgs(process.argv.slice(2));
  const [cmd, sub] = positional;
  if (!cmd || cmd === "help" || flags.help || flags.h) return process.stdout.write(HELP);
  if (cmd === "auth") return auth(sub, flags);
  if (cmd === "accounts") {
    return out({ ok: true, accounts: listAccounts().map((a) => ({ account: a.name, username: a.username, aliases: a.aliases, cdp_port: a.cdpPort, saved_session: sessionStatus(a.name).saved })) });
  }
  const handler = commands[cmd];
  if (!handler) throw new Error(`Unknown command '${cmd}'. Run outlook-web-cli help.`);
  const analyzer = openRouterVisionAnalyzer();
  if (analyzer) setVisionAnalyzer(analyzer);
  const { json, yes, ...rest } = flags;
  const body = { ...rest, ...(json ? JSON.parse(json) : {}) };
  return out(await handler(body, { confirmed: yes === true || yes === "true" }));
}

main().catch((err) => {
  out({ ok: false, status: "error", message: String(err?.message || err) });
});
