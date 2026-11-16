import assert from "node:assert/strict";
import fs from "node:fs";
import http from "node:http";
import os from "node:os";
import path from "node:path";

// A fake Outlook REST getSchedule endpoint.
const b64 = (o) => Buffer.from(JSON.stringify(o)).toString("base64url");
const posts = [];
const server = http.createServer(async (req, res) => {
  let body = "";
  for await (const chunk of req) body += chunk;
  res.setHeader("content-type", "application/json");
  if (req.method === "POST" && req.url === "/api/v2.0/me/calendar/getschedule") {
    const payload = JSON.parse(body);
    posts.push({ payload, headers: req.headers });
    const item = (s, e, status, extra = {}) => ({ Status: status, Start: { DateTime: `${s}:00.0000000`, TimeZone: "Eastern Standard Time" }, End: { DateTime: `${e}:00.0000000`, TimeZone: "Eastern Standard Time" }, ...extra });
    const data = {
      "alice@contoso.com": { ScheduleItems: [
        item("2026-11-17T10:00", "2026-11-17T10:30", "Busy"),
        item("2026-11-17T10:15", "2026-11-17T11:00", "Tentative"),
        item("2026-11-17T13:00", "2026-11-17T13:30", "Free"),
        item("2026-11-17T19:00", "2026-11-17T20:00", "Busy"),
      ] },
      "bob@contoso.com": { ScheduleItems: [
        item("2026-11-17T11:30", "2026-11-17T12:00", "Busy", { Subject: "Shared calendar standup", Location: "Room 2" }),
        item("2026-11-17T14:00", "2026-11-17T16:45", "Oof"),
      ] },
      "nobody@elsewhere.com": { Error: { Message: "The user wasn't found." } },
    };
    res.writeHead(200);
    return res.end(JSON.stringify({ value: payload.Schedules.map((id) => ({ ScheduleId: id, ...(data[id] || { ScheduleItems: [] }) })) }));
  }
  res.writeHead(404);
  res.end("{}");
});
await new Promise((r) => server.listen(0, "127.0.0.1", r));
process.env.OUTLOOK_WEB_CONFIG_DIR = fs.mkdtempSync(path.join(os.tmpdir(), "outlook-web-cli-freebusy-"));
process.env.OUTLOOK_REST_BASE_URL = `http://127.0.0.1:${server.address().port}`;
process.env.OUTLOOK_TEST_TOKEN_CONTOSO = `${b64({ alg: "none" })}.${b64({ aud: "https://outlook.office.com", scp: "Calendars.ReadWrite" })}.sig`;

const { configureAccounts } = await import("../src/accounts.mjs");
const { commands } = await import("../src/commands.mjs");
configureAccounts([
  { name: "contoso", aliases: ["contoso.com"], username: "me@contoso.com" },
  { name: "fabrikam", aliases: ["fabrikam.com"], username: "me@fabrikam.com" },
]);

try {
  // 1. Two people for one day: the account is picked from their email domain.
  const r = await commands.freebusy({ emails: "Alice <alice@contoso.com>, bob@contoso.com; ALICE@contoso.com", start_date: "2026-11-17", end_date: "2026-11-17" });
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.account, "contoso");
  assert.deepEqual(posts[0].payload.Schedules, ["alice@contoso.com", "bob@contoso.com"], "deduplicated, lower-cased");
  assert.equal(posts[0].payload.StartTime.DateTime, "2026-11-17T00:00:00");
  assert.equal(posts[0].payload.StartTime.TimeZone, "Eastern Standard Time");
  assert.equal(posts[0].headers.prefer, 'outlook.timezone="America/New_York"', "answers in Eastern, not UTC");
  const alice = r.people.find((p) => p.email === "alice@contoso.com");
  assert.deepEqual(alice.busy.map((b) => [b.start.slice(11), b.end.slice(11), b.status]), [["10:00", "11:00", "busy"], ["19:00", "20:00", "busy"]], "overlaps merged, Free items dropped");
  const bob = r.people.find((p) => p.email === "bob@contoso.com");
  assert.equal(bob.busy[0].subject, "Shared calendar standup");
  assert.equal(bob.busy[1].status, "oof");
  assert.deepEqual(r.free_for_all.map((s) => [s.start.slice(11), s.end.slice(11)]), [["09:00", "10:00"], ["11:00", "11:30"], ["12:00", "14:00"]], "16:45-17:00 is shorter than 30 minutes");
  assert.match(r.answer_text, /3 slots when everyone is free/);

  // 2. A date-time window, a shorter minimum and a custom working day.
  const w = await commands.freebusy({ emails: ["alice@contoso.com", "bob@contoso.com"], start: "2026-11-17T11:00", end: "2026-11-17T18:00", min_minutes: 15, day_end: "18:00" });
  assert.equal(posts.at(-1).payload.StartTime.DateTime, "2026-11-17T11:00:00", "seconds added");
  assert.deepEqual(w.free_for_all.map((s) => [s.start.slice(11), s.end.slice(11), s.minutes]), [["11:00", "11:30", 30], ["12:00", "14:00", 120], ["16:45", "18:00", 75]]);

  // 3. Someone without free/busy is reported, the rest still answer.
  const x = await commands.freebusy({ account: "contoso", emails: "alice@contoso.com,nobody@elsewhere.com", tomorrow: true });
  assert.equal(x.ok, true);
  assert.equal(x.people.find((p) => p.email === "nobody@elsewhere.com").ok, false);
  assert.match(x.answer_text, /No free\/busy for nobody@elsewhere.com/);

  // 4. Weekends are skipped unless asked for.
  const sat = await commands.freebusy({ emails: "alice@contoso.com", start_date: "2026-11-21", end_date: "2026-11-21" });
  assert.equal(sat.free_for_all.length, 0);
  const satOn = await commands.freebusy({ emails: "alice@contoso.com", start_date: "2026-11-21", end_date: "2026-11-21", weekends: true });
  assert.equal(satOn.free_for_all.length, 1);

  // 5. Input errors.
  assert.equal((await commands.freebusy({ emails: "" })).status, "missing_emails");
  assert.equal((await commands.freebusy({ emails: "someone@gmail.com" })).status, "account_required", "two accounts and no domain match");
  assert.equal((await commands.freebusy({ emails: Array.from({ length: 21 }, (_, i) => `p${i}@contoso.com`) })).status, "too_many_emails");

  console.log("freebusy tests: ok");
} finally {
  server.close();
}
