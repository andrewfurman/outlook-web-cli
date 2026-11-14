import assert from "node:assert/strict";
import http from "node:http";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
process.env.OUTLOOK_WEB_CONFIG_DIR = fs.mkdtempSync(path.join(os.tmpdir(), "outlook-web-test-"));
const { configureAccounts } = await import("../src/accounts.mjs");
configureAccounts([
  { name: "contoso", aliases: ["contoso.com", "contoso work"], username: "me@contoso.com" },
  { name: "fabrikam", aliases: ["fabrikam.com"], username: "me@fabrikam.com" },
]);

// Spin up a tiny fake upstream server simulating Graph and Outlook REST
const requests = [];
const server = http.createServer((req, res) => {
  const url = new URL(req.url || "", `http://${req.headers.host}`);
  requests.push(url);
  res.setHeader("content-type", "application/json; charset=utf-8");
  // Simple routing for Graph defaults (used by initial tests)
  if (url.pathname === "/v1.0/me/mailFolders/Inbox/messages") {
    res.writeHead(200);
    res.end(JSON.stringify({
      value: [
        {
          id: "MSG_A",
          subject: "Hello A",
          from: { emailAddress: { name: "Alice", address: "alice@example.com" } },
          receivedDateTime: "2026-10-09T16:00:00Z",
          isRead: false,
          conversationId: "C1",
          webLink: "https://outlook.office.com/mail/deeplink/read/MSG_A",
        },
        {
          id: "MSG_B",
          subject: "Hello B",
          from: { emailAddress: { name: "Bob", address: "bob@example.com" } },
          receivedDateTime: "2026-10-09T15:00:00Z",
          isRead: true,
          conversationId: "C2",
          webLink: "https://outlook.office.com/mail/deeplink/read/MSG_B",
        },
      ],
    }));
    return;
  }
  if (url.pathname.startsWith("/v1.0/me/messages/")) {
    res.writeHead(200);
    res.end(JSON.stringify({
      id: url.pathname.split("/").pop(),
      subject: "Read Test",
      from: { emailAddress: { name: "Carol", address: "carol@example.com" } },
      receivedDateTime: "2026-10-09T17:00:00Z",
      body: { contentType: "html", content: "<p>Line 1<br>Line 2</p>" },
      webLink: "https://outlook.office.com/mail/deeplink/read/READ_ID",
    }));
    return;
  }
  if (url.pathname === "/v1.0/me/messages" && url.searchParams.get("$search")) {
    res.writeHead(200);
    res.end(JSON.stringify({
      value: [
        {
          id: "MSG_S",
          subject: "Search Hit",
          from: { emailAddress: { name: "Dana", address: "dana@example.com" } },
          receivedDateTime: "2026-10-09T18:00:00Z",
          isRead: false,
          conversationId: "S1",
          webLink: "https://outlook.office.com/mail/deeplink/read/MSG_S",
        },
      ],
    }));
    return;
  }
  if (url.pathname === "/v1.0/me/calendarView") {
    res.writeHead(200);
    res.end(JSON.stringify({
      value: [
        {
          id: "EV1",
          subject: "Standup",
          isAllDay: false,
          start: { dateTime: "2026-10-09T09:00:00", timeZone: "America/New_York" },
          end: { dateTime: "2026-10-09T09:30:00", timeZone: "America/New_York" },
          organizer: { emailAddress: { name: "Andrew", address: "andrew@example.com" } },
          location: { displayName: "Zoom" },
          webLink: "https://outlook.office.com/calendar/item/EV1",
        },
      ],
    }));
    return;
  }
  // Simulate Graph base override returning 401 to trigger fallback
  if (url.pathname === "/graph/v1.0/me/mailFolders/Inbox/messages") {
    res.writeHead(401);
    res.end(JSON.stringify({ error: { code: "InvalidAuthenticationToken", message: "Bad token" } }));
    return;
  }
  // Outlook REST v2.0 folder listing: Sent Items holds a sent reply, Drafts an unsent one
  const folderList = url.pathname.match(/^\/api\/v2\.0\/me\/MailFolders\/(\w+)\/messages$/);
  if (folderList) {
    const folder = folderList[1];
    res.writeHead(200);
    res.end(JSON.stringify({
      value: [
        {
          Id: `MSG_${folder}`,
          Subject: "RE: Duplicate checks",
          From: { EmailAddress: { Name: "AF (Shared)", Address: "me@fabrikam.com" } },
          ReceivedDateTime: "2026-11-11T19:06:05Z",
          IsRead: true,
          IsDraft: folder === "Drafts",
          ConversationId: "C9",
          WebLink: `https://outlook.office.com/mail/deeplink/read/MSG_${folder}`,
        },
      ],
    }));
    return;
  }
  // Outlook REST v2.0 simulation
  if (url.pathname === "/outlook/api/v2.0/me/messages" || url.pathname === "/api/v2.0/me/messages") {
    res.writeHead(200);
    res.end(JSON.stringify({
      value: [
        {
          Id: "MSG_A",
          Subject: "Hello A",
          From: { EmailAddress: { Name: "Alice", Address: "alice@example.com" } },
          ReceivedDateTime: "2026-10-09T16:00:00Z",
          IsRead: false,
          ConversationId: "C1",
          WebLink: "https://outlook.office.com/mail/deeplink/read/MSG_A",
        },
        {
          Id: "MSG_B",
          Subject: "Hello B",
          From: { EmailAddress: { Name: "Bob", Address: "bob@example.com" } },
          ReceivedDateTime: "2026-10-09T15:00:00Z",
          IsRead: true,
          ConversationId: "C2",
          WebLink: "https://outlook.office.com/mail/deeplink/read/MSG_B",
        },
      ],
    }));
    return;
  }
  if (url.pathname.startsWith("/outlook/api/v2.0/me/messages/") || url.pathname.startsWith("/api/v2.0/me/messages/")) {
    res.writeHead(200);
    res.end(JSON.stringify({
      Id: url.pathname.split("/").pop(),
      Subject: "Read Test",
      From: { EmailAddress: { Name: "Carol", Address: "carol@example.com" } },
      ReceivedDateTime: "2026-10-09T17:00:00Z",
      Body: { ContentType: "HTML", Content: "<p>Line 1<br>Line 2</p>" },
      WebLink: "https://outlook.office.com/mail/deeplink/read/READ_ID",
    }));
    return;
  }
  if ((url.pathname === "/outlook/api/v2.0/me/messages" || url.pathname === "/api/v2.0/me/messages") && url.searchParams.get("$search")) {
    res.writeHead(200);
    res.end(JSON.stringify({
      value: [
        {
          Id: "MSG_S",
          Subject: "Search Hit",
          From: { EmailAddress: { Name: "Dana", Address: "dana@example.com" } },
          ReceivedDateTime: "2026-10-09T18:00:00Z",
          IsRead: false,
          ConversationId: "S1",
          WebLink: "https://outlook.office.com/mail/deeplink/read/MSG_S",
        },
      ],
    }));
    return;
  }
  if (url.pathname === "/outlook/api/v2.0/me/calendarview" || url.pathname === "/api/v2.0/me/calendarview") {
    res.writeHead(200);
    res.end(JSON.stringify({
      value: [
        {
          Id: "EV1",
          Subject: "Standup",
          IsAllDay: false,
          Start: { DateTime: "2026-10-09T09:00:00", TimeZone: "America/New_York" },
          End: { DateTime: "2026-10-09T09:30:00", TimeZone: "America/New_York" },
          Organizer: { EmailAddress: { Name: "Andrew", Address: "andrew@example.com" } },
          Location: { DisplayName: "Zoom" },
          WebLink: "https://outlook.office.com/calendar/item/EV1",
        },
      ],
    }));
    return;
  }
  res.writeHead(404);
  res.end(JSON.stringify({ error: { code: "not_found", message: "not found" } }));
});

await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
const { port } = server.address();
const base = `http://127.0.0.1:${port}`;

// Inject test API bases and token so CDP is not needed.
process.env.OUTLOOK_GRAPH_BASE_URL = base;
process.env.OUTLOOK_REST_BASE_URL = base;
// token with a benign payload by default; specific tests will override
process.env.OUTLOOK_TEST_TOKEN = `x.${Buffer.from(JSON.stringify({ aud: "https://graph.microsoft.com", scp: "Mail.Read" })).toString("base64url")}.y`;

// Import after env is set so the module reads the overridden base.
const { outlookAgenda, outlookMessages, outlookRead, outlookSearch, outlookStatus } = await import("../src/outlook.mjs");

// status
{
  const s = await outlookStatus({ account: "contoso" });
  assert.equal(s.ok, true);
  assert.equal(s.accounts?.[0]?.status, "signed_in");
}

// list messages: simulate Graph 401 then Outlook REST success via base overrides and token aud
{
  const r = await outlookMessages({ account: "contoso", limit: 2, unread: false });
  assert.equal(r.ok, true);
  assert.equal(r.returned_count, 2);
  assert.equal(r.items[0].id, "MSG_A");
  assert.equal(typeof r.items[0].received_at, "string");
}

// Outlook REST lists one folder, Inbox by default, and flags drafts (an unsent draft is not a sent reply)
{
  process.env.OUTLOOK_TEST_TOKEN_FABRIKAM = `x.${Buffer.from(JSON.stringify({ aud: "https://outlook.office.com", scp: "Mail.Read" })).toString("base64url")}.y`;
  const last = () => requests[requests.length - 1];

  const inbox = await outlookMessages({ account: "fabrikam", limit: 5 });
  assert.equal(inbox.ok, true);
  assert.equal(inbox.folder, "Inbox");
  assert.equal(last().pathname, "/api/v2.0/me/MailFolders/Inbox/messages");
  assert.equal(inbox.items[0].folder, "Inbox");
  assert.equal(inbox.items[0].is_draft, false);

  const sent = await outlookMessages({ account: "fabrikam", folder: "sent" });
  assert.equal(last().pathname, "/api/v2.0/me/MailFolders/SentItems/messages");
  assert.equal(sent.items[0].folder, "SentItems");
  assert.equal(sent.items[0].is_draft, false);

  const drafts = await outlookMessages({ account: "fabrikam", folder: "Drafts" });
  assert.equal(drafts.items[0].is_draft, true);

  const bad = await outlookMessages({ account: "fabrikam", folder: "everything" });
  assert.equal(bad.ok, false);
  assert.equal(bad.status, "invalid_folder");

  // A sender filter alone still sorts by date: the date clause leads the filter.
  await outlookMessages({ account: "fabrikam", from: "dia@fabrikam.com" });
  assert.equal(last().searchParams.get("$orderby"), "ReceivedDateTime desc");
  assert.match(last().searchParams.get("$filter"), /^ReceivedDateTime ge 1900-01-01T00:00:00Z and From\/EmailAddress\/Address eq 'dia@fabrikam.com'$/);

  // $search can't be combined with $orderby.
  await outlookMessages({ account: "fabrikam", search: "duplicate" });
  assert.equal(last().searchParams.get("$search"), '"duplicate"');
  assert.equal(last().searchParams.get("$orderby"), null);

  delete process.env.OUTLOOK_TEST_TOKEN_FABRIKAM;
}

// Graph lists the requested folder too, and searches without $orderby
{
  await outlookMessages({ account: "contoso", search: "invoice" });
  const url = requests[requests.length - 1];
  assert.equal(url.pathname, "/v1.0/me/mailFolders/Inbox/messages");
  assert.equal(url.searchParams.get("$orderby"), null);
  await outlookSearch({ account: "contoso", query: "invoice" });
  assert.equal(requests[requests.length - 1].searchParams.get("$orderby"), null);
}

// read message
{
  const r = await outlookRead({ account: "contoso", id: "READ_ID", maxBodyChars: 2000 });
  assert.equal(r.ok, true);
  assert.match(r.body_text || "", /Line 1/);
}

// search
{
  const r = await outlookSearch({ account: "contoso", query: "search words", limit: 1 });
  assert.equal(r.ok, true);
  assert.equal(r.returned_count, 1);
  assert.equal(r.items[0].id, "MSG_S");
}

// agenda today should use ET-local day bounds (00:00..23:59 ET) via Prefer header
{
  const r = await outlookAgenda({ account: "contoso", today: true });
  assert.equal(r.ok, true);
  const acc = r.accounts?.[0];
  assert.equal(acc?.status, "ok");
  assert.ok((acc?.returned_count || 0) >= 1);
  assert.equal(acc?.items?.[0]?.subject, "Standup");
  // Start/end should be local ET times without Z
  assert.match(String(r.start_date), /T00:00:00$/);
  assert.match(String(r.end_date), /T23:59:59$/);
}

// account aliases: names, aliases, usernames and email domains
{
  const { resolveWorkAccountAlias } = await import("../src/outlook.mjs");
  for (const a of ["contoso", "Contoso", "contoso work", "me@contoso.com", "contoso.com", "anyone@contoso.com"]) assert.equal(resolveWorkAccountAlias(a), "contoso", a);
  for (const a of ["fabrikam", "FABRIKAM", "fabrikam.com"]) assert.equal(resolveWorkAccountAlias(a), "fabrikam", a);
  for (const a of ["", undefined, "all", "gmail", "andrew@gmail.com"]) assert.equal(resolveWorkAccountAlias(a), null, String(a));
  const viaAlias = await outlookMessages({ account: "fabrikam.com", limit: 1 });
  assert.equal(viaAlias.ok, true);
  const { commands } = await import("../src/commands.mjs");
  const listed = await commands.messages({ account: "contoso", limit: 1 });
  assert.equal(listed.ok, true);
  const pending = await commands["draft-new"]({ account: "contoso", to: "a@b.c", subject: "s", body: "b" }, {});
  assert.equal(pending.status, "confirmation_required");
}

server.close();

console.log(JSON.stringify({ ok: true, status: "ok", tests: "outlook-tools", http_port: port }, null, 2));

