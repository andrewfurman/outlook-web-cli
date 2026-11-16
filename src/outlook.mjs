import { htmlToText } from "html-to-text";
import { getAccount, listAccounts, resolveAccountAlias } from "./accounts.mjs";
import { connectCdp, findOutlookTab } from "./cdp.mjs";
import { getSessionToken } from "./session.mjs";

// Outlook mail and calendar through the Outlook web session (see session.mjs): no app
// registration and no admin consent, the same access the signed-in web app has.
// No tokens are ever logged. Commands are read-only, except the draft commands
// (outlookDraftReply / outlookDraftNew / outlookDeleteDraft), which save or discard drafts
// like Outlook web's Reply / Reply all / New message do. Nothing here ever sends mail.
//
// Tests inject a token (OUTLOOK_TEST_TOKEN[_<ACCOUNT>]) and HTTP base URLs.

const DEFAULT_LIST_LIMIT = 20;
const MAX_LIST_LIMIT = 50;
const DEFAULT_READ_MAX_BODY_CHARS = 20000;
const MAX_READ_MAX_BODY_CHARS = 60000;
const DEFAULT_TIMEOUT_MS = 15000;
const ET_TIMEZONE = "America/New_York";
// `outlook attachments` returns base64 in the response only (nothing is written to disk):
// whole files up to max_data_bytes, or a byte range (offset/length) of one attachment so a
// caller with a small output limit, such as SSM, can pull a big image in pieces.
const DEFAULT_ATTACHMENT_DATA_BYTES = 250_000;
const MAX_ATTACHMENT_DATA_BYTES = 1_000_000;
const MAX_INSPECT_IMAGE_BYTES = 5_000_000;
const ATTACHMENT_TIMEOUT_MS = 30000;
// The process keeps a message's attachments in memory for a few minutes (never on disk), so
// pulling one image in many byte ranges downloads it from Outlook once, not once per piece.
const ATTACHMENT_CACHE_TTL_MS = 5 * 60_000;
const ATTACHMENT_CACHE_MAX_ENTRIES = 4;
const attachmentCache = new Map();
let defaultVisionAnalyze = null;

// Programs that embed this package can plug in their own image model for
// outlookAttachmentInspect; the CLI uses OpenRouter when OPENROUTER_API_KEY is set.
export function setVisionAnalyzer(fn) {
  defaultVisionAnalyze = typeof fn === "function" ? fn : null;
}
const DEFAULT_ATTACHMENT_INSPECT_PROMPT =
  "Describe this screenshot or image in detail, then transcribe ALL visible text verbatim in reading order " +
  "(menus, labels, table cells, numbers, dates). Put the full transcription in ocr_text.";

// `outlook messages` lists one well-known folder, Inbox by default. Listing all of
// /me/messages mixes in Drafts and Sent Items, so an unsent draft looks like a sent reply.
const MAIL_FOLDERS = {
  inbox: "Inbox",
  sent: "SentItems",
  sentitems: "SentItems",
  drafts: "Drafts",
  deleted: "DeletedItems",
  deleteditems: "DeletedItems",
  archive: "Archive",
  junk: "JunkEmail",
  junkemail: "JunkEmail",
};
// Exchange rejects $orderby on a property that is not also the first $filter clause.
const ALL_TIME_UTC = "1900-01-01T00:00:00Z";

const GRAPH_BASE =
  normalizeUrl(process.env.OUTLOOK_GRAPH_BASE_URL) || "https://graph.microsoft.com";
const OUTLOOK_REST_BASE =
  normalizeUrl(process.env.OUTLOOK_REST_BASE_URL) || "https://outlook.office.com";

export async function outlookStatus({ account } = {}) {
  const targets = await resolveAccounts(account);
  const summaries = [];
  for (const target of targets) {
    const auth = await safeGetBestToken(target, "mail").catch(() => null);
    summaries.push({
      account: target.label,
      ok: Boolean(auth?.token),
      status: auth?.token ? "signed_in" : "needs_sign_in",
      token_source: auth?.token ? auth.source || "chrome_tab" : "",
    });
  }
  const signedInCount = summaries.filter((s) => s.ok).length;
  return {
    ok: true,
    status: "ok",
    accounts_checked: summaries.length,
    accounts: summaries,
    answer_text:
      summaries.length === 1
        ? (summaries[0].ok ? "Signed in." : "Needs sign-in.")
        : `${signedInCount}/${summaries.length} Outlook sessions are signed in.`,
  };
}

export async function outlookMessages({
  account,
  limit = DEFAULT_LIST_LIMIT,
  unread = false,
  from,
  since,
  search,
  folder = "inbox",
  maxRawBytes, // accepted but unused to preserve common signature style
} = {}) {
  const folderName = resolveMailFolder(folder);
  if (!folderName) {
    return {
      ok: false,
      status: "invalid_folder",
      folder: String(folder),
      answer_text: `Unknown Outlook folder "${folder}". Use inbox, sent, drafts, deleted, archive or junk.`,
    };
  }
  const targets = await resolveAccounts(account);
  const boundedLimit = clampInteger(limit, 1, MAX_LIST_LIMIT, DEFAULT_LIST_LIMIT);
  const queries = [];
  for (const target of targets) {
    const auth = await safeGetBestToken(target, "mail");
    if (!auth?.token) {
      queries.push({
        ok: false,
        account: target.label,
        status: "needs_sign_in",
        returned_count: 0,
        items: [],
      });
      continue;
    }
    const token = auth.token;
    const first = auth.api || preferredApiForToken(decodeJwtClaims(token));
    const attempt = async (api) => {
      if (api === "graph") {
        const url = new URL(`/v1.0/me/mailFolders/${folderName}/messages`, ensureTrailingSlash(GRAPH_BASE));
        const params = new URLSearchParams();
        params.set("$select", "id,subject,from,receivedDateTime,isRead,isDraft,conversationId,webLink");
        params.set("$top", String(boundedLimit));
        if (normalizeString(search)) {
          // Graph returns search hits by relevance and rejects $orderby with $search.
          params.set("$search", `"${normalizeString(search)}"`);
        } else {
          params.set("$orderby", "receivedDateTime desc");
          const filter = listFilter({ unread, from, since }, { received: "receivedDateTime", isRead: "isRead", fromAddress: "from/emailAddress/address" });
          if (filter) params.set("$filter", filter);
        }
        url.search = params.toString();
        const headers = graphHeaders(token, { needsConsistencyLevel: Boolean(params.get("$search")) });
        return await fetchJsonDetailed(url, { headers, timeoutMs: DEFAULT_TIMEOUT_MS, api: "graph" });
      } else {
        const url = new URL(`/api/v2.0/me/MailFolders/${folderName}/messages`, ensureTrailingSlash(OUTLOOK_REST_BASE));
        const params = new URLSearchParams();
        params.set("$select", "Id,Subject,From,ReceivedDateTime,IsRead,IsDraft,ConversationId,WebLink");
        params.set("$top", String(boundedLimit));
        // Outlook REST v2.0 supports $filter; $search support varies, prefer $filter when possible.
        const filter = listFilter({ unread, from, since }, { received: "ReceivedDateTime", isRead: "IsRead", fromAddress: "From/EmailAddress/Address" });
        if (filter) params.set("$filter", filter);
        if (!filter && normalizeString(search)) {
          params.set("$search", `"${normalizeString(search)}"`);
        } else {
          params.set("$orderby", "ReceivedDateTime desc");
        }
        url.search = params.toString();
        const headers = { authorization: `Bearer ${token}`, accept: "application/json" };
        return await fetchJsonDetailed(url, { headers, timeoutMs: DEFAULT_TIMEOUT_MS, api: "outlook_rest" });
      }
    };
    let res = await attempt(first);
    if (!res.ok && shouldFallback(res)) {
      const second = first === "graph" ? "outlook_rest" : "graph";
      const res2 = await attempt(second);
      res = res2;
    }
    if (!res.ok) {
      queries.push({
        ok: false,
        account: target.label,
        status: "upstream_error",
        upstream_http_status: res.status,
        upstream_error: res.error || "",
        returned_count: 0,
        items: [],
      });
      continue;
    }
    const isGraph = res.api === "graph";
    const rows = Array.isArray(res.body?.value) ? res.body.value : [];
    const items = rows.map((m) => ({
      ...(isGraph ? compactGraphMessage(m, target.label) : compactOutlookRestMessage(m, target.label)),
      folder: folderName,
    }));
    queries.push({ ok: true, account: target.label, status: "ok", folder: folderName, returned_count: items.length, items });
  }
  // Merge results for "all"
  const merged = mergeAccountLists(queries);
  const anyError = queries.some((q) => !q.ok);
  const allFailed = queries.every((q) => !q.ok);
  return {
    ok: !allFailed,
    status: allFailed ? "upstream_error" : "ok",
    scope: accountScopeLabel(account),
    folder: folderName,
    returned_count: merged.length,
    items: merged,
    answer_text:
      merged.length === 0 && anyError
        ? "The Outlook upstream call failed."
        : merged.length === 0
        ? `No messages matched in ${folderName}.`
        : `Found ${merged.length} message${merged.length === 1 ? "" : "s"} in ${folderName}.`,
    accounts: queries,
  };
}

export async function outlookRead({
  account,
  id,
  maxBodyChars = DEFAULT_READ_MAX_BODY_CHARS,
} = {}) {
  const messageId = normalizeString(id);
  if (!messageId) return missingField("id", "An Outlook message id is required.");
  const targets = await resolveAccounts(account);
  // If "all", prefer the first account that can read the id; attempt sequentially.
  for (const target of targets) {
    const auth = await safeGetBestToken(target, "mail");
    if (!auth?.token) continue;
    const token = auth.token;
    const claims = decodeJwtClaims(token);
    const first = auth.api || preferredApiForToken(claims);
    const attempt = async (api) => {
      if (api === "graph") {
        const url = new URL(`/v1.0/me/messages/${encodeURIComponent(messageId)}`, ensureTrailingSlash(GRAPH_BASE));
        const params = new URLSearchParams();
        params.set("$select", "id,subject,from,receivedDateTime,isDraft,body,webLink");
        url.search = params.toString();
        const headers = graphHeaders(token);
        return await fetchJsonDetailed(url, { headers, timeoutMs: DEFAULT_TIMEOUT_MS, api: "graph" });
      } else {
        const url = new URL(`/api/v2.0/me/messages/${encodeURIComponent(messageId)}`, ensureTrailingSlash(OUTLOOK_REST_BASE));
        const params = new URLSearchParams();
        params.set("$select", "Id,Subject,From,ReceivedDateTime,IsDraft,Body,WebLink");
        url.search = params.toString();
        const headers = { authorization: `Bearer ${token}`, accept: "application/json" };
        return await fetchJsonDetailed(url, { headers, timeoutMs: DEFAULT_TIMEOUT_MS, api: "outlook_rest" });
      }
    };
    let res = await attempt(first);
    if (!res.ok && shouldFallback(res)) res = await attempt(first === "graph" ? "outlook_rest" : "graph");
    if (!res.ok) return { ok: false, status: "message_read_failed", account: target.label, id: messageId, upstream_http_status: res.status, answer_text: "The Outlook upstream call failed." };
    const isGraph = res.api === "graph";
    const raw = res.body || {};
    const compact = (isGraph ? compactGraphMessage(raw, target.label, {
      includeBody: true,
      maxBodyChars: clampInteger(maxBodyChars, 2000, MAX_READ_MAX_BODY_CHARS, DEFAULT_READ_MAX_BODY_CHARS),
    }) : compactOutlookRestMessage(raw, target.label, {
      includeBody: true,
      maxBodyChars: clampInteger(maxBodyChars, 2000, MAX_READ_MAX_BODY_CHARS, DEFAULT_READ_MAX_BODY_CHARS),
    }));
    return {
      ok: true,
      status: "ok",
      account: target.label,
      id: compact.id,
      subject: compact.subject,
      from: compact.from,
      received_at: compact.received_at,
      received_datetime: compact.received_datetime,
      is_draft: compact.is_draft,
      web_link: compact.web_link || "",
      body_text: compact.body_text || "",
      answer_text: compact.is_draft
        ? `Unsent draft: ${compact.subject || "(no subject)"}.`
        : compact.body_text
        ? `${compact.subject || "Email"} from ${compact.from?.address || "unknown sender"}.`
        : "Message has no readable body.",
    };
  }
  // Tried all accounts
  return {
    ok: false,
    status: "message_not_found",
    scope: accountScopeLabel(account),
    id: messageId,
    answer_text: "Message not found.",
  };
}

// Lists a message's attachments, including inline images pasted into the body (the
// cid: images Outlook calls image001.png). Pick one with index, content_id or
// attachment_id; offset/length return that byte range of it. Read-only: GET requests only.
export async function outlookAttachments({
  account,
  id,
  includeData,
  maxDataBytes = DEFAULT_ATTACHMENT_DATA_BYTES,
  imagesOnly = false,
  index,
  contentId,
  attachmentId,
  offset,
  length,
} = {}) {
  const messageId = normalizeString(id);
  if (!messageId) return missingField("id", "An Outlook message id is required.");
  const listed = await fetchAttachments(account, messageId, normalizeString(attachmentId));
  if (!listed.ok) return listed;
  const dataCap = clampInteger(maxDataBytes, 1_000, MAX_ATTACHMENT_DATA_BYTES, DEFAULT_ATTACHMENT_DATA_BYTES);
  const wantsRange = offset !== undefined && offset !== null && offset !== "";
  const pickOne = wantsRange || normalizeString(attachmentId) || normalizeString(contentId) || (index !== undefined && index !== null && index !== "");
  if (pickOne) {
    const att = selectAttachment(listed.attachments, { index, contentId, attachmentId });
    if (!att) return { ok: false, status: "attachment_not_found", account: listed.account, id: messageId, answer_text: "No attachment matched that index, content id or attachment id." };
    const { data_base64, ...item } = att;
    const bytes = Buffer.from(data_base64 || "", "base64");
    if (wantsRange) {
      const start = clampInteger(offset, 0, bytes.length, 0);
      const end = Math.min(bytes.length, start + clampInteger(length, 1_000, MAX_ATTACHMENT_DATA_BYTES, DEFAULT_ATTACHMENT_DATA_BYTES));
      return {
        ok: true, status: "ok", account: listed.account, id: messageId, attachment: item,
        total_bytes: bytes.length, offset: start, returned_bytes: end - start,
        next_offset: end < bytes.length ? end : null,
        data_base64: bytes.subarray(start, end).toString("base64"),
        answer_text: `Bytes ${start}-${end} of ${bytes.length} of ${item.name}.`,
      };
    }
    const fits = bytes.length <= dataCap;
    return {
      ok: true, status: "ok", account: listed.account, id: messageId, attachment: item, total_bytes: bytes.length,
      ...(toBoolean(includeData, true) && fits ? { data_base64 } : {}),
      answer_text: fits ? `${item.name}, ${bytes.length} bytes.` : `${item.name} is ${bytes.length} bytes; fetch it in pieces with offset and length.`,
    };
  }
  const withData = toBoolean(includeData);
  const items = [];
  for (const att of listed.attachments) {
    if (toBoolean(imagesOnly) && !att.is_image) continue;
    const { data_base64, ...item } = att;
    if (withData && data_base64 && Buffer.byteLength(data_base64, "base64") <= dataCap) item.data_base64 = data_base64;
    items.push(item);
  }
  const images = items.filter((x) => x.is_image);
  return {
    ok: true,
    status: "ok",
    account: listed.account,
    id: messageId,
    returned_count: items.length,
    image_count: images.length,
    attachments: items,
    answer_text: items.length
      ? `${items.length} attachment${items.length === 1 ? "" : "s"}${images.length ? `, ${images.length} image${images.length === 1 ? "" : "s"}` : ""}: ${items.slice(0, 5).map((x) => x.name).join("; ")}.`
      : "That email has no attachments.",
  };
}

function selectAttachment(attachments, { index, contentId, attachmentId }) {
  const aid = normalizeString(attachmentId);
  if (aid) return attachments.find((x) => x.attachment_id === aid) || (attachments.length === 1 ? attachments[0] : null);
  const cid = normalizeString(contentId).replace(/^<|>$/g, "").toLowerCase();
  if (cid) return attachments.find((x) => x.content_id.toLowerCase() === cid) || null;
  if (index !== undefined && index !== null && index !== "") return attachments.find((x) => x.index === Number(index)) || null;
  return attachments.find((x) => x.is_image) || null;
}

// Sends one image attachment (default: the first image) to the vision model and returns a
// description plus the full transcribed text. OpenRouter is used when its key is set.
export async function outlookAttachmentInspect({
  account,
  id,
  index,
  contentId,
  prompt,
  visionAnalyze = defaultVisionAnalyze,
} = {}) {
  const messageId = normalizeString(id);
  if (!messageId) return missingField("id", "An Outlook message id is required.");
  const listed = await fetchAttachments(account, messageId);
  if (!listed.ok) return listed;
  const images = listed.attachments.filter((x) => x.is_image && x.data_base64);
  const image = selectAttachment(listed.attachments, { index, contentId });
  if (!image || !image.is_image || !image.data_base64) {
    return {
      ok: false,
      status: images.length ? "image_not_found" : "no_images",
      account: listed.account,
      id: messageId,
      images: images.map(({ data_base64, ...x }) => x),
      answer_text: images.length ? "No image matched that index or content id." : "That email has no images.",
    };
  }
  if (image.size > MAX_INSPECT_IMAGE_BYTES) {
    return { ok: false, status: "image_too_large", account: listed.account, id: messageId, answer_text: "That image is too large to inspect." };
  }
  if (typeof visionAnalyze !== "function") {
    const message = "No image model is set up. Set OPENROUTER_API_KEY, or call setVisionAnalyzer().";
    return { ok: false, status: "vision_not_configured", account: listed.account, id: messageId, message, answer_text: message };
  }
  const vision = await visionAnalyze({
    provider: "openrouter",
    prompt: normalizeString(prompt) || DEFAULT_ATTACHMENT_INSPECT_PROMPT,
    images: [{ media_type: image.content_type, data_base64: image.data_base64, filename: image.name }],
  });
  const { data_base64, ...meta } = image;
  if (!vision.ok) {
    return { ok: false, status: vision.status || "vision_error", account: listed.account, id: messageId, image: meta, message: vision.message, answer_text: vision.message || "Vision analysis failed." };
  }
  return {
    ok: true,
    status: "ok",
    account: listed.account,
    id: messageId,
    image: meta,
    provider: vision.provider || "",
    upstream_provider: vision.upstream_provider || "",
    model: vision.model || "",
    usage: vision.usage || null,
    description: vision.description || "",
    ocr_text: vision.ocr_text || "",
    answer_text: vision.answer_text || vision.description || "I inspected the image.",
  };
}

async function fetchAttachments(account, messageId, attachmentId = "") {
  const key = JSON.stringify([normalizeString(account).toLowerCase(), messageId, attachmentId]);
  const cached = attachmentCache.get(key);
  if (cached && Date.now() - cached.at < ATTACHMENT_CACHE_TTL_MS) return cached.result;
  const result = await fetchAttachmentsUncached(account, messageId, attachmentId);
  if (result.ok) {
    attachmentCache.delete(key);
    attachmentCache.set(key, { at: Date.now(), result });
    while (attachmentCache.size > ATTACHMENT_CACHE_MAX_ENTRIES) attachmentCache.delete(attachmentCache.keys().next().value);
  }
  return result;
}

async function fetchAttachmentsUncached(account, messageId, attachmentId) {
  const targets = await resolveAccounts(account);
  for (const target of targets) {
    const auth = await safeGetBestToken(target, "mail");
    if (!auth?.token) continue;
    const first = auth.api || preferredApiForToken(decodeJwtClaims(auth.token));
    const attempt = async (api) => {
      const one = attachmentId ? `/${encodeURIComponent(attachmentId)}` : "";
      const url = api === "graph"
        ? new URL(`/v1.0/me/messages/${encodeURIComponent(messageId)}/attachments${one}`, ensureTrailingSlash(GRAPH_BASE))
        : new URL(`/api/v2.0/me/messages/${encodeURIComponent(messageId)}/attachments${one}`, ensureTrailingSlash(OUTLOOK_REST_BASE));
      const headers = api === "graph" ? graphHeaders(auth.token) : { authorization: `Bearer ${auth.token}`, accept: "application/json" };
      return await fetchJsonDetailed(url, { headers, timeoutMs: ATTACHMENT_TIMEOUT_MS, api: api === "graph" ? "graph" : "outlook_rest" });
    };
    let res = await attempt(first);
    if (!res.ok && shouldFallback(res)) res = await attempt(first === "graph" ? "outlook_rest" : "graph");
    if (!res.ok) {
      if (targets.length > 1 && res.status === 404) continue;
      return { ok: false, status: "attachments_failed", account: target.label, id: messageId, upstream_http_status: res.status, answer_text: "The Outlook upstream call failed." };
    }
    const values = attachmentId ? (res.body ? [res.body] : []) : Array.isArray(res.body?.value) ? res.body.value : [];
    return { ok: true, account: target.label, attachments: values.map((raw, i) => compactAttachment(raw, i)) };
  }
  return { ok: false, status: "message_not_found", scope: accountScopeLabel(account), id: messageId, answer_text: "Message not found." };
}

// Graph (camelCase) and Outlook REST (PascalCase) shapes -> one compact shape.
function compactAttachment(raw = {}, index) {
  const pick = (name) => raw[name] ?? raw[name.charAt(0).toUpperCase() + name.slice(1)];
  const type = String(raw["@odata.type"] || "").split(".").pop().toLowerCase();
  const kind = type.includes("file") ? "file" : type.includes("item") ? "item" : type.includes("reference") ? "reference" : type || "file";
  const contentType = String(pick("contentType") || "").toLowerCase();
  return {
    index,
    attachment_id: String(pick("id") || ""),
    name: String(pick("name") || `attachment-${index}`),
    content_type: contentType,
    size: Number(pick("size") || 0),
    is_inline: pick("isInline") === true,
    content_id: String(pick("contentId") || "").replace(/^<|>$/g, ""),
    kind,
    is_image: kind === "file" && contentType.startsWith("image/"),
    data_base64: kind === "file" ? String(pick("contentBytes") || "") : "",
  };
}

export async function outlookSearch({
  account,
  query,
  limit = DEFAULT_LIST_LIMIT,
} = {}) {
  if (!normalizeString(query)) return missingField("query", "A search query is required.");
  const targets = await resolveAccounts(account);
  const boundedLimit = clampInteger(limit, 1, MAX_LIST_LIMIT, DEFAULT_LIST_LIMIT);
  const lists = [];
  for (const target of targets) {
    const auth = await safeGetBestToken(target, "mail");
    if (!auth?.token) {
      lists.push({
        ok: false,
        account: target.label,
        status: "needs_sign_in",
        returned_count: 0,
        items: [],
      });
      continue;
    }
    const token = auth.token;
    const first = auth.api || preferredApiForToken(decodeJwtClaims(token));
    const attempt = async (api) => {
      if (api === "graph") {
        const url = new URL("/v1.0/me/messages", ensureTrailingSlash(GRAPH_BASE));
        const params = new URLSearchParams();
        params.set("$select", "id,subject,from,receivedDateTime,isRead,isDraft,conversationId,webLink");
        params.set("$search", `"${normalizeString(query)}"`);
        params.set("$top", String(boundedLimit));
        url.search = params.toString();
        const headers = graphHeaders(token, { needsConsistencyLevel: true });
        return await fetchJsonDetailed(url, { headers, timeoutMs: DEFAULT_TIMEOUT_MS, api: "graph" });
      } else {
        const url = new URL("/api/v2.0/me/messages", ensureTrailingSlash(OUTLOOK_REST_BASE));
        const params = new URLSearchParams();
        params.set("$select", "Id,Subject,From,ReceivedDateTime,IsRead,IsDraft,ConversationId,WebLink");
        params.set("$top", String(boundedLimit));
        params.set("$search", `"${normalizeString(query)}"`); // best-effort; some tenants may not support it
        url.search = params.toString();
        const headers = { authorization: `Bearer ${token}`, accept: "application/json" };
        return await fetchJsonDetailed(url, { headers, timeoutMs: DEFAULT_TIMEOUT_MS, api: "outlook_rest" });
      }
    };
    let res = await attempt(first);
    if (!res.ok && shouldFallback(res)) res = await attempt(first === "graph" ? "outlook_rest" : "graph");
    if (!res.ok) {
      lists.push({ ok: false, account: target.label, status: "search_failed", upstream_http_status: res.status, upstream_error: res.error || "", returned_count: 0, items: [] });
      continue;
    }
    const isGraph = res.api === "graph";
    const items = Array.isArray(res.body?.value) ? res.body.value : [];
    lists.push({ ok: true, account: target.label, status: "ok", returned_count: items.length, items: items.map((m) => (isGraph ? compactGraphMessage(m, target.label) : compactOutlookRestMessage(m, target.label))) });
  }
  const merged = mergeAccountLists(lists);
  return {
    ok: lists.some(l => l.ok),
    status: lists.some(l => l.ok) ? "ok" : "search_failed",
    scope: accountScopeLabel(account),
    returned_count: merged.length,
    items: merged,
    answer_text:
      merged.length === 0
        ? "No results."
        : `Found ${merged.length} result${merged.length === 1 ? "" : "s"}.`,
    accounts: lists,
  };
}

export async function outlookAgenda({
  account,
  // One of: { today: true }, { tomorrow: true }, or explicit { start_date, end_date } (ISO or YYYY-MM-DD)
  today = false,
  tomorrow = false,
  start_date,
  end_date,
  limit = 50,
} = {}) {
  const targets = await resolveAccounts(account);
  const { startLocal, endLocal } = normalizeEtLocalRange({ today, tomorrow, start_date, end_date });
  const boundedLimit = clampInteger(limit, 1, 200, 50);
  const lists = [];
  for (const target of targets) {
    const auth = await safeGetBestToken(target, "calendar");
    if (!auth?.token) {
      lists.push({
        ok: false,
        account: target.label,
        status: "needs_sign_in",
        returned_count: 0,
        items: [],
      });
      continue;
    }
    const token = auth.token;
    const first = auth.api || preferredApiForToken(decodeJwtClaims(token));
    const attempt = async (api) => {
      if (api === "graph") {
        const url = new URL("/v1.0/me/calendarView", ensureTrailingSlash(GRAPH_BASE));
        const params = new URLSearchParams();
        params.set("startDateTime", startLocal);
        params.set("endDateTime", endLocal);
        params.set("$top", String(boundedLimit));
        params.set("$orderby", "start/dateTime asc");
        params.set("$select", "id,subject,start,end,location,organizer,isAllDay,webLink");
        url.search = params.toString();
        const headers = graphHeaders(token, { preferOutlookTz: ET_TIMEZONE });
        return await fetchJsonDetailed(url, { headers, timeoutMs: DEFAULT_TIMEOUT_MS, api: "graph" });
      } else {
        const url = new URL("/api/v2.0/me/calendarview", ensureTrailingSlash(OUTLOOK_REST_BASE));
        const params = new URLSearchParams();
        params.set("startDateTime", startLocal);
        params.set("endDateTime", endLocal);
        params.set("$top", String(boundedLimit));
        params.set("$orderby", "Start/DateTime asc");
        params.set("$select", "Id,Subject,Start,End,Location,Organizer,IsAllDay,WebLink");
        url.search = params.toString();
        const headers = { authorization: `Bearer ${token}`, accept: "application/json", prefer: `outlook.timezone="${ET_TIMEZONE}"` };
        return await fetchJsonDetailed(url, { headers, timeoutMs: DEFAULT_TIMEOUT_MS, api: "outlook_rest" });
      }
    };
    let res = await attempt(first);
    if (!res.ok && shouldFallback(res)) res = await attempt(first === "graph" ? "outlook_rest" : "graph");
    if (!res.ok) {
      lists.push({ ok: false, account: target.label, status: "agenda_failed", upstream_http_status: res.status, upstream_error: res.error || "", returned_count: 0, items: [] });
      continue;
    }
    const isGraph = res.api === "graph";
    const values = Array.isArray(res.body?.value) ? res.body.value : [];
    lists.push({ ok: true, account: target.label, status: "ok", returned_count: values.length, items: values.map((e) => (isGraph ? compactGraphEvent(e, target.label) : compactOutlookRestEvent(e, target.label))) });
  }
  // Return a per-account map; do not merge calendar events across accounts into one list by default
  const result = {
    ok: true,
    status: "ok",
    scope: accountScopeLabel(account),
    start_date: startLocal,
    end_date: endLocal,
    accounts: lists.map((list) => ({
      account: list.account,
      ok: list.ok,
      status: list.status,
      returned_count: list.returned_count,
      items: list.items,
    })),
  };
  const total =
    result.accounts?.reduce((sum, a) => sum + (a.ok ? a.returned_count : 0), 0) || 0;
  result.answer_text =
    total === 0
      ? "No events in the selected range."
      : `${total} event${total === 1 ? "" : "s"} in the selected range.`;
  return result;
}

// Free/busy for one or more people, like Outlook's Scheduling Assistant. Uses the
// signed-in account's getSchedule, so it works for anyone in the same organization
// (and other organizations that publish free/busy) without them sharing their calendar.
// Returns each person's busy blocks (subject and location only when they share more
// than free/busy) and the slots inside the working day when everyone is free.
// Read-only. All times are US Eastern wall-clock times.
const FREEBUSY_MAX_PEOPLE = 20;

export async function outlookFreeBusy({
  account,
  emails,
  // Same range options as outlookAgenda; start_date/end_date may also be date-times (2026-11-17T11:00).
  today = false,
  tomorrow = false,
  start_date,
  end_date,
  day_start = "09:00",
  day_end = "17:00",
  min_minutes = 30,
  weekends = false,
} = {}) {
  const people = [...new Set(parseRecipients(emails).map((r) => r.EmailAddress.Address.toLowerCase()))];
  if (!people.length) {
    return { ok: false, status: "missing_emails", answer_text: "Give one or more email addresses, separated by commas." };
  }
  if (people.length > FREEBUSY_MAX_PEOPLE) {
    return { ok: false, status: "too_many_emails", answer_text: `Check at most ${FREEBUSY_MAX_PEOPLE} people at a time.` };
  }
  const target = freeBusyAccount(account, people);
  if (!target) {
    return {
      ok: false,
      status: "account_required",
      answer_text: `Say which account to check from: ${accountNames().join(", ") || "none set up"}.`,
    };
  }
  const withSeconds = (v) => (/T\d{2}:\d{2}$/.test(v) ? `${v}:00` : v);
  const range = normalizeEtLocalRange({ today, tomorrow, start_date, end_date });
  const startLocal = withSeconds(range.startLocal), endLocal = withSeconds(range.endLocal);
  const auth = await safeGetBestToken(target, "calendar");
  if (!auth?.token) return { ok: false, status: "needs_sign_in", account: target.label, answer_text: `The ${target.label} Outlook session needs signing in again.` };

  const windowsTz = "Eastern Standard Time";
  const res = await postJson(new URL("/api/v2.0/me/calendar/getschedule", ensureTrailingSlash(OUTLOOK_REST_BASE)), auth.token, {
    Schedules: people,
    StartTime: { DateTime: startLocal, TimeZone: windowsTz },
    EndTime: { DateTime: endLocal, TimeZone: windowsTz },
    AvailabilityViewInterval: 15,
  });
  if (!res.ok) {
    return { ok: false, status: "freebusy_failed", account: target.label, upstream_http_status: res.status, upstream_error: res.error, answer_text: "Outlook didn't return free/busy times." };
  }

  const schedules = Array.isArray(res.body?.value) ? res.body.value : [];
  const results = people.map((email) => {
    const sched = schedules.find((x) => String(x.ScheduleId || "").toLowerCase() === email) || {};
    if (sched.Error) return { email, ok: false, error: normalizeString(sched.Error.Message) || "unavailable", busy: [] };
    const busy = (sched.ScheduleItems || [])
      .filter((i) => i.Status && i.Status !== "Free")
      .map((i) => ({
        start: wallTime(i.Start?.DateTime),
        end: wallTime(i.End?.DateTime),
        status: String(i.Status).toLowerCase(),
        ...(i.Subject ? { subject: i.Subject } : {}),
        ...(i.Location ? { location: i.Location } : {}),
      }))
      .sort((a, b) => a.start.localeCompare(b.start));
    return { email, ok: true, busy: mergeBusy(busy) };
  });

  const known = results.filter((r) => r.ok);
  const freeForAll = known.length
    ? commonFreeSlots(known.flatMap((r) => r.busy), { startLocal, endLocal, dayStart: day_start, dayEnd: day_end, minMinutes: clampInteger(min_minutes, 5, 480, 30), weekends: toBoolean(weekends) })
    : [];
  const unavailable = results.filter((r) => !r.ok).map((r) => r.email);
  return {
    ok: known.length > 0,
    status: known.length ? "ok" : "freebusy_unavailable",
    account: target.label,
    timezone: ET_TIMEZONE,
    start: startLocal,
    end: endLocal,
    working_day: `${day_start}-${day_end}`,
    people: results,
    free_for_all: freeForAll,
    answer_text: [
      known.length ? `${freeForAll.length} slot${freeForAll.length === 1 ? "" : "s"} when ${known.length === 1 ? "they are" : "everyone is"} free (Eastern, ${day_start}-${day_end}).` : "No free/busy information came back.",
      unavailable.length ? `No free/busy for ${unavailable.join(", ")}.` : "",
    ].filter(Boolean).join(" "),
  };
}

// The account to query from: the one named, else the one whose domain matches the
// people being checked, else the only account.
function freeBusyAccount(account, people) {
  if (normalizeString(account) && normalizeString(account).toLowerCase() !== "all") return singleWorkAccount(account);
  const all = listAccounts();
  for (const email of people) {
    const key = resolveAccountAlias(email.replace(/^.*@/, ""));
    if (key) return getAccount(key);
  }
  return all.length === 1 ? all[0] : null;
}

async function postJson(url, token, payload) {
  const response = await fetch(url, {
    method: "POST",
    // Without the Prefer header, getSchedule answers in UTC whatever time zone the request used.
    headers: { authorization: `Bearer ${token}`, accept: "application/json", "content-type": "application/json", prefer: `outlook.timezone="${ET_TIMEZONE}"` },
    body: JSON.stringify(payload),
    redirect: "error",
    signal: AbortSignal.timeout(DEFAULT_TIMEOUT_MS),
  });
  const text = await response.text().catch(() => "");
  let body;
  try { body = JSON.parse(text); } catch { body = null; }
  return { ok: response.ok, status: response.status, body, error: !response.ok ? (body?.error?.message || text.slice(0, 200)) : "" };
}

// "2026-11-17T11:00:00.0000000" -> "2026-11-17T11:00"
function wallTime(value) {
  return String(value || "").slice(0, 16);
}

// Wall-clock "YYYY-MM-DDTHH:mm" <-> minutes, with no time zone math (everything is Eastern).
function wallMinutes(value) {
  const m = String(value).match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/);
  return m ? Date.UTC(+m[1], +m[2] - 1, +m[3], +m[4], +m[5]) / 60000 : NaN;
}

function minutesToWall(minutes) {
  return new Date(minutes * 60000).toISOString().slice(0, 16);
}

// Joins overlapping or touching busy blocks for one person.
function mergeBusy(blocks) {
  const out = [];
  for (const b of blocks) {
    const last = out[out.length - 1];
    if (last && b.start <= last.end) {
      if (b.end > last.end) last.end = b.end;
      if (last.status !== b.status) last.status = [last.status, b.status].includes("busy") ? "busy" : last.status;
      if (b.subject && !last.subject) last.subject = b.subject;
    } else out.push({ ...b });
  }
  return out;
}

function commonFreeSlots(busy, { startLocal, endLocal, dayStart, dayEnd, minMinutes, weekends }) {
  const hm = (v) => { const [h, m] = String(v).split(":").map(Number); return (h || 0) * 60 + (m || 0); };
  const ds = hm(dayStart), de = hm(dayEnd);
  const rangeStart = wallMinutes(startLocal), rangeEnd = wallMinutes(endLocal);
  if (!Number.isFinite(rangeStart) || !Number.isFinite(rangeEnd) || de <= ds) return [];
  const blocks = busy.map((b) => [wallMinutes(b.start), wallMinutes(b.end)]).filter(([a, b]) => b > a).sort((x, y) => x[0] - y[0]);
  const slots = [];
  for (let day = Math.floor(rangeStart / 1440) * 1440; day < rangeEnd; day += 1440) {
    const weekday = new Date(day * 60000).getUTCDay();
    if (!weekends && (weekday === 0 || weekday === 6)) continue;
    let cursor = Math.max(day + ds, rangeStart);
    const stop = Math.min(day + de, rangeEnd);
    for (const [bs, be] of blocks) {
      if (be <= cursor || bs >= stop) continue;
      if (bs - cursor >= minMinutes) slots.push([cursor, bs]);
      cursor = Math.max(cursor, be);
    }
    if (stop - cursor >= minMinutes) slots.push([cursor, stop]);
  }
  return slots.map(([a, b]) => ({ start: minutesToWall(a), end: minutesToWall(b), minutes: b - a }));
}

// ------------------------------
// Helpers
// ------------------------------

// Maps spoken/typed account names, labels and email domains to a work account key.
// Returns the account name, or null when the value is not a configured account
// (e.g. empty, "all", or a Gmail/Himalaya account name).
// ---------------------------------------------------------------------------
// Drafts. Mirrors Outlook web: Reply / Reply all create a draft in the thread
// with the quoted original below your text; New message saves to Drafts.
// These only create or update drafts. There is deliberately no send path.
// ---------------------------------------------------------------------------

export async function outlookDraftReply({
  account,
  id,
  body,
  bodyHtml,
  replyAll = false,
  addTo,
  addCc,
  inlineImages,
  confirmed = false,
} = {}) {
  const messageId = normalizeString(id);
  if (!messageId) return missingField("id", "An Outlook message id is required.");
  const content = draftContent({ body, bodyHtml });
  if (!content) return missingField("body", "Draft text is required.");
  const target = singleWorkAccount(account);
  if (!target) return missingField("account", `Pick one account: ${accountNames().join(", ") || "none configured"}.`);
  const all = toBoolean(replyAll);
  if (!toBoolean(confirmed)) {
    return draftConfirmationRequired(`Confirm that Andrew wants to save a ${all ? "reply-all" : "reply"} draft in ${target.label} Outlook.`);
  }
  const auth = await getDraftToken(target);
  if (!auth) return noDraftToken(target);

  const created = await restWrite(auth.token, "POST", `me/messages/${encodeURIComponent(messageId)}/${all ? "createreplyall" : "createreply"}`, {});
  if (!created.ok) return draftUpstreamError(target, created, "reply_draft_failed");
  const draft = created.body || {};
  const existing = draft.Body || {};
  const merged = prependToBody(existing, content);
  // Optional extra recipients (e.g. the rest of the team on a reply-all);
  // anyone already on the draft is kept once.
  const patch = { Body: merged };
  const extraTo = parseRecipients(addTo);
  const extraCc = parseRecipients(addCc);
  if (extraTo.length || extraCc.length) {
    const seen = new Set();
    const keep = (list) =>
      (list || []).filter((r) => {
        const key = String(r?.EmailAddress?.Address || "").toLowerCase();
        if (!key || seen.has(key)) return false;
        seen.add(key);
        return true;
      });
    patch.ToRecipients = keep([...(draft.ToRecipients || []), ...extraTo]);
    patch.CcRecipients = keep([...(draft.CcRecipients || []), ...extraCc]);
  }
  const patched = await restWrite(auth.token, "PATCH", `me/messages/${encodeURIComponent(draft.Id)}`, patch);
  if (!patched.ok) return draftUpstreamError(target, patched, "reply_draft_failed");
  const saved = patched.body || draft;
  const label = all ? "Reply-all draft saved" : "Reply draft saved";
  const attached = await attachInlineImages(auth.token, draft.Id, inlineImages);
  if (!attached.ok) return { ...draftResult(target, saved, `${label}, but an image couldn't be attached.`), ok: false, status: "inline_image_failed", error: attached.error };
  return draftResult(target, saved, attached.count ? `${label} with ${attached.count} inline image(s).` : `${label}.`);
}

export async function outlookDraftNew({
  account,
  to,
  cc,
  subject,
  body,
  bodyHtml,
  inlineImages,
  confirmed = false,
} = {}) {
  const toList = parseRecipients(to);
  if (!toList.length) return missingField("to", "At least one recipient is required.");
  const content = draftContent({ body, bodyHtml });
  if (!content) return missingField("body", "Draft text is required.");
  const target = singleWorkAccount(account);
  if (!target) return missingField("account", `Pick one account: ${accountNames().join(", ") || "none configured"}.`);
  if (!toBoolean(confirmed)) {
    return draftConfirmationRequired(`Confirm that Andrew wants to save a new draft in ${target.label} Outlook.`);
  }
  const auth = await getDraftToken(target);
  if (!auth) return noDraftToken(target);
  // POST /me/messages saves the message to Drafts without sending it.
  const created = await restWrite(auth.token, "POST", "me/messages", {
    Subject: normalizeString(subject),
    Body: content,
    ToRecipients: toList,
    CcRecipients: parseRecipients(cc),
  });
  if (!created.ok) return draftUpstreamError(target, created, "new_draft_failed");
  const draft = created.body || {};
  const attached = await attachInlineImages(auth.token, draft.Id, inlineImages);
  if (!attached.ok) return { ...draftResult(target, draft, "Draft saved, but an image couldn't be attached."), ok: false, status: "inline_image_failed", error: attached.error };
  return draftResult(target, draft, attached.count ? `Draft saved with ${attached.count} inline image(s).` : "Draft saved.");
}

// Inline images for a draft body that references them as <img src="cid:...">.
// Each image is fetched from an allowed host (OUTLOOK_INLINE_IMAGE_HOSTS) and
// attached with IsInline, so it shows without "download pictures".
// No host is allowed until one is set, so a draft can't be made to fetch arbitrary URLs.
let inlineImageHostsOverride = null;
export function setInlineImageHosts(hosts) {
  inlineImageHostsOverride = Array.isArray(hosts) ? hosts.map(String) : null;
}
function inlineImageHosts() {
  const list = inlineImageHostsOverride ?? String(process.env.OUTLOOK_INLINE_IMAGE_HOSTS || "").split(",");
  return new Set(list.map((h) => h.trim()).filter(Boolean));
}
const MAX_INLINE_IMAGE_BYTES = 3 * 1024 * 1024;

async function attachInlineImages(token, draftId, images) {
  const list = Array.isArray(images) ? images : [];
  for (const img of list) {
    const contentId = normalizeString(img?.content_id || img?.contentId);
    let url;
    try { url = new URL(String(img?.url || "")); } catch { return { ok: false, error: "Each inline image needs a valid url." }; }
    const local = url.hostname === "127.0.0.1" || url.hostname === "localhost";
    const allowed = inlineImageHosts();
    if (!contentId || !allowed.has(url.hostname) || (url.protocol !== "https:" && !local)) {
      return { ok: false, error: allowed.size ? `Inline image must have a content_id and an https url on ${[...allowed].join(", ")}.` : "Inline images are off: set OUTLOOK_INLINE_IMAGE_HOSTS to the image hosts to allow." };
    }
    const res = await fetch(url, { redirect: "error", signal: AbortSignal.timeout(DEFAULT_TIMEOUT_MS) }).catch(() => null);
    const type = res?.headers.get("content-type") || "";
    if (!res?.ok || !type.startsWith("image/")) return { ok: false, error: `Couldn't fetch image ${url.pathname}.` };
    const bytes = Buffer.from(await res.arrayBuffer());
    if (bytes.length > MAX_INLINE_IMAGE_BYTES) return { ok: false, error: `Image ${url.pathname} is over 3 MB.` };
    const saved = await restWrite(token, "POST", `me/messages/${encodeURIComponent(draftId)}/attachments`, {
      "@odata.type": "#Microsoft.OutlookServices.FileAttachment",
      Name: normalizeString(img?.name) || url.pathname.split("/").pop(),
      ContentType: type,
      ContentBytes: bytes.toString("base64"),
      ContentId: contentId,
      IsInline: true,
    });
    if (!saved.ok) return { ok: false, error: saved.error || "Outlook rejected the image." };
  }
  return { ok: true, count: list.length };
}

// Moves a draft to Deleted Items (recoverable). Anything that isn't a draft is refused.
export async function outlookDeleteDraft({ account, id, confirmed = false } = {}) {
  const messageId = normalizeString(id);
  if (!messageId) return missingField("id", "An Outlook draft id is required.");
  const target = singleWorkAccount(account);
  if (!target) return missingField("account", `Pick one account: ${accountNames().join(", ") || "none configured"}.`);
  if (!toBoolean(confirmed)) {
    return draftConfirmationRequired(`Confirm that Andrew wants to delete a draft in ${target.label} Outlook.`);
  }
  const auth = await getDraftToken(target);
  if (!auth) return noDraftToken(target);
  const path = `me/messages/${encodeURIComponent(messageId)}`;
  const found = await restWrite(auth.token, "GET", `${path}?$select=Subject,IsDraft`);
  if (!found.ok) return draftUpstreamError(target, found, "delete_draft_failed");
  if (found.body?.IsDraft !== true) {
    return {
      ok: false,
      status: "not_a_draft",
      account: target.label,
      answer_text: "That message isn't a draft, so it was left alone.",
    };
  }
  const moved = await restWrite(auth.token, "POST", `${path}/move`, { DestinationId: "DeletedItems" });
  if (!moved.ok) return draftUpstreamError(target, moved, "delete_draft_failed");
  return {
    ok: true,
    status: "draft_deleted",
    account: target.label,
    subject: String(found.body?.Subject || ""),
    answer_text: `Draft moved to ${target.label} Outlook Deleted Items.`,
  };
}

function singleWorkAccount(account) {
  const key = resolveAccountAlias(account);
  return key ? getAccount(key) : null;
}

async function getDraftToken(target) {
  const auth = await safeGetBestToken(target, "mail_write").catch(() => null);
  return auth?.token ? auth : null;
}

function draftContent({ body, bodyHtml }) {
  const html = normalizeString(bodyHtml);
  if (html) return { ContentType: "HTML", Content: html };
  const text = normalizeString(body);
  if (!text) return null;
  return { ContentType: "HTML", Content: textToHtml(text) };
}

function textToHtml(text) {
  const esc = String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
  return esc
    .split(/\n{2,}/)
    .map((p) => `<p>${p.replace(/\n/g, "<br>")}</p>`)
    .join("");
}

// Puts the new text above the quoted original, where Outlook web's cursor starts.
function prependToBody(existing, content) {
  const current = String(existing?.Content || "");
  if (String(existing?.ContentType || "").toLowerCase() === "text") {
    const plain = content.Content.replace(/<br>/g, "\n").replace(/<\/p><p>/g, "\n\n").replace(/<[^>]+>/g, "")
      .replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&");
    return { ContentType: "Text", Content: `${plain}\n\n${current}` };
  }
  const block = `<div>${content.Content}</div><br>`;
  const open = current.match(/<body[^>]*>/i);
  if (!open) return { ContentType: "HTML", Content: block + current };
  const at = open.index + open[0].length;
  return { ContentType: "HTML", Content: current.slice(0, at) + block + current.slice(at) };
}

function parseRecipients(value) {
  const items = Array.isArray(value) ? value : String(value || "").split(/[,;]/);
  return items
    .map((v) => (typeof v === "object" && v ? v : String(v).trim()))
    .filter(Boolean)
    .map((v) => {
      if (typeof v === "object") {
        const address = normalizeString(v.address || v.email);
        return address ? { EmailAddress: { Address: address, Name: normalizeString(v.name) || undefined } } : null;
      }
      const m = v.match(/^(.*?)<\s*([^>]+)\s*>$/);
      const address = (m ? m[2] : v).trim();
      const name = m ? m[1].replace(/^"|"$/g, "").trim() : "";
      return /@/.test(address) ? { EmailAddress: { Address: address, ...(name ? { Name: name } : {}) } } : null;
    })
    .filter(Boolean);
}

function recipientsOut(list) {
  return (list || []).map((r) => ({ name: r?.EmailAddress?.Name || "", address: r?.EmailAddress?.Address || "" }));
}

function draftResult(target, msg, answer) {
  return {
    ok: true,
    status: "draft_saved",
    account: target.label,
    id: String(msg.Id || ""),
    subject: String(msg.Subject || ""),
    to: recipientsOut(msg.ToRecipients),
    cc: recipientsOut(msg.CcRecipients),
    is_draft: msg.IsDraft !== false,
    web_link: String(msg.WebLink || ""),
    answer_text: `${answer} It's in ${target.label} Outlook Drafts and has not been sent.`,
  };
}

function draftConfirmationRequired(message) {
  return { ok: false, status: "confirmation_required", message, answer_text: message };
}

function noDraftToken(target) {
  return {
    ok: false,
    status: "needs_sign_in",
    account: target.label,
    answer_text: `The ${target.label} Outlook session has no token that can save drafts. Run outlook-web-cli auth status, then sign in again (outlook-web-cli auth chrome or auth cookies).`,
  };
}

function draftUpstreamError(target, res, status) {
  return {
    ok: false,
    status,
    account: target.label,
    upstream_http_status: res.status,
    upstream_error: String(res.error || "").slice(0, 300),
    answer_text: "Outlook didn't save the draft.",
  };
}

async function restWrite(token, method, path, payload) {
  if (/\/send(\?|$)/i.test(path)) throw new Error("Refusing to send mail: drafts only.");
  const url = new URL(`/api/v2.0/${path}`, ensureTrailingSlash(OUTLOOK_REST_BASE));
  const response = await fetch(url, {
    method,
    headers: { authorization: `Bearer ${token}`, accept: "application/json", "content-type": "application/json" },
    body: payload === undefined ? undefined : JSON.stringify(payload),
    redirect: "error",
    signal: AbortSignal.timeout(DEFAULT_TIMEOUT_MS),
  });
  const text = await response.text().catch(() => "");
  let body;
  try { body = JSON.parse(text); } catch { body = null; }
  return { ok: response.ok, status: response.status, body, error: !response.ok ? (body?.error?.message || text.slice(0, 200)) : "" };
}

export function resolveWorkAccountAlias(value) {
  return resolveAccountAlias(value);
}

export function listWorkAccounts() {
  return listAccounts().map((a) => ({ account: a.name, email_domain: a.emailDomain, provider: "outlook", access: a.access }));
}

function accountNames() {
  return listAccounts().map((a) => a.name);
}

async function resolveAccounts(scope) {
  const s = normalizeString(scope).toLowerCase() === "all" ? "all" : (resolveAccountAlias(scope) || normalizeString(scope));
  const all = listAccounts();
  if (!s || s === "all") {
    if (!all.length) throw new Error("No Outlook accounts are set up. Run `outlook-web-cli auth chrome --account <name> --cdp-port <port>` first.");
    return all;
  }
  const conf = getAccount(s);
  if (!conf) throw new Error(`Unknown account '${scope}'. Use ${[...accountNames(), "all"].join(", ")}.`);
  return [conf];
}

function accountScopeLabel(scope) {
  const s = normalizeString(scope);
  return s || "all";
}

function graphHeaders(token, { needsConsistencyLevel = false, preferOutlookTz } = {}) {
  const headers = {
    authorization: `Bearer ${token}`,
    accept: "application/json",
  };
  if (needsConsistencyLevel) headers["consistency-level"] = "eventual";
  if (preferOutlookTz) headers["prefer"] = `outlook.timezone="${preferOutlookTz}"`;
  return headers;
}

async function safeGetBestToken(target, purpose /* "mail" | "calendar" | "mail_write" */) {
  // Allow test override without a session
  const envOverride =
    process.env[`OUTLOOK_TEST_TOKEN_${target.label.toUpperCase().replace(/[^A-Z0-9]/g, "_")}`] ||
    process.env.OUTLOOK_TEST_TOKEN ||
    "";
  if (normalizeString(envOverride)) return { token: envOverride, api: preferredApiForToken(decodeJwtClaims(envOverride)) };
  // The saved web session's token is for https://outlook.office.com with the web app's
  // Mail.ReadWrite and Calendars scopes, so it serves every purpose.
  const session = await getSessionToken(target.label);
  if (session) return session;
  // Last resort: read a live token from the Outlook web tab, as earlier versions did.
  return target.cdpPort ? await getBestTokenViaCdp(target, purpose) : null;
}

async function getBestTokenViaCdp({ cdpPort, mainUrl }, purpose) {
  const tab = await findOutlookTab(cdpPort, mainUrl);
  if (!tab) return null;
  const cdp = await connectCdp(tab.webSocketDebuggerUrl);
  if (!cdp) return null;
  const expr = `
(function() {
  try {
    var out = [];
    for (var i = 0; i < localStorage.length; i++) {
      var k = localStorage.key(i) || "";
      var raw = null;
      try { raw = localStorage.getItem(k) || ""; } catch(e) {}
      if (!raw) continue;
      try {
        var obj = JSON.parse(raw);
        if (k.startsWith("msal.3|") && obj && obj.credentialType === "AccessToken" && typeof obj.secret === "string") {
          out.push({ source: "msal_v3", target: String(obj.target||""), expiresOn: Number(obj.expiresOn||0), token: obj.secret });
          continue;
        }
        if (k.startsWith("accesstoken-") && (obj.secret || obj.accessToken || obj.credential)) {
          out.push({ source: "msal_v2", target: String(obj.target||obj.scopes||""), expiresOn: Number(obj.expiresOn||obj.expires_on||0), token: obj.secret || obj.accessToken || obj.credential });
          continue;
        }
      } catch (e) {
        // ignore parse errors
      }
    }
    return { ok: true, tokens: out };
  } catch (e3) {
    return { ok: false, status: "token_probe_failed" };
  }
})()`.trim();
  const scan = async () => {
    const value = (await cdp.evaluate(expr)) || {};
    return value.ok && Array.isArray(value.tokens) ? value.tokens : [];
  };
  try {
    let best = pickBestToken(await scan(), purpose);
    if (!best) {
      // Reload once so MSAL can refresh silently, then look again
      await cdp.send("Page.reload", { ignoreCache: true }).catch(() => {});
      await new Promise((r) => setTimeout(r, 1500));
      best = pickBestToken(await scan(), purpose);
    }
    return best ? { token: best.token, api: best.api } : null;
  } catch {
    return null;
  } finally {
    cdp.close();
  }
}

function resolveMailFolder(folder) {
  const key = normalizeString(folder || "inbox").toLowerCase().replace(/[\s_-]+/g, "");
  return MAIL_FOLDERS[key] || null;
}

// The received-date clause goes first so `$orderby` on it is accepted with other filters.
function listFilter({ unread, from, since }, fields) {
  const sender = normalizeString(from);
  const sinceIso = toIsoUtcString(since, { assumeEtMidnight: true });
  const filters = [];
  if (sinceIso || sender) filters.push(`${fields.received} ge ${sinceIso || ALL_TIME_UTC}`);
  if (toBoolean(unread)) filters.push(`${fields.isRead} eq false`);
  if (sender) filters.push(`${fields.fromAddress} eq '${escapeODataLiteral(sender)}'`);
  return filters.join(" and ");
}

function compactGraphMessage(msg, account, { includeBody = false, maxBodyChars = DEFAULT_READ_MAX_BODY_CHARS } = {}) {
  const id = String(msg.id || "");
  const from = {
    name: String(msg.from?.emailAddress?.name || ""),
    address: String(msg.from?.emailAddress?.address || ""),
  };
  const receivedUtc = String(msg.receivedDateTime || "");
  const receivedEt = toEtLocalString(receivedUtc);
  const base = {
    account,
    id,
    subject: trimTo(String(msg.subject || ""), 400),
    from,
    is_unread: msg.isRead === false,
    is_draft: msg.isDraft === true,
    received_datetime: receivedUtc,
    received_at: receivedEt,
    web_link: String(msg.webLink || ""),
    conversation_id: String(msg.conversationId || ""),
  };
  if (includeBody) {
    const bodyType = String(msg.body?.contentType || "").toLowerCase();
    const raw = String(msg.body?.content || "");
    const text =
      bodyType === "text"
        ? raw
        : htmlToText(raw || "", { wordwrap: 120, selectors: [{ selector: "a", options: { hideLinkHrefIfSameAsText: true } }] });
    base.body_text = trimTo(text, maxBodyChars);
  }
  return base;
}

function compactGraphEvent(evt, account) {
  const start = asEtIso(evt.start);
  const end = asEtIso(evt.end);
  return {
    account,
    id: String(evt.id || ""),
    subject: trimTo(String(evt.subject || ""), 400),
    is_all_day: Boolean(evt.isAllDay),
    location: String(evt.location?.displayName || ""),
    organizer: {
      name: String(evt.organizer?.emailAddress?.name || ""),
      address: String(evt.organizer?.emailAddress?.address || ""),
    },
    start: start,
    end: end,
    web_link: String(evt.webLink || ""),
  };
}

function compactOutlookRestMessage(msg, account, { includeBody = false, maxBodyChars = DEFAULT_READ_MAX_BODY_CHARS } = {}) {
  // Outlook REST uses PascalCase fields
  const id = String(msg.Id || "");
  const from = {
    name: String(msg.From?.EmailAddress?.Name || ""),
    address: String(msg.From?.EmailAddress?.Address || ""),
  };
  const receivedUtc = String(msg.ReceivedDateTime || "");
  const receivedEt = toEtLocalString(receivedUtc);
  const base = {
    account,
    id,
    subject: trimTo(String(msg.Subject || ""), 400),
    from,
    is_unread: msg.IsRead === false,
    is_draft: msg.IsDraft === true,
    received_datetime: receivedUtc,
    received_at: receivedEt,
    web_link: String(msg.WebLink || ""),
    conversation_id: String(msg.ConversationId || ""),
  };
  if (includeBody) {
    const bodyType = String(msg.Body?.ContentType || "").toLowerCase();
    const raw = String(msg.Body?.Content || "");
    const text = bodyType === "text" ? raw : htmlToText(raw || "", { wordwrap: 120, selectors: [{ selector: "a", options: { hideLinkHrefIfSameAsText: true } }] });
    base.body_text = trimTo(text, maxBodyChars);
  }
  return base;
}

function compactOutlookRestEvent(evt, account) {
  const start = { dateTime: String(evt.Start?.DateTime || ""), timeZone: String(evt.Start?.TimeZone || ET_TIMEZONE) };
  const end = { dateTime: String(evt.End?.DateTime || ""), timeZone: String(evt.End?.TimeZone || ET_TIMEZONE) };
  return {
    account,
    id: String(evt.Id || ""),
    subject: trimTo(String(evt.Subject || ""), 400),
    is_all_day: Boolean(evt.IsAllDay),
    location: String(evt.Location?.DisplayName || ""),
    organizer: {
      name: String(evt.Organizer?.EmailAddress?.Name || ""),
      address: String(evt.Organizer?.EmailAddress?.Address || ""),
    },
    start,
    end,
    web_link: String(evt.WebLink || ""),
  };
}

function graphError(error, account, code = "graph_error") {
  const message = String(error?.message || "Graph error");
  return {
    ok: false,
    account,
    status: code,
    error_code: String(error?.code || ""),
    error_message: message.slice(0, 400),
  };
}

function mergeAccountLists(lists) {
  const merged = [];
  for (const list of lists) {
    if (!list.ok) continue;
    for (const item of list.items || []) merged.push(item);
  }
  merged.sort((a, b) => {
    const aT = Date.parse(a.received_datetime || a.start?.dateTime || 0) || 0;
    const bT = Date.parse(b.received_datetime || b.start?.dateTime || 0) || 0;
    return bT - aT;
  });
  return merged;
}

async function httpJson(url, { headers = {}, timeoutMs = DEFAULT_TIMEOUT_MS } = {}) {
  const response = await fetch(url, {
    method: "GET",
    headers,
    redirect: "error",
    signal: AbortSignal.timeout(timeoutMs),
  });
  let text = "";
  try { text = await response.text(); } catch {}
  let json;
  try { json = JSON.parse(text); } catch { json = { error: { code: "invalid_json", message: "Upstream did not return JSON." } }; }
  if (!response.ok) {
    const error = json?.error || { code: `HTTP_${response.status}`, message: text.slice(0, 400) };
    return { error };
  }
  return json;
}

async function fetchJsonDetailed(url, { headers = {}, timeoutMs = DEFAULT_TIMEOUT_MS, api } = {}) {
  const response = await fetch(url, { method: "GET", headers, redirect: "error", signal: AbortSignal.timeout(timeoutMs) });
  const text = await response.text().catch(() => "");
  let body;
  try { body = JSON.parse(text); } catch { body = null; }
  return { ok: response.ok, status: response.status, body, error: !response.ok ? (body?.error?.message || String(text).slice(0, 200)) : "", api: api || (String(url).includes("graph.microsoft.com") ? "graph" : "outlook_rest") };
}

function normalizeDateRangeEt({ today, tomorrow, start_date, end_date }) {
  if (toBoolean(today)) {
    const { start, end } = etDayBounds(0);
    return { startIso: start.toISOString(), endIso: end.toISOString() };
  }
  if (toBoolean(tomorrow)) {
    const { start, end } = etDayBounds(1);
    return { startIso: start.toISOString(), endIso: end.toISOString() };
  }
  const startIso = toIsoUtcString(start_date, { assumeEtMidnight: true });
  const endIso = toIsoUtcString(end_date, { assumeEtMidnight: false, endOfDay: true });
  if (startIso && endIso) return { startIso, endIso };
  // Default: today
  const { start, end } = etDayBounds(0);
  return { startIso: start.toISOString(), endIso: end.toISOString() };
}

function normalizeEtLocalRange({ today, tomorrow, start_date, end_date }) {
  // Return ET-local wall times as YYYY-MM-DDTHH:mm:ss strings for API calls that honor Prefer: outlook.timezone
  const nowMs = Number(process.env.OUTLOOK_TEST_NOW_MS || 0) || Date.now();
  const fmt = new Intl.DateTimeFormat("en-CA", { timeZone: ET_TIMEZONE, year: "numeric", month: "2-digit", day: "2-digit" });
  const ymd = (offset) => {
    const d = new Date(nowMs + (offset || 0) * 24 * 60 * 60 * 1000);
    const parts = fmt.formatToParts(d);
    const y = parts.find(p => p.type === "year")?.value;
    const m = parts.find(p => p.type === "month")?.value;
    const da = parts.find(p => p.type === "day")?.value;
    return `${y}-${m}-${da}`;
  };
  if (toBoolean(today)) {
    const d = ymd(0);
    return { startLocal: `${d}T00:00:00`, endLocal: `${d}T23:59:59` };
  }
  if (toBoolean(tomorrow)) {
    const d = ymd(1);
    return { startLocal: `${d}T00:00:00`, endLocal: `${d}T23:59:59` };
  }
  const s = normalizeString(start_date);
  const e = normalizeString(end_date);
  const startLocal = s && /^\d{4}-\d{2}-\d{2}$/.test(s) ? `${s}T00:00:00` : s || "";
  const endLocal = e && /^\d{4}-\d{2}-\d{2}$/.test(e) ? `${e}T23:59:59` : e || "";
  if (startLocal && endLocal) return { startLocal, endLocal };
  const d = ymd(0);
  return { startLocal: `${d}T00:00:00`, endLocal: `${d}T23:59:59` };
}

function etDayBounds(offsetDays) {
  const now = new Date();
  // Compute ET midnight for current date + offset
  const options = { timeZone: ET_TIMEZONE, year: "numeric", month: "2-digit", day: "2-digit" };
  const parts = new Intl.DateTimeFormat("en-CA", options).formatToParts(now);
  const year = Number(parts.find((p) => p.type === "year")?.value);
  const month = Number(parts.find((p) => p.type === "month")?.value);
  const day = Number(parts.find((p) => p.type === "day")?.value) + Number(offsetDays || 0);
  // Build a Date in ET, then convert to UTC ISO
  const base = new Date(Date.UTC(year, month - 1, day, 0, 0, 0));
  const start = shiftUtcFromEt(base, 0, 0, 0);
  const end = shiftUtcFromEt(base, 23, 59, 59, 999);
  return { start, end };
}

function shiftUtcFromEt(dateUtc, h, m, s, ms = 0) {
  // dateUtc is the ET calendar date at 00:00 UTC; we need to offset by the ET UTC offset for that calendar day.
  const et = new Date(dateUtc);
  et.setUTCHours(h, m, s, ms);
  // Determine ET offset by formatting the target time in ET and parsing back the offset via toLocaleString with timeZoneName.
  const fmt = new Intl.DateTimeFormat("en-US", {
    timeZone: ET_TIMEZONE,
    hour12: false,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
  });
  // The Date object et is UTC; we just return it — Graph will accept the UTC ISO range.
  return et;
}

function toEtLocalString(isoUtc) {
  if (!isoUtc) return "";
  const d = new Date(isoUtc);
  const f = new Intl.DateTimeFormat("en-CA", {
    timeZone: ET_TIMEZONE,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hour12: false,
  });
  // en-CA yields YYYY-MM-DD, HH:mm:ss
  const parts = f.formatToParts(d);
  const as = (t) => parts.find((p) => p.type === t)?.value || "";
  return `${as("year")}-${as("month")}-${as("day")}T${as("hour")}:${as("minute")}:${as("second")}`;
}

function asEtIso(g) {
  // Graph calendarView returns start/end as { dateTime, timeZone }. Respect that if present.
  if (!g) return { dateTime: "", timeZone: ET_TIMEZONE };
  const tz = g.timeZone || ET_TIMEZONE;
  const dt = String(g.dateTime || "");
  return { dateTime: dt, timeZone: tz };
}

function decodeJwtClaims(token) {
  try {
    const [, payload] = String(token).split(".");
    if (!payload) return {};
    const b64 = payload.replace(/-/g, "+").replace(/_/g, "/") + "===".slice((payload.length + 3) % 4);
    const json = Buffer.from(b64, "base64").toString("utf8");
    return JSON.parse(json);
  } catch {
    return {};
  }
}

function preferredApiForToken(claims) {
  const aud = String(claims?.aud || "").toLowerCase();
  if (aud.includes("graph.microsoft.com") || aud === "00000003-0000-0000-c000-000000000000") return "graph";
  if (aud.includes("outlook.office.com") || aud === "00000002-0000-0ff1-ce00-000000000000" || aud.includes("substrate.office.com")) return "outlook_rest";
  return "graph"; // default
}

function shouldFallback(res) {
  return !res.ok && (res.status === 401 || res.status === 403 || res.status === 404);
}

function ensureTrailingSlash(base) {
  return base.endsWith("/") ? base : base + "/";
}

function pickBestToken(tokens, purpose) {
  const nowSec = Math.floor(Date.now() / 1000);
  const margin = 60;
  const need = purpose === "calendar" ? "Calendars.Read" : purpose === "mail_write" ? "Mail.ReadWrite" : "Mail.Read";
  const filtered = tokens
    .map(t => ({ ...t, target: String(t.target || ""), expiresOn: Number(t.expiresOn || 0) }))
    .filter(t => t.token && t.token.length > 100);
  const valid = filtered.filter(t => t.expiresOn > nowSec + margin);
  const byPref = (domain) =>
    valid
      .filter(t => t.target.includes(domain) && t.target.includes(need))
      .sort((a, b) => b.expiresOn - a.expiresOn)[0];
  const outlook = byPref("outlook.office.com");
  if (outlook) return { token: outlook.token, api: "outlook_rest", target: outlook.target, expiresOn: outlook.expiresOn };
  // Drafts are written through the Outlook REST API only.
  if (purpose === "mail_write") return null;
  const graph = byPref("graph.microsoft.com");
  if (graph) return { token: graph.token, api: "graph", target: graph.target, expiresOn: graph.expiresOn };
  return null;
}

// Test hook for msal.3 parsing/selection
export function __test_selectBestTokenFromMsalEntries(entries, purpose) {
  return pickBestToken(entries, purpose);
}

function toIsoUtcString(input, { assumeEtMidnight = false, endOfDay = false } = {}) {
  const s = normalizeString(input);
  if (!s) return "";
  // If only a date is supplied (YYYY-MM-DD), treat it as ET midnight boundaries.
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  if (m) {
    const year = Number(m[1]);
    const month = Number(m[2]) - 1;
    const day = Number(m[3]);
    const h = endOfDay ? 23 : 0;
    const mi = endOfDay ? 59 : 0;
    const se = endOfDay ? 59 : 0;
    const d = new Date(Date.UTC(year, month, day, h, mi, se, endOfDay ? 999 : 0));
    return d.toISOString();
  }
  // Otherwise expect an ISO with zone or plain ISO — pass through.
  try {
    const d = new Date(s);
    return d.toISOString();
  } catch {
    return "";
  }
}

function trimTo(s, n) {
  if (typeof s !== "string") return "";
  if (s.length <= n) return s;
  return s.slice(0, Math.max(0, n - 3)) + "...";
}

function normalizeUrl(s) {
  if (!s || typeof s !== "string") return "";
  return s.trim();
}

function clampInteger(value, min, max, fallback) {
  const v = Number(value);
  if (!Number.isFinite(v)) return fallback;
  return Math.min(Math.max(Math.trunc(v), min), max);
}

function toBoolean(value, fallback = false) {
  if (value === true || value === "true" || value === 1 || value === "1") return true;
  if (value === false || value === "false" || value === 0 || value === "0") return false;
  return Boolean(fallback);
}

function normalizeString(value, fallback = "") {
  if (typeof value !== "string") return fallback;
  const s = value.trim();
  return s || fallback;
}

function escapeODataLiteral(s) {
  return String(s).replace(/'/g, "''");
}

function missingField(field, message) {
  return { ok: false, status: "missing_field", field, message };
}

