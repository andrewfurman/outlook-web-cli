import {
  outlookAgenda,
  outlookAttachmentInspect,
  outlookAttachments,
  outlookDeleteDraft,
  outlookDraftNew,
  outlookDraftReply,
  outlookMessages,
  outlookRead,
  outlookSearch,
  outlookStatus,
} from "./outlook.mjs";

// Command name -> handler taking a snake_case (or camelCase) body, as the CLI's flags and
// --json produce. Draft commands do nothing until context.confirmed is true.
const id = (body) => body.id || body.message_id || body.messageId;
const text = (body) => body.body || body.message || body.text;
const html = (body) => body.body_html || body.bodyHtml || body.html;

function draftReply(body, context, replyAll) {
  return outlookDraftReply({
    account: body.account,
    id: id(body),
    body: text(body),
    bodyHtml: html(body),
    replyAll,
    addTo: body.add_to || body.addTo,
    addCc: body.add_cc || body.addCc,
    inlineImages: body.inline_images || body.inlineImages,
    confirmed: context.confirmed,
  });
}

export const commands = {
  status: (body) => outlookStatus({ account: body.account }),
  messages: (body) => outlookMessages({
    account: body.account,
    limit: body.limit || body.max_results || body.maxResults,
    unread: body.unread,
    from: body.from || body.sender,
    since: body.since || body.start_date || body.startDate,
    search: body.search || body.query,
    folder: body.folder,
    maxRawBytes: body.max_raw_bytes || body.maxRawBytes,
  }),
  read: (body) => outlookRead({ account: body.account, id: id(body), maxBodyChars: body.max_body_chars || body.maxBodyChars }),
  search: (body) => outlookSearch({ account: body.account, query: body.query || body.search, limit: body.limit || body.max_results }),
  agenda: (body) => outlookAgenda({
    account: body.account,
    today: body.today,
    tomorrow: body.tomorrow,
    start_date: body.start_date || body.startDate,
    end_date: body.end_date || body.endDate,
    limit: body.limit || body.max_results,
  }),
  attachments: (body) => outlookAttachments({
    account: body.account,
    id: id(body),
    includeData: body.include_data ?? body.includeData,
    maxDataBytes: body.max_data_bytes || body.maxDataBytes,
    imagesOnly: body.images_only ?? body.imagesOnly,
    index: body.index ?? body.attachment_index,
    contentId: body.content_id || body.contentId || body.cid,
    attachmentId: body.attachment_id || body.attachmentId,
    offset: body.offset,
    length: body.length,
  }),
  "attachment-inspect": (body) => outlookAttachmentInspect({
    account: body.account,
    id: id(body),
    index: body.index ?? body.attachment_index,
    contentId: body.content_id || body.contentId || body.cid,
    prompt: body.prompt,
  }),
  "draft-reply": (body, context = {}) => draftReply(body, context, body.reply_all ?? body.replyAll),
  "draft-new": (body, context = {}) => outlookDraftNew({
    account: body.account,
    to: body.to,
    cc: body.cc,
    subject: body.subject,
    body: text(body),
    bodyHtml: html(body),
    inlineImages: body.inline_images || body.inlineImages,
    confirmed: context.confirmed,
  }),
  "draft-delete": (body, context = {}) => outlookDeleteDraft({ account: body.account, id: id(body), confirmed: context.confirmed }),
};
