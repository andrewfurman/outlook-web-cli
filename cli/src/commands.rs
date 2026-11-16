//! Command name -> handler taking a snake_case (or camelCase) body, as the CLI's flags
//! and `--json` produce. Draft commands do nothing until `confirmed` is true.

use crate::ctx::Ctx;
use crate::drafts::{self, NewArgs, ReplyArgs};
use crate::freebusy::{self, FreeBusyArgs};
use crate::js::{nn, or, Body};
use crate::outlook::{self, AttachmentArgs, MessagesArgs, RangeArgs};
use anyhow::{bail, Result};
use serde_json::Value;

pub const COMMANDS: &[&str] = &[
    "status",
    "messages",
    "read",
    "search",
    "agenda",
    "freebusy",
    "attachments",
    "attachment-inspect",
    "draft-reply",
    "draft-new",
    "draft-delete",
];

pub fn run(ctx: &Ctx, name: &str, body: &Body, confirmed: bool) -> Result<Value> {
    let g = |k: &str| body.get(k).cloned();
    let o = |keys: &[&str]| or(body, keys);
    let n = |keys: &[&str]| nn(body, keys);
    let account = g("account");
    let id = o(&["id", "message_id", "messageId"]);
    let text = o(&["body", "message", "text"]);
    let html = o(&["body_html", "bodyHtml", "html"]);
    let images = o(&["inline_images", "inlineImages"]);
    match name {
        "status" => outlook::outlook_status(ctx, account.as_ref()),
        "messages" => {
            let limit = o(&["limit", "max_results", "maxResults"]);
            let from = o(&["from", "sender"]);
            let since = o(&["since", "start_date", "startDate"]);
            let search = o(&["search", "query"]);
            let unread = g("unread");
            let folder = g("folder");
            outlook::outlook_messages(
                ctx,
                MessagesArgs {
                    account: account.as_ref(),
                    limit: limit.as_ref(),
                    unread: unread.as_ref(),
                    from: from.as_ref(),
                    since: since.as_ref(),
                    search: search.as_ref(),
                    folder: folder.as_ref(),
                },
            )
        }
        "read" => {
            let max = o(&["max_body_chars", "maxBodyChars"]);
            outlook::outlook_read(ctx, account.as_ref(), id.as_ref(), max.as_ref())
        }
        "search" => {
            let q = o(&["query", "search"]);
            let limit = o(&["limit", "max_results"]);
            outlook::outlook_search(ctx, account.as_ref(), q.as_ref(), limit.as_ref())
        }
        "agenda" => {
            let (today, tomorrow) = (g("today"), g("tomorrow"));
            let start = o(&["start_date", "startDate"]);
            let end = o(&["end_date", "endDate"]);
            let limit = o(&["limit", "max_results"]);
            outlook::outlook_agenda(
                ctx,
                account.as_ref(),
                RangeArgs {
                    today: today.as_ref(),
                    tomorrow: tomorrow.as_ref(),
                    start_date: start.as_ref(),
                    end_date: end.as_ref(),
                },
                limit.as_ref(),
            )
        }
        "freebusy" => {
            let emails = o(&["emails", "email", "people", "attendees"]);
            let (today, tomorrow) = (g("today"), g("tomorrow"));
            let start = o(&["start_date", "startDate", "start"]);
            let end = o(&["end_date", "endDate", "end"]);
            let day_start = o(&["day_start", "dayStart"]);
            let day_end = o(&["day_end", "dayEnd"]);
            let min = o(&["min_minutes", "minMinutes"]);
            let weekends = g("weekends");
            freebusy::outlook_free_busy(
                ctx,
                FreeBusyArgs {
                    account: account.as_ref(),
                    emails: emails.as_ref(),
                    range: RangeArgs {
                        today: today.as_ref(),
                        tomorrow: tomorrow.as_ref(),
                        start_date: start.as_ref(),
                        end_date: end.as_ref(),
                    },
                    day_start: day_start.as_ref(),
                    day_end: day_end.as_ref(),
                    min_minutes: min.as_ref(),
                    weekends: weekends.as_ref(),
                },
            )
        }
        "attachments" => {
            let include = n(&["include_data", "includeData"]);
            let max = o(&["max_data_bytes", "maxDataBytes"]);
            let images_only = n(&["images_only", "imagesOnly"]);
            let index = n(&["index", "attachment_index"]);
            let cid = o(&["content_id", "contentId", "cid"]);
            let aid = o(&["attachment_id", "attachmentId"]);
            let (offset, length) = (g("offset"), g("length"));
            outlook::outlook_attachments(
                ctx,
                AttachmentArgs {
                    account: account.as_ref(),
                    id: id.as_ref(),
                    include_data: include.as_ref(),
                    max_data_bytes: max.as_ref(),
                    images_only: images_only.as_ref(),
                    index: index.as_ref(),
                    content_id: cid.as_ref(),
                    attachment_id: aid.as_ref(),
                    offset: offset.as_ref(),
                    length: length.as_ref(),
                },
            )
        }
        "attachment-inspect" => {
            let index = n(&["index", "attachment_index"]);
            let cid = o(&["content_id", "contentId", "cid"]);
            let prompt = g("prompt");
            outlook::outlook_attachment_inspect(
                ctx,
                account.as_ref(),
                id.as_ref(),
                index.as_ref(),
                cid.as_ref(),
                prompt.as_ref(),
            )
        }
        "draft-reply" => {
            let reply_all = n(&["reply_all", "replyAll"]);
            let add_to = o(&["add_to", "addTo"]);
            let add_cc = o(&["add_cc", "addCc"]);
            drafts::outlook_draft_reply(
                ctx,
                ReplyArgs {
                    account: account.as_ref(),
                    id: id.as_ref(),
                    body: text.as_ref(),
                    body_html: html.as_ref(),
                    reply_all: reply_all.as_ref(),
                    add_to: add_to.as_ref(),
                    add_cc: add_cc.as_ref(),
                    inline_images: images.as_ref(),
                    confirmed,
                },
            )
        }
        "draft-new" => {
            let (to, cc, subject) = (g("to"), g("cc"), g("subject"));
            drafts::outlook_draft_new(
                ctx,
                NewArgs {
                    account: account.as_ref(),
                    to: to.as_ref(),
                    cc: cc.as_ref(),
                    subject: subject.as_ref(),
                    body: text.as_ref(),
                    body_html: html.as_ref(),
                    inline_images: images.as_ref(),
                    confirmed,
                },
            )
        }
        "draft-delete" => {
            drafts::outlook_delete_draft(ctx, account.as_ref(), id.as_ref(), confirmed)
        }
        other => bail!("Unknown command '{other}'. Run outlook-web-cli help."),
    }
}
