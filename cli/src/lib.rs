//! Outlook mail and calendar through your Outlook web session. The `outlook-web-cli`
//! binary is the product; this library exists so the commands can be tested in-process.

pub mod accounts;
pub mod cdp;
pub mod commands;
pub mod ctx;
pub mod drafts;
pub mod freebusy;
pub mod html;
pub mod js;
pub mod outlook;
pub mod session;
pub mod vision;

pub use ctx::Ctx;
