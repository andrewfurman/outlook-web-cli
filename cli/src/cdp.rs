//! Minimal Chrome DevTools Protocol client for one Outlook web tab. Used only to
//! import the web session (and, as a last resort, to read a live token from the tab).

use crate::ctx::Ctx;
use anyhow::{anyhow, Result};
use regex::Regex;
use serde_json::{json, Value};
use std::net::{SocketAddr, TcpStream};
use std::sync::LazyLock;
use std::time::{Duration, Instant};
use tungstenite::{Message, WebSocket};

static OUTLOOK_MAIL_RE: LazyLock<Regex> = LazyLock::new(|| {
    Regex::new(r"^https://outlook\.(cloud\.microsoft|office\.com|office365\.com)/mail/").unwrap()
});

#[derive(Clone, Debug)]
pub struct Tab {
    pub url: String,
    pub web_socket_debugger_url: String,
}

/// The tab showing `main_url`, else any Outlook web mail tab.
pub fn find_outlook_tab(ctx: &Ctx, cdp_port: u16, main_url: Option<&str>) -> Option<Tab> {
    let res = ctx
        .http
        .get(format!("http://127.0.0.1:{cdp_port}/json/list"))
        .timeout(Duration::from_secs(3))
        .send()
        .ok()?;
    let targets: Value = serde_json::from_str(&res.text().ok()?).ok()?;
    let pages: Vec<Tab> = targets
        .as_array()?
        .iter()
        .filter(|t| t.get("type").and_then(Value::as_str) == Some("page"))
        .filter_map(|t| {
            let url = t.get("url")?.as_str()?.to_string();
            let ws = t.get("webSocketDebuggerUrl")?.as_str()?.to_string();
            (!ws.is_empty()).then_some(Tab {
                url,
                web_socket_debugger_url: ws,
            })
        })
        .collect();
    let main_url = main_url.filter(|m| !m.is_empty());
    pages
        .iter()
        .find(|t| main_url.is_some_and(|m| t.url == m || t.url.starts_with(m)))
        .or_else(|| pages.iter().find(|t| OUTLOOK_MAIL_RE.is_match(&t.url)))
        .cloned()
}

pub struct Cdp {
    ws: WebSocket<TcpStream>,
    next_id: u64,
    timeout: Duration,
}

/// Opens the tab's DevTools websocket, or None when it can't be reached in 2 seconds.
pub fn connect_cdp(web_socket_debugger_url: &str) -> Option<Cdp> {
    let url = url::Url::parse(web_socket_debugger_url).ok()?;
    let host = url.host_str()?.trim_matches(|c| c == '[' || c == ']');
    let port = url.port_or_known_default()?;
    let addr: SocketAddr = std::net::ToSocketAddrs::to_socket_addrs(&(host, port))
        .ok()?
        .next()?;
    let stream = TcpStream::connect_timeout(&addr, Duration::from_secs(2)).ok()?;
    stream.set_read_timeout(Some(Duration::from_secs(2))).ok()?;
    let (ws, _) = tungstenite::client(web_socket_debugger_url, stream).ok()?;
    Some(Cdp {
        ws,
        next_id: 0,
        timeout: Duration::from_secs(4),
    })
}

impl Cdp {
    pub fn send(&mut self, method: &str, params: Value) -> Result<Value> {
        self.next_id += 1;
        let id = self.next_id;
        let msg = json!({ "id": id, "method": method, "params": params });
        self.ws.send(Message::text(msg.to_string()))?;
        let deadline = Instant::now() + self.timeout;
        loop {
            let left = deadline.saturating_duration_since(Instant::now());
            if left.is_zero() {
                return Err(anyhow!("cdp_timeout"));
            }
            self.ws
                .get_ref()
                .set_read_timeout(Some(left.max(Duration::from_millis(10))))?;
            let frame = match self.ws.read() {
                Ok(f) => f,
                Err(tungstenite::Error::Io(e))
                    if matches!(
                        e.kind(),
                        std::io::ErrorKind::WouldBlock | std::io::ErrorKind::TimedOut
                    ) =>
                {
                    return Err(anyhow!("cdp_timeout"))
                }
                Err(e) => return Err(e.into()),
            };
            let Message::Text(text) = frame else { continue };
            let Ok(reply) = serde_json::from_str::<Value>(text.as_str()) else {
                continue;
            };
            if reply.get("id").and_then(Value::as_u64) != Some(id) {
                continue;
            }
            if let Some(err) = reply.get("error") {
                let message = err
                    .get("message")
                    .and_then(Value::as_str)
                    .unwrap_or("unknown");
                return Err(anyhow!("cdp_error: {message}"));
            }
            return Ok(reply.get("result").cloned().unwrap_or_else(|| json!({})));
        }
    }

    /// Runs JavaScript in the page and returns its value.
    pub fn evaluate(&mut self, expression: &str) -> Result<Value> {
        let result = self.send(
            "Runtime.evaluate",
            json!({ "expression": expression, "returnByValue": true, "awaitPromise": true }),
        )?;
        Ok(result
            .get("result")
            .and_then(|r| r.get("value"))
            .cloned()
            .unwrap_or(Value::Null))
    }
}

impl Drop for Cdp {
    fn drop(&mut self) {
        let _ = self.ws.close(None);
        let _ = self.ws.flush();
    }
}
