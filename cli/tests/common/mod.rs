//! A tiny fake HTTP server (Outlook REST, Graph, Microsoft sign-in) that records every
//! request, plus helpers to build a test context with its own config directory.
#![allow(dead_code)]

use base64::Engine;
use outlook_web_cli::Ctx;
use serde_json::{json, Value};
use std::sync::{Arc, Mutex};

#[derive(Clone, Debug)]
pub struct Req {
    pub method: String,
    pub path: String,
    pub query: Vec<(String, String)>,
    pub headers: Vec<(String, String)>,
    pub body: String,
}

impl Req {
    pub fn q(&self, k: &str) -> Option<String> {
        self.query
            .iter()
            .find(|(n, _)| n == k)
            .map(|(_, v)| v.clone())
    }
    pub fn header(&self, k: &str) -> Option<String> {
        self.headers
            .iter()
            .find(|(n, _)| n.eq_ignore_ascii_case(k))
            .map(|(_, v)| v.clone())
    }
    pub fn form(&self, k: &str) -> Option<String> {
        url::form_urlencoded::parse(self.body.as_bytes())
            .find(|(n, _)| n == k)
            .map(|(_, v)| v.into_owned())
    }
    pub fn json(&self) -> Value {
        serde_json::from_str(&self.body).unwrap_or(Value::Null)
    }
}

pub struct Reply {
    pub status: u16,
    pub headers: Vec<(String, String)>,
    pub body: Vec<u8>,
}

impl Reply {
    pub fn json(status: u16, v: Value) -> Reply {
        Reply {
            status,
            headers: vec![(
                "content-type".into(),
                "application/json; charset=utf-8".into(),
            )],
            body: v.to_string().into_bytes(),
        }
    }
    pub fn not_found() -> Reply {
        Reply::json(
            404,
            json!({ "error": { "code": "not_found", "message": "not found" } }),
        )
    }
    pub fn header(mut self, k: &str, v: &str) -> Reply {
        self.headers.push((k.into(), v.into()));
        self
    }
}

pub struct Fake {
    pub base: String,
    pub log: Arc<Mutex<Vec<Req>>>,
    server: Arc<tiny_http::Server>,
}

impl Fake {
    pub fn start(handler: impl Fn(&Req) -> Reply + Send + 'static) -> Fake {
        let server = Arc::new(tiny_http::Server::http("127.0.0.1:0").unwrap());
        let port = server.server_addr().to_ip().unwrap().port();
        let log = Arc::new(Mutex::new(Vec::new()));
        let (s, l) = (server.clone(), log.clone());
        std::thread::spawn(move || {
            for mut r in s.incoming_requests() {
                let mut body = String::new();
                let _ = r.as_reader().read_to_string(&mut body);
                let full = url::Url::parse(&format!("http://x{}", r.url())).unwrap();
                let req = Req {
                    method: r.method().as_str().to_string(),
                    path: full.path().to_string(),
                    query: full.query_pairs().into_owned().collect(),
                    headers: r
                        .headers()
                        .iter()
                        .map(|h| {
                            (
                                h.field.as_str().as_str().to_string(),
                                h.value.as_str().to_string(),
                            )
                        })
                        .collect(),
                    body,
                };
                l.lock().unwrap().push(req.clone());
                let reply = handler(&req);
                let mut resp =
                    tiny_http::Response::from_data(reply.body).with_status_code(reply.status);
                for (k, v) in reply.headers {
                    resp.add_header(
                        tiny_http::Header::from_bytes(k.as_bytes(), v.as_bytes()).unwrap(),
                    );
                }
                let _ = r.respond(resp);
            }
        });
        Fake {
            base: format!("http://127.0.0.1:{port}"),
            log,
            server,
        }
    }
    pub fn requests(&self) -> Vec<Req> {
        self.log.lock().unwrap().clone()
    }
    pub fn clear(&self) {
        self.log.lock().unwrap().clear();
    }
    pub fn count(&self) -> usize {
        self.log.lock().unwrap().len()
    }
    pub fn last(&self) -> Req {
        self.log.lock().unwrap().last().cloned().expect("a request")
    }
}

impl Drop for Fake {
    fn drop(&mut self) {
        self.server.unblock();
    }
}

pub fn b64(v: &Value) -> String {
    base64::engine::general_purpose::URL_SAFE_NO_PAD.encode(v.to_string())
}

/// An unsigned JWT with these claims.
pub fn jwt(claims: Value) -> String {
    format!("{}.{}.sig", b64(&json!({ "alg": "none" })), b64(&claims))
}

pub fn rest_token() -> String {
    jwt(json!({ "aud": "https://outlook.office.com", "scp": "Mail.ReadWrite Calendars.ReadWrite" }))
}

pub fn graph_token() -> String {
    jwt(json!({ "aud": "https://graph.microsoft.com", "scp": "Mail.Read" }))
}

/// A temp config dir with `config.json` holding these accounts.
pub fn config_dir(accounts: Value) -> tempfile::TempDir {
    let dir = tempfile::tempdir().unwrap();
    std::fs::write(
        dir.path().join("config.json"),
        json!({ "accounts": accounts }).to_string(),
    )
    .unwrap();
    dir
}

/// The contoso + fabrikam accounts most suites use.
pub fn two_accounts() -> tempfile::TempDir {
    config_dir(json!([
        { "name": "contoso", "aliases": ["contoso.com", "contoso work"], "username": "me@contoso.com" },
        { "name": "fabrikam", "aliases": ["fabrikam.com"], "username": "me@fabrikam.com" },
    ]))
}

pub fn ctx(dir: &tempfile::TempDir, vars: &[(&str, &str)]) -> Ctx {
    let mut all: Vec<(String, String)> = vec![(
        "OUTLOOK_WEB_CONFIG_DIR".into(),
        dir.path().display().to_string(),
    )];
    all.extend(vars.iter().map(|(k, v)| (k.to_string(), v.to_string())));
    Ctx::with_vars(all)
}

pub fn body(v: Value) -> serde_json::Map<String, Value> {
    v.as_object().cloned().unwrap()
}
