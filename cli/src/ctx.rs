//! What a command runs with: environment settings, the HTTP client, the optional
//! image model and the in-memory attachment cache. `Ctx::from_env()` is the real
//! process environment; tests build one from a map so they never share globals.

use anyhow::Result;
use serde_json::Value;
use std::cell::RefCell;
use std::collections::HashMap;
use std::path::PathBuf;
use std::time::{Duration, Instant};

/// An image model for `attachment-inspect`: takes
/// `{provider, prompt, images: [{media_type, data_base64, filename}]}` and returns
/// `{ok, provider, model, description, ocr_text, ...}`.
pub type VisionAnalyzer = Box<dyn Fn(&Value) -> Result<Value>>;

pub struct Ctx {
    vars: HashMap<String, String>,
    pub vision: Option<VisionAnalyzer>,
    pub(crate) attachment_cache: RefCell<Vec<(String, Instant, Value)>>,
    pub(crate) http: reqwest::blocking::Client,
}

impl Ctx {
    pub fn from_env() -> Self {
        let mut ctx = Self::with_vars(std::env::vars());
        ctx.vision = crate::vision::openrouter_analyzer(&ctx);
        ctx
    }

    /// A context with exactly these variables (nothing read from the process).
    pub fn with_vars<I, K, V>(vars: I) -> Self
    where
        I: IntoIterator<Item = (K, V)>,
        K: Into<String>,
        V: Into<String>,
    {
        let http = reqwest::blocking::Client::builder()
            .redirect(reqwest::redirect::Policy::none())
            .no_proxy()
            .connect_timeout(Duration::from_secs(10))
            .build()
            .expect("HTTP client");
        Ctx {
            vars: vars
                .into_iter()
                .map(|(k, v)| (k.into(), v.into()))
                .collect(),
            vision: None,
            attachment_cache: RefCell::new(Vec::new()),
            http,
        }
    }

    pub fn with_vision(mut self, analyzer: VisionAnalyzer) -> Self {
        self.vision = Some(analyzer);
        self
    }

    /// A variable's value, or None when unset or empty.
    pub fn var(&self, key: &str) -> Option<&str> {
        self.vars
            .get(key)
            .map(String::as_str)
            .filter(|v| !v.is_empty())
    }

    /// `$OUTLOOK_WEB_CONFIG_DIR`, else `~/.config/outlook-web-cli` on every platform.
    pub fn config_dir(&self) -> PathBuf {
        match self.var("OUTLOOK_WEB_CONFIG_DIR") {
            Some(dir) => PathBuf::from(dir),
            None => dirs::home_dir()
                .unwrap_or_else(|| PathBuf::from("."))
                .join(".config")
                .join("outlook-web-cli"),
        }
    }

    pub fn sessions_dir(&self) -> PathBuf {
        self.config_dir().join("sessions")
    }

    /// Milliseconds since the epoch, or `OUTLOOK_TEST_NOW_MS` for the calendar range.
    pub fn now_ms_for_ranges(&self) -> i64 {
        self.var("OUTLOOK_TEST_NOW_MS")
            .and_then(|v| v.trim().parse::<f64>().ok())
            .filter(|v| *v != 0.0 && v.is_finite())
            .map(|v| v as i64)
            .unwrap_or_else(now_ms)
    }
}

pub fn now_ms() -> i64 {
    chrono::Utc::now().timestamp_millis()
}

pub fn now_sec_f() -> f64 {
    now_ms() as f64 / 1000.0
}

/// `new Date().toISOString()`.
pub fn now_iso() -> String {
    chrono::Utc::now().to_rfc3339_opts(chrono::SecondsFormat::Millis, true)
}
