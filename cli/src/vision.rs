//! Optional image reader for `attachment-inspect`: one OpenRouter chat call with the
//! image. Used when OPENROUTER_API_KEY is set; OUTLOOK_WEB_VISION_MODEL picks the model.

use crate::ctx::{Ctx, VisionAnalyzer};
use crate::js;
use serde_json::{json, Value};
use std::time::Duration;

const DEFAULT_MODEL: &str = "google/gemini-2.5-flash";

pub fn openrouter_analyzer(ctx: &Ctx) -> Option<VisionAnalyzer> {
    let api_key = ctx.var("OPENROUTER_API_KEY")?.trim().to_string();
    if api_key.is_empty() {
        return None;
    }
    let model = ctx
        .var("OUTLOOK_WEB_VISION_MODEL")
        .unwrap_or(DEFAULT_MODEL)
        .to_string();
    let base = ctx
        .var("OPENROUTER_BASE_URL")
        .unwrap_or("https://openrouter.ai/api/v1")
        .trim_end_matches('/')
        .to_string();
    let http = ctx.http.clone();
    Some(Box::new(move |req: &Value| {
        let prompt = js::get_str(req, &["prompt"]);
        let mut content = vec![json!({
            "type": "text",
            "text": format!("{prompt}\n\nReply as JSON: {{\"description\": \"...\", \"ocr_text\": \"...\"}}"),
        })];
        for img in req
            .get("images")
            .and_then(Value::as_array)
            .into_iter()
            .flatten()
        {
            content.push(json!({
                "type": "image_url",
                "image_url": { "url": format!("data:{};base64,{}", js::get_str(img, &["media_type"]), js::get_str(img, &["data_base64"])) },
            }));
        }
        let body = json!({
            "model": model,
            "messages": [{ "role": "user", "content": content }],
            "response_format": { "type": "json_object" },
        });
        let res = http
            .post(format!("{base}/chat/completions"))
            .bearer_auth(&api_key)
            .header("content-type", "application/json")
            .body(body.to_string())
            .timeout(Duration::from_secs(60))
            .send()?;
        let status = res.status();
        let json: Value = res
            .text()
            .ok()
            .and_then(|t| serde_json::from_str(&t).ok())
            .unwrap_or_else(|| json!({}));
        if !status.is_success() {
            let message: String = js::get_str(&json, &["error", "message"])
                .chars()
                .take(200)
                .collect();
            return Ok(json!({
                "ok": false, "status": "vision_failed",
                "http_status": status.as_u16(), "message": message,
            }));
        }
        let text = json
            .get("choices")
            .and_then(|c| c.get(0))
            .map(|c| js::get_str(c, &["message", "content"]))
            .unwrap_or_default();
        let parsed = serde_json::from_str::<Value>(strip_fence(&text))
            .ok()
            .filter(Value::is_object)
            .unwrap_or_else(|| json!({ "description": text }));
        Ok(json!({
            "ok": true,
            "provider": "openrouter",
            "model": model,
            "description": js::get_str(&parsed, &["description"]),
            "ocr_text": js::get_str(&parsed, &["ocr_text"]),
        }))
    }))
}

fn strip_fence(s: &str) -> &str {
    let mut t = s;
    if let Some(rest) = t.strip_prefix("```") {
        t = rest.strip_prefix("json").unwrap_or(rest).trim_start();
    }
    if let Some(rest) = t.trim_end().strip_suffix("```") {
        t = rest.trim_end();
    }
    t
}
