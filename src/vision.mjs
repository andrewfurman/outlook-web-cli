// Optional image reader for `attachment-inspect`: one OpenRouter chat call with the image.
// Programs that embed this package usually pass their own analyzer to setVisionAnalyzer().
export function openRouterVisionAnalyzer({ apiKey = process.env.OPENROUTER_API_KEY, model = process.env.OUTLOOK_WEB_VISION_MODEL || "google/gemini-2.5-flash" } = {}) {
  if (!String(apiKey || "").trim()) return null;
  return async ({ prompt, images }) => {
    const res = await fetch("https://openrouter.ai/api/v1/chat/completions", {
      method: "POST",
      headers: { authorization: `Bearer ${apiKey}`, "content-type": "application/json" },
      body: JSON.stringify({
        model,
        messages: [{
          role: "user",
          content: [
            { type: "text", text: `${prompt}\n\nReply as JSON: {"description": "...", "ocr_text": "..."}` },
            ...images.map((img) => ({ type: "image_url", image_url: { url: `data:${img.media_type};base64,${img.data_base64}` } })),
          ],
        }],
        response_format: { type: "json_object" },
      }),
      signal: AbortSignal.timeout(60_000),
    });
    const json = await res.json().catch(() => ({}));
    if (!res.ok) return { ok: false, status: "vision_failed", http_status: res.status, message: String(json?.error?.message || "").slice(0, 200) };
    const content = String(json?.choices?.[0]?.message?.content || "");
    let parsed = {};
    try { parsed = JSON.parse(content.replace(/^```(?:json)?\s*|\s*```$/g, "")); } catch { parsed = { description: content }; }
    return { ok: true, provider: "openrouter", model, description: String(parsed.description || ""), ocr_text: String(parsed.ocr_text || "") };
  };
}
