// Minimal Chrome DevTools Protocol client for one Outlook web tab. Used only to import
// the web session (and, as a last resort, to read a live token from the tab).

export async function findOutlookTab(cdpPort, mainUrl) {
  let targets;
  try {
    const res = await fetch(`http://127.0.0.1:${cdpPort}/json/list`, { redirect: "error", signal: AbortSignal.timeout(3000) });
    targets = await res.json();
  } catch {
    return null;
  }
  if (!Array.isArray(targets)) return null;
  const pages = targets.filter((t) => t.type === "page" && typeof t.url === "string" && t.webSocketDebuggerUrl);
  return pages.find((t) => mainUrl && (t.url === mainUrl || t.url.startsWith(mainUrl)))
    || pages.find((t) => /^https:\/\/outlook\.(cloud\.microsoft|office\.com|office365\.com)\/mail\//.test(t.url))
    || null;
}

export async function connectCdp(webSocketDebuggerUrl, { timeoutMs = 4000 } = {}) {
  const ws = new WebSocket(webSocketDebuggerUrl);
  const opened = await new Promise((resolve) => {
    ws.addEventListener("open", () => resolve(true));
    ws.addEventListener("error", () => resolve(false));
    setTimeout(() => resolve(false), 2000);
  });
  if (!opened) return null;
  let nextId = 0;
  const pending = new Map();
  ws.addEventListener("message", (ev) => {
    try {
      const msg = JSON.parse(ev.data);
      const waiter = pending.get(msg.id);
      if (waiter) {
        pending.delete(msg.id);
        waiter(msg);
      }
    } catch {
      // ignore non-JSON frames
    }
  });
  function send(method, params = {}) {
    const id = ++nextId;
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        pending.delete(id);
        reject(new Error("cdp_timeout"));
      }, timeoutMs);
      pending.set(id, (msg) => {
        clearTimeout(timer);
        if (msg.error) reject(new Error(`cdp_error: ${msg.error.message || "unknown"}`));
        else resolve(msg.result || {});
      });
      ws.send(JSON.stringify({ id, method, params }));
    });
  }
  return {
    send,
    async evaluate(expression) {
      const result = await send("Runtime.evaluate", { expression, returnByValue: true, awaitPromise: true });
      return result?.result?.value;
    },
    close() {
      try { ws.close(); } catch {}
    },
  };
}
