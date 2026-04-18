# Improvements Backlog

Tracked ideas to harden the project. Each entry: **problem → fix → tradeoff → status**.

When an item ships, move it to the **Done** section at the bottom AND add a short note to `PROJECT_OVERVIEW.md` under a "Recent Improvements" section.

---

## Status Legend
- `[ ]` open
- `[~]` in progress
- `[x]` done — also reflected in `PROJECT_OVERVIEW.md`

---

## Architecture

### 1. [ ] OAuth2 / MSAL instead of user+password
**Problem:** `UserCredential(email, password)` stores password in `.env`. Breaks on MFA. Manual rotation.
**Fix:** MSAL client-credentials flow. Register app in Entra ID. Cert-based auth. Least-privilege Graph scopes (`Sites.Read.All` or site-scoped `Sites.Selected`).
**Tradeoff:** Requires tenant admin consent. More setup vs. dramatically better security posture.

### 2. [ ] Microsoft Graph API instead of Office365-REST-Python-Client
**Problem:** Legacy SOAP-ish client. No batch endpoints. No delta sync.
**Fix:** Migrate to Graph (`/drives/{id}/root/children`, `/search/query`). Use `/drives/{id}/root/delta` for incremental sync.
**Tradeoff:** Rewrite every tool body. Unblocks caching + batching.

### 3. [ ] Deploy MCP server (drop ngrok dependency)
**Problem:** Dev-only topology. Anthropic cannot reach `localhost`.
**Fix:** Azure Container Apps / Cloud Run. Bearer auth on `/mcp` endpoint via `authorization_token` in Anthropic's `mcp_servers` config.
**Tradeoff:** Hosting cost + CI/CD setup. Enables real users.

---

## Performance

### 4. [x] Caching layer
**Problem:** Every `list_files` / `search_*` / metadata call hits SharePoint. Same folder listed N times per session.
**Fix:** LRU (in-process) or Redis, keyed on `folder_url`. TTL 60s. Read-only workload → no invalidation complexity.
**Tradeoff:** Staleness risk. Mitigate with short TTL + manual bust on write tools (none yet).

### 5. [ ] OCR for scanned PDFs
**Problem:** `search_content` returns nothing for scanned PDFs (no text layer indexed by SharePoint).
**Fix:** Background job: Azure Document Intelligence or Tesseract → extract text → store in sidecar index (SQLite / vector DB).
**Tradeoff:** One-time ingest cost vs. permanent searchability. Must handle new uploads.

### 6. [ ] Vector search / RAG
**Problem:** KQL is substring-only. Misses semantic matches ("recurrence" vs "recursive formula").
**Fix:** Chunk PDFs on ingest → embed (voyage-3 or `text-embedding-3-small`) → pgvector / Qdrant. Expose `semantic_search(query, top_k)` tool.
**Tradeoff:** Infra overhead. Hybrid BM25 + dense is the sweet spot (better recall than either alone).

---

## Correctness & Observability

### 7. [ ] Tool call observability
**Problem:** No metrics. No tracing. Errors → string → Claude, nothing else.
**Fix:** OpenTelemetry spans per `@mcp.tool()` call. Log: tool name, args, duration, HTTP status, bytes. Ship to App Insights / Grafana.
**Tradeoff:** Minor latency from span creation. Big win: p99 visibility, retry budget tuning.

### 8. [x] Streaming backpressure + cancellation
**Problem:** Session switch mid-stream → `ChatBox` remounts, but upstream `fetch` keeps reading → Anthropic keeps generating → wasted tokens = $.
**Fix (client):** `AbortController` in `ChatBox`; `abort()` in cleanup effect.
**Fix (server):** `route.ts` listens to `req.signal`; forward abort to upstream Anthropic fetch via its own `AbortController`.
**Tradeoff:** None really — pure win. Tricky: verify `ReadableStream` controller closes cleanly on abort.

### 9. [x] Rate limiting / concurrency control
**Problem:** No backpressure against SharePoint. Claude can fire many parallel `read_pdf` → SharePoint 429 → cascade failure.
**Fix:** `asyncio.Semaphore` (or thread-pool equivalent) per tool. Global concurrency cap (e.g. 4). Retry with exponential backoff + jitter on 429.
**Tradeoff:** Slightly slower bulk operations. Prevents outage.

---

## UX

### 10. [ ] Render tool inputs + outputs, not just names
**Problem:** Tool badges show name only. User sees "called read_pdf" — can't verify.
**Fix:** Expandable panel per tool call → args + truncated output.
**Tradeoff:** UI real estate. Solves trust + debuggability.

### 11. [ ] Citations with deep links
**Problem:** Claude quotes a file in prose → no way to jump to source.
**Fix:** Emit clickable URL → `https://.../Shared Documents/path#page=3`. SharePoint supports `#page=N` anchors for PDFs.
**Tradeoff:** None. Pure UX win.

### 12. [ ] Multi-turn context compaction
**Problem:** Long sessions blow context window. Every message re-sent in full.
**Fix:** After N messages, summarize older turns via cheap Haiku pass. Keep last K verbatim.
**Tradeoff:** Risk of losing detail. Mitigate with "show full history" toggle.

---

## Data Layer

### 13. [ ] Server-side session persistence
**Problem:** `localStorage` traps sessions on one browser. Cache clear → data loss. No cross-device.
**Fix:** Postgres + NextAuth (or Clerk). Sessions keyed on user ID.
**Tradeoff:** Infra cost. localStorage was right initially (YAGNI) — trigger = first user asks "can I access from my phone".

---

## Testing

### 14. [ ] Integration test harness
**Problem:** `test_tools.py` imports `from server` — file was renamed to `tools.py`. Broken.
**Fix:** pytest + VCR.py → record SharePoint responses once, replay offline. Runs in CI without creds.
**Tradeoff:** Cassettes rot if SharePoint API changes. Re-record periodically.

---

## Done

- **4. Redis caching layer** — `cache.py` adds `_with_cache` decorator (Redis GET/SET, TTL-per-tool, error-skip, size cap). Stacked `@mcp.tool() → @_with_cache → @_with_backpressure` on all 7 tools. Hits short-circuit semaphore + SharePoint entirely. `cache_stats()` MCP tool exposes in-process counters. Graceful degradation: Redis down → pass-through, no exception raised. See `PROJECT_OVERVIEW.md` → Recent Improvements.
- **8. Streaming backpressure + cancellation** — `route.ts` now forwards `req.signal` aborts to both the upstream Anthropic `fetch` (via inner `AbortController`) and the `ReadableStream` reader (`reader.cancel()` + loop guard). Client already aborted on unmount. See `PROJECT_OVERVIEW.md` → Recent Improvements.
- **9. Rate limiting / concurrency control** — `tools.py` adds `_with_backpressure` decorator: `threading.Semaphore(4)` concurrency cap + exponential-backoff retry with jitter on 429/503. Applied to all 7 tools. Tunable via `SHAREPOINT_MAX_CONCURRENCY` env var. See `PROJECT_OVERVIEW.md` → Recent Improvements.
