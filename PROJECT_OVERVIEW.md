# SharePoint AI Assistant — Project Overview

A two-repo system that turns a Microsoft SharePoint document library into a conversational knowledge base. A user chats in plain English; Claude answers by searching and reading actual SharePoint files in real time.

---

## Repos

| Repo | Role | Stack |
|---|---|---|
| `sharepoint-agent` | **MCP server** — exposes SharePoint as a set of tools (list, search, read PDFs/DOCX/PPTX) | Python, FastMCP, Office365-REST-Python-Client, PyMuPDF |
| `sharepoint-next` | **Chat UI + proxy** — Next.js app that streams Claude responses and wires Claude to the MCP server | Next.js 16, React 19, TypeScript, Tailwind v4, Anthropic Messages API |

---

## High-Level Architecture

```
┌──────────────┐    fetch + SSE    ┌──────────────────┐    HTTPS     ┌────────────────┐
│ Browser (UI) │ ───────────────▶  │ Next.js /api/chat│ ───────────▶ │ Anthropic API  │
│  ChatBox     │ ◀───────────────  │   (proxy)        │ ◀─────────── │ (Claude)       │
└──────────────┘   custom SSE      └──────────────────┘   SSE stream └────────┬───────┘
                                                                              │ MCP
                                                                              ▼
                                                    ┌───────────────────────────────┐
                                                    │ sharepoint-agent (FastMCP)    │
                                                    │  streamable-http / stdio      │
                                                    └────────────┬──────────────────┘
                                                                 │ Office365 REST
                                                                 ▼
                                                    ┌───────────────────────────────┐
                                                    │ Microsoft SharePoint library  │
                                                    └───────────────────────────────┘
```

Key: browser never touches SharePoint or Anthropic directly. API key stays server-side. SharePoint credentials stay inside the MCP server.

---

## Interesting Concepts

### 1. MCP (Model Context Protocol)

Open standard. LLM gets uniform interface to external data/tools. Claude reads tool schemas at inference time → decides when + how to call them. No hardcoded SharePoint logic in the chat app.

Two transports used here:
- **stdio** (`server_stdio.py`) — Claude Code spawns process, talks over stdin/stdout. No port. Local dev only.
- **streamable-http** (`server_http.py`) — public URL required. Anthropic's infra calls back into the MCP server. Tunneled via `ngrok` during dev.

Anthropic side enables MCP via:
```ts
requestBody.mcp_servers = [{ type: "url", url, name: "sharepoint" }];
headers["anthropic-beta"] = "mcp-client-2025-04-04";
```

Key insight: Claude → MCP loop runs **inside** the Anthropic request. Next.js proxy just observes events; it doesn't drive the tool loop.

### 2. HTTP Streaming (SSE pipeline)

Two SSE hops, re-encoded in the middle:

**Hop 1 — Anthropic → Next.js route.**
Raw Anthropic SSE has 6+ event types (`message_start`, `content_block_start`, `content_block_delta`, `content_block_stop`, `message_delta`, `message_stop`, `error`). Text arrives as `text_delta`. Tool-use inputs arrive as `input_json_delta` — fragmented JSON across multiple events.

**Hop 2 — Next.js route → browser.**
Route collapses Anthropic's format into **2 event types**:
- `event: text` → `{ text: "chunk" }`
- `event: tool` → `{ tool: "name", input: {...} }` (emitted on `content_block_stop` after reassembling fragmented JSON)
- `event: error` → `{ error: "msg" }`

Implementation detail: `blockMeta = Map<index, { tool, partialJson }>`. Tool block's JSON input accumulates across deltas keyed by `content_block` index. Full object emitted only at `content_block_stop`.

Browser consumes via `res.body.getReader()` + `TextDecoder`. Text chunks patched **in place** on a placeholder assistant message (`setMessages` mutates content at `placeholderIndex`) — no ref concatenation, React stays in control.

### 3. SharePoint Context Reuse

`office365.ClientContext` is expensive: auth handshake per construction. Solution:
```python
_ctx: ClientContext | None = None

def _get_ctx():
    global _ctx
    if _ctx: return _ctx
    _ctx = ClientContext(SITE_URL).with_credentials(UserCredential(email, pw))
    _ctx.load(_ctx.web); _ctx.execute_query()
    return _ctx
```
Module-level singleton. First tool call warms it; rest reuse. `warmup()` called at server start → eager auth, fails loud before first request.

### 4. Course Registry via MCP Instructions

`FastMCP("sharepoint-agent", instructions=_CLAUDE_MD)` injects the full `claude.md` (course → path map) into the **server's MCP `instructions` field**. Claude receives it as system context when it connects. No browsing needed for known courses — Claude maps course name → server-relative URL directly.

### 5. Smart PDF Reading

Most SharePoint course files are scanned PDFs (>1MB). Full-document read = context blown. Mitigation:
- `read_pdf(pages="1-5", max_chars=8000)` — default to bounded page range + char cap
- Header prepends `Pages | Language | Size` so Claude knows scale before deciding next step
- `_detect_language()` flags Hebrew (U+0590–U+05FF) vs Latin → RTL content handled correctly
- PyMuPDF `TEXT_PRESERVE_LIGATURES` flag → math notation survives extraction

### 6. Chat Session Persistence (No DB)

`useChatSessions` hook. Single source of truth. Stored in `localStorage` under `sharepoint-chat-sessions`. Cap: 50 sessions (FIFO eviction).

SSR trap avoided: `localStorage` doesn't exist on server. Hook initializes empty, hydrates in `useEffect`, exposes `hydrated` flag. Root page renders spinner until `hydrated === true` → no hydration mismatch.

Session switch resets stream state via **remount trick**: `<ChatBox key={activeSessionId} />`. Different key = React unmounts old component (kills in-flight reader) and mounts fresh one. Cleanest reset.

---

## Tools Exposed by `sharepoint-agent` (In Depth)

All tools registered via `@mcp.tool()` decorator on the `FastMCP` instance. Schemas auto-generated from type hints + docstrings.

### `list_files(folder_url)`
Lists files + subdirs of a SharePoint folder. Uses `ctx.web.get_folder_by_server_relative_url(...)` then `ctx.load(folder, ["Files", "Folders"])`. Filters out `Forms/` (SharePoint's internal form folder — always noise). Output format: `[FILE] name (bytes)` and `[DIR] name/ -> path`. Path emitted next to dir entry so Claude can feed it straight back into the next call.

### `read_file_content(file_url)`
Downloads raw bytes via `file_obj.download(buf).execute_query()`. Decode chain: UTF-8 → Latin-1 → "[binary content]" fallback. Text-only; wrong tool for PDFs/DOCX.

### `search_files(query, folder_url)`
Filename substring search. Uses SharePoint's native **KQL** via `SearchService`:
```
filename:"query" path:"https://.../sites/..."
```
`row_limit=500`, `trim_duplicates=False`. Path-scoped so hits stay inside the target site. Returns paths only — no content.

### `search_content(query, folder_url, file_types, max_results)`
Full-text search inside file **bodies** (SharePoint indexes content). KQL pattern:
```
"query" path:"..." (FileExtension:pdf OR FileExtension:docx)
```
Key caveat: **scanned PDFs don't index** (no OCR layer) → this tool returns nothing for them. For scanned content, fall back to `read_pdf` directly. Row cells parsed defensively — supports both `{Key, Value}` objects and dict shapes returned by different Office365 library versions.

### `read_pdf(file_url, pages, max_chars)`
Downloads bytes → `pymupdf.open(stream=..., filetype="pdf")`. Page range parser supports `"all"` / `"3"` / `"1-5"`. Text extracted with `TEXT_PRESERVE_LIGATURES` (keeps math ligatures, fi/fl glyphs). Emits per-page headers `--- Page N ---` so Claude can cite locations. Language auto-detected. Capped at `max_chars` (default 8000).

### `read_docx(file_url, max_chars)`
`python-docx`. Walks the document's XML body element (`doc.element.body`), dispatching on tag:
- `p` (paragraph) → checks style: `Heading 1-9` → markdown `#`/`##`/... prefix; body text as plain line
- `tbl` (table) → rows joined with `|` → markdown-ish table

Reading order preserved because XML traversal is DOM-order (text and tables interleaved as authored).

### `read_pptx(file_url, slides, max_chars)`
`python-pptx`. Iterates slides → iterates shapes → grabs every `shape.text_frame.paragraphs` text. Titles + bullets come out flat, in shape order. `--- Slide N ---` headers between slides. Same slide-range parser as `read_pdf`.

### `get_file_metadata(file_url, include_versions)`
Loads `Name, Length, TimeCreated, TimeLastModified, UIVersionLabel` from the file. Then loads `Author` + `Editor` from `listItemAllFields` (SharePoint splits file props from list-item props — author lives on the list item). Fallback extractor handles both dict (`{Title, LoginName}`) and raw string shapes.

Bonus: `_extract_grade()` regex-parses filenames matching `..._NN.ext` or `..._NNN.ext` → pulls numeric suffix (course convention: students encode grade in filename). Optional version history via `file_obj.versions`.

---

## Error Handling

Single `_format_error(tool, target, exc)` helper. Detects `403`/`404` in exception string → human-readable message. Otherwise generic `[tool] Error accessing 'target': msg`. Every tool wraps its body in `try/except → _format_error`. Claude receives the error as a normal string result → it can choose to retry with a different path, apologize, or escalate. No unhandled exceptions leak to the MCP protocol layer.

---

## Environment & Secrets

| Var | Where | Why |
|---|---|---|
| `SHAREPOINT_EMAIL` / `SHAREPOINT_PASSWORD` | `sharepoint-agent/.env` | `UserCredential` auth to SharePoint site |
| `ANTHROPIC_API_KEY` | `sharepoint-next/.env.local` | Server-only. No `NEXT_PUBLIC_` → never shipped to browser |
| `SHAREPOINT_MCP_URL` | `sharepoint-next/.env.local` | Public ngrok URL of running MCP server. Omit → plain Claude, no tools |

---

## Run Topology (Dev)

3 processes:
1. `python server_http.py` → MCP on `localhost:8000/mcp`
2. `ngrok http 8000` → public HTTPS tunnel
3. `npm run dev` in `sharepoint-next` → `localhost:3000`

Update `SHAREPOINT_MCP_URL` in `.env.local` each time ngrok restarts (free-tier URL rotates).

---

## Why This Design

- **MCP over hardcoded API**: swap SharePoint for Google Drive → rewrite server, chat app untouched
- **Server-side proxy**: API key never leaves server; also lets us reshape Anthropic's SSE into a simpler UI contract
- **Streaming**: perceived latency ~0 for first token; tool calls surface as badges mid-response
- **Singleton context**: avoids ~300ms auth round-trip per tool call
- **Bounded reads (`max_chars`, `pages`)**: prevents blowing Claude's context window on multi-MB scanned PDFs

---

## Recent Improvements

Shipped items from `IMPROVEMENTS.md`. Each links back to the underlying motivation.

### Redis caching layer (#4)
**Problem:** Every SharePoint tool call pays auth round-trip + Office365 REST fetch. Claude re-lists the same folders and re-reads the same PDFs across turns.
**Fix (`cache.py` + `tools.py`):** New `_with_cache` decorator sits between `@mcp.tool()` and `@_with_backpressure`. On a cache hit, the semaphore is never acquired and SharePoint is never touched — pure in-memory string return in <1 ms. Key scheme: `{CACHE_NAMESPACE}:{tool_name}:{sha1(canonical_json)[:16]}`. Canonical args via `inspect.signature().bind() + apply_defaults()` so `list_files()` and `list_files(SHAREPOINT_BASE_FOLDER)` hit the same key. TTLs per tool: `read_pdf`/`read_docx`/`read_pptx` = 86400 s (deterministic extraction), `get_file_metadata` = 30 s (embeds `TimeLastModified`), listing/search = 300–600 s. Error strings (`[tool_name]` prefix or `[binary content`) and results >1 MiB are never stored. Graceful degradation: any `RedisError` flips `_client_healthy=False` → decorator becomes a pass-through; periodic probe re-enables it. New `cache_stats()` MCP tool returns `{hits, misses, bypasses, bytes_served}` from in-process counters. Configured via `REDIS_URL` (unset → disabled), `CACHE_ENABLED`, `CACHE_DEFAULT_TTL`, `CACHE_MAX_BYTES`, `CACHE_NAMESPACE`.
**Interview angles:** cache key normalization via `inspect.signature`, TTL-per-tool policy, graceful degradation without try/except leaking to callers, error-skip rule, size cap to prevent OOM, why Redis over `lru_cache` (survives restart, horizontal scale).

### Streaming cancellation end-to-end (#8)
**Problem:** Session switch mid-stream left the upstream Anthropic generation running — wasted tokens = wasted $.
**Fix (client — `app/components/ChatBox.tsx`):** `AbortController` stored in `abortRef`, aborted in the unmount `useEffect`. `signal` passed to `/api/chat` fetch.
**Fix (server — `app/api/chat/route.ts`):** `req.signal` (the incoming `Request`'s abort signal) is bridged to an inner `AbortController` whose `signal` is passed to the upstream Anthropic `fetch`. A second listener on `req.signal` calls `reader.cancel()` so the SSE parse loop unblocks instead of hanging on `reader.read()`. Expected `AbortError`s are swallowed to keep logs clean.
**Interview angle:** cost awareness ($$ per token), connection lifecycle, React 19 cleanup semantics.

### Backpressure + retry on SharePoint (#9)
**Problem:** Claude can fan out N parallel `read_pdf` / `search_*` calls. SharePoint responds with 429 → cascading failures.
**Fix (`tools.py`):** `_with_backpressure` decorator wraps every `@mcp.tool()`:
- `threading.Semaphore(_SP_CONCURRENCY)` caps in-flight SharePoint requests (default 4, override via `SHAREPOINT_MAX_CONCURRENCY`).
- On transient errors (`429`, `503`, `Too Many Requests`, `Service Unavailable` in exception text) retries up to 4 times with `base * 2**attempt + jitter` backoff.
- Non-transient errors propagate immediately.
Stacked as `@mcp.tool()` (outer) → `@_with_backpressure` (inner) so `functools.wraps` keeps the tool's signature visible to FastMCP's schema generator.
**Interview angle:** semaphore vs. rate-limit-token-bucket, thundering herd, retry storms, why jitter matters.
