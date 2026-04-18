# Running the SharePoint MCP Server

All tools live in `tools.py`. Two entry points expose them over different transports.

## Prerequisites

```bash
pip install -r requirements.txt
```

Ensure `.env` exists with:
```
SHAREPOINT_EMAIL=your@email.com
SHAREPOINT_PASSWORD=yourpassword
```

---

## Claude Code (stdio)

**Entry point:** `server_stdio.py`  
**Transport:** stdio — Claude Code spawns the process and communicates over stdin/stdout. No port is opened.

**Config:** `.mcp.json` already points to this file. Claude Code picks it up automatically.

To run manually (for testing):
```bash
python server_stdio.py
```

---

## Next.js Frontend (HTTP)

**Entry point:** `server_http.py`  
**Transport:** streamable-http — Anthropic's API calls back to the MCP server, so it must be publicly reachable.

### Option A — Docker Compose (recommended)

Starts Redis + the MCP server in one command.

**Terminal 1:**
```bash
docker compose up --build
```

**Terminal 2 — ngrok tunnel:**
```bash
ngrok http 8000
```

After code changes, rebuild the MCP container only:
```bash
docker compose up --build mcp
```

Stop everything:
```bash
docker compose down
```

> `REDIS_URL` is set automatically inside the Compose network (`redis://redis:6379/0`).  
> Your `.env` only needs `SHAREPOINT_EMAIL` and `SHAREPOINT_PASSWORD` — do not add `REDIS_URL` to `.env` when using Compose.

---

### Option B — Manual (3 terminals)

**Terminal 1 — Redis:**
```bash
docker run -d -p 6379:6379 --name sp-redis redis:7-alpine \
  --maxmemory 256mb --maxmemory-policy allkeys-lru
```
On subsequent runs: `docker start sp-redis`

Add to `.env`:
```
REDIS_URL=redis://localhost:6379/0
```

**Terminal 2 — MCP server:**
```bash
python server_http.py
```

**Terminal 3 — ngrok tunnel:**
```bash
ngrok http 8000
```

---

### Configure Next.js

Copy the `https://....ngrok-free.app` URL from ngrok output and set in `../sharepoint-next/.env.local`:
```
SHAREPOINT_MCP_URL=https://<your-ngrok-subdomain>.ngrok-free.app/mcp
```

Restart the Next.js dev server. The ngrok URL changes every time ngrok restarts — update `.env.local` each time.

---

## Redis Cache

The caching layer is opt-in. Without `REDIS_URL` set, all tools work normally — no cache.

| Env var | Default | Purpose |
|---|---|---|
| `REDIS_URL` | (unset = disabled) | e.g. `redis://localhost:6379/0` |
| `CACHE_ENABLED` | `true` | Runtime kill-switch |
| `CACHE_DEFAULT_TTL` | `300` | Fallback TTL (seconds) for unlisted tools |
| `CACHE_MAX_BYTES` | `1048576` | Max entry size (1 MiB); oversized results returned but not stored |
| `CACHE_NAMESPACE` | `sp:v1` | Key prefix — bump to `sp:v2` to bust all cached entries |

Inspect cache at runtime via the `cache_stats` MCP tool — returns `{hits, misses, bypasses, bytes_served}`.

---

## File Structure

```
tools.py              — shared MCP instance + all tool definitions
cache.py              — Redis client, _with_cache decorator, TTL map, stats
server_stdio.py       — entry point for Claude Code
server_http.py        — entry point for Next.js frontend
Dockerfile            — builds the MCP server image
docker-compose.yml    — starts Redis + MCP server together
.mcp.json             — Claude Code MCP config (points to server_stdio.py)
```
