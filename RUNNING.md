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
**Transport:** streamable-http — Anthropic's API calls back to the MCP server, so it must be publicly reachable. `localhost` is not an option — ngrok is required.

### Startup (2 terminals required)

**Terminal 1 — MCP server:**
```bash
cd sharepoint-agent
python server_http.py
```

**Terminal 2 — ngrok tunnel:**
```bash
ngrok http 8000
```

Copy the `https://....ngrok-free.app` URL from the ngrok output.

### Configure Next.js

Set in `../sharepoint-next/.env.local`:
```
SHAREPOINT_MCP_URL=https://<your-ngrok-subdomain>.ngrok-free.app/mcp
```

Then restart the Next.js dev server. The ngrok URL changes every time you restart ngrok — update `.env.local` each time.

---

## File Structure

```
tools.py         — shared MCP instance + all tool definitions
server_stdio.py  — entry point for Claude Code
server_http.py   — entry point for Next.js frontend
.mcp.json        — Claude Code MCP config (points to server_stdio.py)
```
