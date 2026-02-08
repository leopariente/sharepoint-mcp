import os
import io
from dotenv import load_dotenv
from mcp.server.fastmcp import FastMCP
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext

load_dotenv()

SHAREPOINT_SITE_URL = "https://postidcac.sharepoint.com/sites/ComputerScienceLibrary-StudentsTeam2"
SHAREPOINT_BASE_FOLDER = "/sites/ComputerScienceLibrary-StudentsTeam2/Shared Documents"
STUDENT_EMAIL = os.getenv("SHAREPOINT_EMAIL")
STUDENT_PASSWORD = os.getenv("SHAREPOINT_PASSWORD")

mcp = FastMCP("sharepoint-agent")

# Reusable SharePoint context — created once per server lifetime.
_ctx: ClientContext | None = None


def _get_ctx() -> ClientContext:
    """Return a cached SharePoint ClientContext, creating it on first call."""
    global _ctx
    if _ctx is not None:
        return _ctx

    if not STUDENT_PASSWORD:
        raise RuntimeError(
            "SHAREPOINT_PASSWORD environment variable is not set. "
            "Set it before starting the server."
        )

    credentials = UserCredential(STUDENT_EMAIL, STUDENT_PASSWORD)
    ctx = ClientContext(SHAREPOINT_SITE_URL).with_credentials(credentials)
    ctx.load(ctx.web)
    ctx.execute_query()
    _ctx = ctx
    return _ctx



# ---------------------------------------------------------------------------
# Tool: list_files
# ---------------------------------------------------------------------------
@mcp.tool()
def list_files(folder_url: str = "") -> str:
    """List files and subdirectories in a SharePoint folder.

    Args:
        folder_url: Server-relative URL of the folder to list
                    (e.g. "/sites/MySite/Shared Documents/subfolder").
                    Leave empty to list the root of the default document library.

    Returns:
        A formatted listing of files (with sizes) and subfolders.
    """
    try:
        ctx = _get_ctx()

        if not folder_url:
            folder_url = SHAREPOINT_BASE_FOLDER

        folder = ctx.web.get_folder_by_server_relative_url(folder_url)
        ctx.load(folder, ["Files", "Folders"])
        ctx.execute_query()

        lines: list[str] = [f"Contents of {folder_url}:\n"]

        for f in folder.files:
            lines.append(f"  [FILE] {f.name}  ({f.length} bytes)")

        for sub in folder.folders:
            if sub.name == "Forms":
                continue
            lines.append(f"  [DIR]  {sub.name}/  -> {sub.serverRelativeUrl}")

        if len(lines) == 1:
            lines.append("  (empty)")

        return "\n".join(lines)

    except Exception as e:
        return _format_error("list_files", folder_url, e)


# ---------------------------------------------------------------------------
# Tool: read_file_content
# ---------------------------------------------------------------------------
@mcp.tool()
def read_file_content(file_url: str) -> str:
    """Read and return the text content of a file stored in SharePoint.

    Suitable for text-based files (.txt, .md, .csv, .json, .xml, .html, .py, etc.).
    Binary files will return an error message instead.

    Args:
        file_url: Server-relative URL of the file
                  (e.g. "/sites/MySite/Shared Documents/notes.txt").

    Returns:
        The text content of the file, or an error message.
    """
    try:
        ctx = _get_ctx()
        response = ctx.web.get_file_by_server_relative_url(file_url)
        buf = io.BytesIO()
        response.download(buf).execute_query()
        buf.seek(0)
        raw = buf.read()

        try:
            return raw.decode("utf-8")
        except UnicodeDecodeError:
            try:
                return raw.decode("latin-1")
            except Exception:
                return (
                    f"[binary content — {len(raw)} bytes] "
                    "The file does not appear to be text-based."
                )

    except Exception as e:
        return _format_error("read_file_content", file_url, e)


# ---------------------------------------------------------------------------
# Tool: search_files
# ---------------------------------------------------------------------------
@mcp.tool()
def search_files(query: str, folder_url: str = "") -> str:
    """Search for files by name within a SharePoint folder (recursive).

    Args:
        query: Substring to match against file names (case-insensitive).
        folder_url: Server-relative URL of the folder to search in.
                    Leave empty to search the default document library.

    Returns:
        A list of matching file paths, or a message if none were found.
    """
    try:
        ctx = _get_ctx()

        if not folder_url:
            folder_url = SHAREPOINT_BASE_FOLDER

        matches: list[str] = []
        _search_recursive(ctx, folder_url, query.lower(), matches)

        if not matches:
            return f"No files matching '{query}' found under {folder_url}."

        header = f"Files matching '{query}' under {folder_url}:\n"
        return header + "\n".join(f"  {m}" for m in matches)

    except Exception as e:
        return _format_error("search_files", folder_url, e)


def _search_recursive(
    ctx: ClientContext,
    folder_url: str,
    query_lower: str,
    matches: list[str],
) -> None:
    """Walk the folder tree and collect file paths whose names contain *query_lower*."""
    folder = ctx.web.get_folder_by_server_relative_url(folder_url)
    ctx.load(folder, ["Files", "Folders"])
    ctx.execute_query()

    for f in folder.files:
        if query_lower in f.name.lower():
            matches.append(f.serverRelativeUrl)

    for sub in folder.folders:
        if sub.name == "Forms":
            continue
        _search_recursive(ctx, sub.serverRelativeUrl, query_lower, matches)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _format_error(tool: str, target: str, exc: Exception) -> str:
    msg = str(exc)
    if "403" in msg or "Forbidden" in msg:
        return (
            f"[{tool}] 403 Forbidden for '{target}'. "
            "The current user does not have permission to access this resource."
        )
    if "404" in msg or "Not Found" in msg:
        return (
            f"[{tool}] 404 Not Found for '{target}'. "
            "Check that the path exists and is spelled correctly."
        )
    return f"[{tool}] Error accessing '{target}': {msg}"


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    mcp.run()
