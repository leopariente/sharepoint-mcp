import os
import io
import re
import sys
from pathlib import Path
from dotenv import load_dotenv
from mcp.server.fastmcp import FastMCP
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.search.service import SearchService

load_dotenv(dotenv_path=Path(__file__).parent / ".env")

SHAREPOINT_SITE_URL = "https://postidcac.sharepoint.com/sites/ComputerScienceLibrary-StudentsTeam2"
SHAREPOINT_BASE_FOLDER = "/sites/ComputerScienceLibrary-StudentsTeam2/Shared Documents"
STUDENT_EMAIL = os.getenv("SHAREPOINT_EMAIL")
STUDENT_PASSWORD = os.getenv("SHAREPOINT_PASSWORD")

mcp = FastMCP("sharepoint-agent")

_CLAUDE_MD = (Path(__file__).parent / "claude.md").read_text(encoding="utf-8")


@mcp.prompt()
def sharepoint_instructions() -> str:
    """Course registry and usage instructions for the SharePoint agent."""
    return _CLAUDE_MD

# Reusable SharePoint context — created once per server lifetime.
_ctx: ClientContext | None = None


def _build_ctx() -> ClientContext:
    """Create and authenticate a new SharePoint ClientContext."""
    if not STUDENT_PASSWORD:
        raise RuntimeError(
            "SHAREPOINT_PASSWORD environment variable is not set. "
            "Set it before starting the server."
        )
    credentials = UserCredential(STUDENT_EMAIL, STUDENT_PASSWORD)
    ctx = ClientContext(SHAREPOINT_SITE_URL).with_credentials(credentials)
    ctx.load(ctx.web)
    ctx.execute_query()
    return ctx


def _get_ctx(force_new: bool = False) -> ClientContext:
    """Return a cached SharePoint ClientContext, creating it if needed."""
    global _ctx
    if _ctx is not None and not force_new:
        return _ctx
    _ctx = _build_ctx()
    return _ctx


def _run_with_retry(fn):
    """Run *fn(ctx)*, retrying once with a fresh context on failure."""
    ctx = _get_ctx()
    try:
        return fn(ctx)
    except Exception:
        ctx = _get_ctx(force_new=True)
        return fn(ctx)



# ---------------------------------------------------------------------------
# Tool: list_files
# ---------------------------------------------------------------------------
@mcp.tool()
def list_files(folder_url: str = SHAREPOINT_BASE_FOLDER) -> str:
    """List files and subdirectories in a SharePoint folder.

    Args:
        folder_url: Server-relative URL of the folder to list
                    (e.g. "/sites/MySite/Shared Documents/subfolder").
                    Leave empty to list the root of the default document library.

    Returns:
        A formatted listing of files (with sizes) and subfolders.
    """

    def _do(ctx):
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

    try:
        return _run_with_retry(_do)
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
    def _do(ctx):
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

    try:
        return _run_with_retry(_do)
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
    if not folder_url:
        folder_url = SHAREPOINT_BASE_FOLDER

    def _do(ctx):
        search = SearchService(ctx)
        kql = f'filename:"{query}" path:"{SHAREPOINT_SITE_URL}"'
        result = search.post_query(
            query_text=kql,
            select_properties=["Path", "FileName"],
            row_limit=500,
            trim_duplicates=False,
        )
        ctx.execute_query()

        matches: list[str] = []
        for row in result.value.PrimaryQueryResult.RelevantResults.Table.Rows:
            cells = {c.Key: c.Value for c in row.Cells}
            path = cells.get("Path", "")
            if path:
                matches.append(path)

        if not matches:
            return f"No files matching '{query}' found under {folder_url}."

        header = f"Files matching '{query}' under {folder_url}:\n"
        return header + "\n".join(f"  {m}" for m in matches)

    try:
        return _run_with_retry(_do)
    except Exception as e:
        return _format_error("search_files", folder_url, e)


# ---------------------------------------------------------------------------
# Tool: search_content
# ---------------------------------------------------------------------------
@mcp.tool()
def search_content(
    query: str,
    folder_url: str = "",
    file_types: list[str] = ["pdf"],
    max_results: int = 10,
) -> str:
    """Full-text search inside SharePoint file contents.

    Args:
        query: Text to search for inside file contents (e.g. "dynamic programming recurrence").
        folder_url: Narrow the search to a specific folder (optional).
        file_types: File extensions to include (e.g. ["pdf", "docx"]). Default: ["pdf"].
        max_results: Maximum number of results to return. Default: 10.

    Returns:
        Ranked list of matching files with relevance score and a content snippet.
    """
    def _do(ctx):
        search = SearchService(ctx)

        kql = f'"{query}"'
        if folder_url:
            host = SHAREPOINT_SITE_URL.split("/sites/")[0]
            kql += f' path:"{host}{folder_url}"'
        else:
            kql += f' path:"{SHAREPOINT_SITE_URL}"'
        if file_types:
            type_filter = " OR ".join(f"FileExtension:{t}" for t in file_types)
            kql += f" ({type_filter})"

        result = search.post_query(
            query_text=kql,
            select_properties=["Title", "Path", "HitHighlightedSummary", "Rank", "FileType"],
            row_limit=max_results,
            trim_duplicates=True,
        )
        ctx.execute_query()

        rows = result.value.PrimaryQueryResult.RelevantResults.Table.Rows
        if not rows:
            return f"No files found containing '{query}'."

        lines = [f"Files containing '{query}':\n"]
        for i, row in enumerate(rows, 1):
            cells = {c.Key: c.Value for c in row.Cells}
            title = cells.get("Title") or cells.get("FileName", "Unknown")
            path = cells.get("Path", "")
            rank = cells.get("Rank", "")
            snippet = re.sub(r"<[^>]+>", "", cells.get("HitHighlightedSummary", "")).strip()
            if len(snippet) > 200:
                snippet = snippet[:200] + "..."

            lines.append(f"[{i}] {title}")
            lines.append(f"    Path:  {path}")
            if rank:
                lines.append(f"    Score: {rank}")
            if snippet:
                lines.append(f"    \"{snippet}\"")
            lines.append("")

        return "\n".join(lines)

    try:
        return _run_with_retry(_do)
    except Exception as e:
        return _format_error("search_content", query, e)


# ---------------------------------------------------------------------------
# Tool: read_pdf
# ---------------------------------------------------------------------------
@mcp.tool()
def read_pdf(
    file_url: str,
    pages: str = "all",
    max_chars: int = 8000,
) -> str:
    """Extract text from a PDF stored in SharePoint.

    Handles Hebrew RTL text and mathematical notation properly.

    Args:
        file_url: SharePoint server-relative URL of the PDF file.
        pages: Page range to read — "all", a single page ("3"), or a range ("1-5").
        max_chars: Maximum characters to return. Default: 8000.

    Returns:
        Extracted text with page headers, file info, and detected language.
    """
    def _parse_page_range(pages_str: str, total: int) -> list[int]:
        s = pages_str.strip().lower()
        if s == "all":
            return list(range(total))
        if "-" in s:
            start, end = s.split("-", 1)
            return list(range(int(start) - 1, min(int(end), total)))
        return [int(s) - 1]

    def _detect_language(text: str) -> str:
        has_hebrew = any("\u0590" <= c <= "\u05FF" for c in text)
        has_latin = any("a" <= c.lower() <= "z" for c in text)
        if has_hebrew and has_latin:
            return "Hebrew + English"
        return "Hebrew" if has_hebrew else "English"

    def _do(ctx):
        try:
            import pymupdf
        except ImportError:
            return "[read_pdf] pymupdf is not installed. Run: pip install pymupdf"

        file_obj = ctx.web.get_file_by_server_relative_url(file_url)
        buf = io.BytesIO()
        file_obj.download(buf).execute_query()
        buf.seek(0)
        raw = buf.read()

        doc = pymupdf.open(stream=io.BytesIO(raw), filetype="pdf")
        page_indices = _parse_page_range(pages, len(doc))

        extracted = []
        for i in page_indices:
            if 0 <= i < len(doc):
                text = doc[i].get_text("text", flags=pymupdf.TEXT_PRESERVE_LIGATURES)
                extracted.append(f"--- Page {i + 1} ---\n{text}")

        full_text = "\n".join(extracted)
        lang = _detect_language(full_text)
        size_mb = round(len(raw) / 1024 / 1024, 2)
        filename = file_url.split("/")[-1]

        header = (
            f"File: {filename}\n"
            f"Pages: {len(doc)} | Language: {lang} | Size: {size_mb} MB\n\n"
        )
        return header + full_text[:max_chars]

    try:
        return _run_with_retry(_do)
    except Exception as e:
        return _format_error("read_pdf", file_url, e)


# ---------------------------------------------------------------------------
# Tool: get_file_metadata
# ---------------------------------------------------------------------------
@mcp.tool()
def get_file_metadata(
    file_url: str,
    include_versions: bool = False,
) -> str:
    """Return rich metadata for a SharePoint file.

    Args:
        file_url: SharePoint server-relative URL of the file.
        include_versions: Whether to include version history. Default: False.

    Returns:
        Author, dates, size, grade parsed from filename, and optional version history.
    """
    def _extract_grade(filename: str) -> str | None:
        match = re.search(r"_(\d{2,3})\.\w+$", filename, re.IGNORECASE)
        return match.group(1) if match else None

    def _do(ctx):
        file_obj = ctx.web.get_file_by_server_relative_url(file_url)
        ctx.load(file_obj, ["Name", "Length", "TimeCreated", "TimeLastModified", "UIVersionLabel"])
        ctx.execute_query()

        author = "Unknown"
        try:
            list_item = file_obj.listItemAllFields
            ctx.load(list_item, ["Author"])
            ctx.execute_query()
            author_val = list_item.get_property("Author")
            if isinstance(author_val, dict):
                author = author_val.get("Title", author_val.get("LoginName", "Unknown"))
            elif author_val:
                author = str(author_val)
        except Exception:
            pass

        props = file_obj.properties
        name = props.get("Name", file_url.split("/")[-1])
        size_mb = round(int(props.get("Length", 0)) / 1024 / 1024, 2)
        created = props.get("TimeCreated", "Unknown")
        modified = props.get("TimeLastModified", "Unknown")
        version = props.get("UIVersionLabel", "Unknown")
        grade = _extract_grade(name)

        lines = [
            f"File: {name}",
            "=" * 40,
            f"Path:     {file_url}",
            f"Author:   {author}",
            f"Uploaded: {created}",
            f"Modified: {modified}",
            f"Size:     {size_mb} MB",
            f"Grade:    {grade if grade else 'N/A (not in filename)'}",
            f"Version:  {version}",
        ]

        if include_versions:
            try:
                versions = file_obj.versions
                ctx.load(versions)
                ctx.execute_query()
                lines.append("\nVersion History:")
                for v in versions:
                    vp = v.properties
                    lines.append(f"  v{vp.get('VersionLabel', '?')} — {vp.get('Created', '?')}")
            except Exception as ve:
                lines.append(f"\nVersion history unavailable: {ve}")

        return "\n".join(lines)

    try:
        return _run_with_retry(_do)
    except Exception as e:
        return _format_error("get_file_metadata", file_url, e)


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
# Authenticate eagerly so the first tool call doesn't pay the auth cost.
try:
    _get_ctx()
    print("SharePoint context ready.", file=sys.stderr)
except Exception as e:
    print(f"Warning: eager auth failed ({e}), will retry on first tool call.", file=sys.stderr)

if __name__ == "__main__":
    mcp.run()
