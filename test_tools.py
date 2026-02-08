"""
Standalone test script for SharePoint MCP tools.

Usage:
    python test_tools.py                  # run all tests
    python test_tools.py list_files       # test only list_files
    python test_tools.py read_file        # test only read_file_content
    python test_tools.py search_files     # test only search_files
"""

import sys

# Import the tools and the connection helper from server.py
from server import list_files, read_file_content, search_files, _get_ctx


def print_header(title: str) -> None:
    sep = "=" * 60
    print(f"\n{sep}")
    print(f"  {title}")
    print(sep)


def test_list_files() -> None:
    print_header("TEST: list_files (root - default library)")
    result = list_files()
    print(result)

    print_header("TEST: list_files (subfolder)")
    subfolder = "/sites/ComputerScienceLibrary-StudentsTeam2/Shared Documents/Year A"
    result = list_files(folder_url=subfolder)
    print(result)


def test_read_file_content() -> None:
    print_header("TEST: read_file_content")
    # First, find a file to read by listing root
    ctx = _get_ctx()
    from server import SHAREPOINT_BASE_FOLDER
    folder = ctx.web.get_folder_by_server_relative_url(SHAREPOINT_BASE_FOLDER)
    ctx.load(folder, ["Files", "Folders"])
    ctx.execute_query()

    # Try to find a text file in root or first subfolder
    file_url = None
    for f in folder.files:
        if any(f.name.endswith(ext) for ext in (".txt", ".md", ".csv", ".json", ".py")):
            file_url = f.serverRelativeUrl
            break

    if not file_url:
        # Pick the first file available regardless of type
        for f in folder.files:
            file_url = f.serverRelativeUrl
            break

    if not file_url:
        print("  No files found in root to test read_file_content.")
        return

    print(f"  Reading: {file_url}")
    result = read_file_content(file_url=file_url)
    if len(result) > 2000:
        print(result[:2000])
        print(f"\n  ... (truncated, total {len(result)} chars)")
    else:
        print(result)


def test_search_files() -> None:
    print_header("TEST: search_files (search for '.pdf')")
    result = search_files(query=".pdf")
    print(result)


TESTS = {
    "list_files": test_list_files,
    "read_file": test_read_file_content,
    "search_files": test_search_files,
}


def main() -> None:
    # Pick which tests to run from CLI args
    requested = sys.argv[1:] if len(sys.argv) > 1 else list(TESTS.keys())

    for name in requested:
        if name not in TESTS:
            print(f"Unknown test '{name}'. Available: {', '.join(TESTS.keys())}")
            sys.exit(1)

    print("SharePoint Tools — Standalone Test Runner")
    print("Connecting to SharePoint...")
    try:
        ctx = _get_ctx()
        print(f"Connected to: {ctx.web.properties.get('Title', 'SharePoint Site')}")
    except Exception as e:
        print(f"Failed to connect to SharePoint: {e}")
        sys.exit(1)

    print(f"Running: {', '.join(requested)}")

    for name in requested:
        try:
            TESTS[name]()
        except KeyboardInterrupt:
            print("\n  Interrupted.")
            break
        except Exception as e:
            print(f"\n  ERROR in {name}: {e}")

    print("\nDone.")


if __name__ == "__main__":
    main()
