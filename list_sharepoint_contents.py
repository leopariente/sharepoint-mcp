from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from dotenv import load_dotenv
import sys
import os

# Load environment variables from .env file
load_dotenv()

# --- CONFIGURATION ---
# The specific SharePoint site URL provided by the user
SHAREPOINT_SITE_URL = "https://postidcac.sharepoint.com/sites/ComputerScienceLibrary-StudentsTeam2"
STUDENT_EMAIL = "leo.pariente@post.runi.ac.il"
STUDENT_PASSWORD = os.getenv("SHAREPOINT_PASSWORD") # IMPORTANT: Replace with your actual password or use environment variable
# ---------------------
# ---------------------

def get_sharepoint_context( ):
    """
    Establishes a connection to SharePoint using either direct user credentials
    or interactive login as a fallback.
    """
    print("Attempting to connect to SharePoint...")
    ctx = None
    try:
        # Try direct authentication first
        print("Trying direct authentication with username/password...")
        user_credentials = UserCredential(STUDENT_EMAIL, STUDENT_PASSWORD)
        ctx = ClientContext(SHAREPOINT_SITE_URL).with_credentials(user_credentials)
        ctx.load(ctx.web)
        ctx.execute_query()
        print(f"Successfully connected to SharePoint site: {ctx.web.properties["Title"]}")
        return ctx
    except Exception as e:
        print(f"Direct authentication failed: {e}")
        print("Falling back to interactive login. A browser window will open for you to authenticate.")
        print("Please complete the login process in the browser, including any MFA steps.")
        try:
            # Fallback to interactive login
            ctx = ClientContext(SHAREPOINT_SITE_URL).with_interactive_login()
            ctx.load(ctx.web)
            ctx.execute_query()
            print(f"Successfully connected to SharePoint site via interactive login: {ctx.web.properties["Title"]}")
            return ctx
        except Exception as interactive_e:
            print(f"Interactive login also failed: {interactive_e}")
            print("Could not establish a connection to SharePoint. Please check your credentials and network.")
            return None

def list_folder_contents_recursive(ctx, folder_url, indent=0):
    """
    Recursively lists all files and subfolders within a given SharePoint folder.
    """
    prefix = "  " * indent
    try:
        # 1. Get the folder object
        folder = ctx.web.get_folder_by_server_relative_url(folder_url)
        
        # FIX 1: Pass a LIST ["Files", "Folders"], not a string "Files, Folders"
        ctx.load(folder, ["Files", "Folders"])
        ctx.execute_query()

        # 2. List files in the current folder
        # We can iterate directly because we loaded "Files" above
        for file in folder.files:
            # We access properties directly. If 'length' fails, use file.properties.get('Length')
            print(f"{prefix}- File: {file.name} (Size: {file.length} bytes)")

        # 3. List subfolders and recurse
        # We can iterate directly because we loaded "Folders" above
        for subfolder in folder.folders:
            if subfolder.name != "Forms": # Ignore system folder
                print(f"{prefix}- Folder: {subfolder.name}/")
                
                # FIX 2: Use the camelCase attribute 'serverRelativeUrl'
                # (This fixes the AttributeError you saw earlier)
                list_folder_contents_recursive(ctx, subfolder.serverRelativeUrl, indent + 1)

    except Exception as e:
        print(f"{prefix}Error accessing {folder_url}: {e}")

def get_default_library(ctx):
    """
    Retrieves the default document library (usually 'Documents' or 'Shared Documents')
    which is safe to access for standard users.
    """
    print("\nConnecting to Default Document Library...")
    # Use the built-in helper to find the main documents library
    lib = ctx.web.default_document_library()
    
    # Load the Title and the RootFolder details specifically
    ctx.load(lib, ["Title", "RootFolder"])
    try:
        ctx.execute_query()
        print(f"Successfully found library: {lib.title}")
        return lib
    except Exception as e:
        print(f"Could not retrieve default library: {e}")
        return None

if __name__ == "__main__":
    sharepoint_context = get_sharepoint_context()
    if sharepoint_context:
        print("\nListing contents of the specified SharePoint site...")
        
        default_lib = get_default_library(sharepoint_context)
        
        if default_lib:
            print(f"\n--- Contents of Document Library: {default_lib.title} ---")
            
            # FIX: Use 'serverRelativeUrl' (camelCase) as suggested by the error
            relative_url = default_lib.root_folder.serverRelativeUrl
            
            list_folder_contents_recursive(sharepoint_context, relative_url)
        else:
            print("No accessible document library found.")

    else:
        print("Script terminated due to authentication failure.")