"""
Opens a browser for NotebookLM login, waits until logged in, then saves cookies.
"""
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

NOTEBOOKLM_HOME = Path.home() / ".notebooklm"
NOTEBOOKLM_HOME.mkdir(parents=True, exist_ok=True)
STORAGE_FILE = NOTEBOOKLM_HOME / "storage_state.json"
BROWSER_PROFILE = NOTEBOOKLM_HOME / "browser_profile"
BROWSER_PROFILE.mkdir(parents=True, exist_ok=True)

CHROMIUM_PATH = r"C:\Users\admin\AppData\Local\ms-playwright\chromium-1208\chrome-win64\chrome.exe"

print("Opening browser...")

with sync_playwright() as p:
    context = p.chromium.launch_persistent_context(
        user_data_dir=str(BROWSER_PROFILE),
        executable_path=CHROMIUM_PATH,
        headless=False,
        args=["--no-sandbox"],
        slow_mo=100,
    )
    page = context.new_page()
    page.goto("https://notebooklm.google.com")

    print("Waiting for you to log in...")
    print("Sign in with Google in the browser window.")

    # Wait until URL is notebooklm.google.com (not accounts.google.com)
    page.wait_for_url(
        lambda url: "notebooklm.google.com" in url and "accounts.google" not in url,
        timeout=300000  # 5 minutes
    )

    print("Logged in! Saving session...")
    context.storage_state(path=str(STORAGE_FILE))
    print(f"Saved to: {STORAGE_FILE}")

    # Keep browser open briefly so user can see it worked
    page.wait_for_timeout(3000)
    context.close()

print("Done! Login saved.")
