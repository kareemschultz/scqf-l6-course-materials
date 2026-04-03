"""
Opens the existing notebooklm browser profile (already authenticated),
navigates to NotebookLM, and saves the storage state.
"""
import json
from pathlib import Path
from playwright.sync_api import sync_playwright

CHROMIUM = r"C:\Users\admin\AppData\Local\ms-playwright\chromium-1208\chrome-win64\chrome.exe"
BROWSER_PROFILE = Path.home() / ".notebooklm/browser_profile"
OUTPUT = Path.home() / ".notebooklm/profiles/default/storage_state.json"
OUTPUT.parent.mkdir(parents=True, exist_ok=True)

with sync_playwright() as p:
    context = p.chromium.launch_persistent_context(
        user_data_dir=str(BROWSER_PROFILE),
        executable_path=CHROMIUM,
        headless=False,
        args=["--disable-blink-features=AutomationControlled"],
        ignore_default_args=["--enable-automation"],
    )
    page = context.new_page()
    page.goto("https://notebooklm.google.com")
    page.wait_for_load_state("networkidle")
    print("Current URL:", page.url)

    context.storage_state(path=str(OUTPUT))
    print("Session saved to:", OUTPUT)
    context.close()
