"""Smoke test: confirms Playwright + Chromium + openpyxl work from this venv."""
from playwright.sync_api import sync_playwright
import openpyxl

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page()
    page.set_content("<h1>ok</h1>")
    print("Chromium", browser.version, "->", page.inner_text("h1"))
    browser.close()
print("openpyxl", openpyxl.__version__)
