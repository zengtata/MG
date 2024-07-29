import re
import webbrowser
from playwright.sync_api import Playwright, sync_playwright, expect


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://www.ssangyong.hu/autoink/arlista/")

    # Handle cookie consent or other initial interactions
    page.get_by_role("button", name="Elutasítom").click()

    # Wait for the popup
    with page.expect_popup() as page1_info:
        page.locator("div:nth-child(3) > .sc-fxwrCY > .sc-jnOGJG > div > .sc-kAkpmW").first.click()

    # Get the URL from the popup
    page1 = page1_info.value
    print(f"Popup URL: {page1.url}")

    # Open the URL in the default web browser
    webbrowser.open(page1.url)

    # Clean up
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)
