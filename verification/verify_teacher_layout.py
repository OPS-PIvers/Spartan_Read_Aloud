from playwright.sync_api import sync_playwright
import os

def run():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()

        # Load local file
        file_path = os.path.abspath("verification/teacher_mock.html")
        page.goto(f"file://{file_path}")

        # Take screenshot of the table cell
        element = page.locator("td")
        element.screenshot(path="verification/teacher_layout.png")

        browser.close()

if __name__ == "__main__":
    run()
