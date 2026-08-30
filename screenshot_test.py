from playwright.sync_api import sync_playwright

url = "https://x.com"

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False)

    page = browser.new_page(
        viewport={"width": 1400, "height": 2000}
    )

    page.goto(url)

    page.wait_for_timeout(5000)

    page.screenshot(path="screenshots/test.png")

    browser.close()

print("Screenshot saved!")