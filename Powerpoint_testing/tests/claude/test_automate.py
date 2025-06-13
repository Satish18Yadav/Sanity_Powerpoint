import os
import pytest
import time
from dotenv import load_dotenv
from patchright.sync_api import sync_playwright, expect, Playwright, Page

load_dotenv()

@pytest.fixture(scope='session')
def open_powerpoint():
    with sync_playwright() as p:
        yield p

# Consider changing to scope="function" with cleanup to avoid state pollution:
# @pytest.fixture(scope="function")
# def create_browser(open_powerpoint: Playwright):
#     browser = open_powerpoint.chromium.launch(headless=False)
#     context = browser.new_context(viewport={'width': 1360, 'height': 770})
#     page = context.new_page()
#     try:
#         yield page
#     finally:
#         page.close()
#         context.close()
#         browser.close()
@pytest.fixture(scope="session")
def create_browser(open_powerpoint: Playwright):
    browser = open_powerpoint.chromium.launch(headless=False)
    context = browser.new_context(viewport={'width': 1360, 'height': 770})
    page = context.new_page()
    yield page

@pytest.mark.sanity
def test_create_presentation(create_browser: Page):
    page = create_browser
    page.goto("https://www.powerpoint.com")

    sign_in = page.get_by_role("textbox", name="Enter your email, phone, or")
    sign_in.wait_for(state='visible', timeout=3000)
    sign_in.fill(os.getenv("TEST_EMAIL"))

    next_button = page.get_by_role("button", name="Next")
    next_button.wait_for(state="visible")
    expect(next_button).to_be_visible()
    next_button.click()

    password_section = page.get_by_role("textbox", name="Password")
    password_section.wait_for(state='visible')
    expect(password_section).to_be_visible()
    password_section.fill(os.getenv("TEST_PASSWORD"))

    page.wait_for_timeout(2000)

    page.get_by_test_id("primaryButton").click()
    page.get_by_test_id("secondaryButton").click()

    blank_presentation = page.get_by_role("link", name="Create a new blank")
    blank_presentation.wait_for(state='visible', timeout=30000)
    expect(blank_presentation).to_be_visible()

    # Capture new tab
    with page.context.expect_page() as new_page_info:
        blank_presentation.click()
        new_page = new_page_info.value

    new_page.bring_to_front()

    # Wait for expected URL
    print("Waiting for new tab to navigate to the correct URL...")
    max_wait_time = 60
    start_time = time.time()
    while time.time() - start_time < max_wait_time:
        current_url = new_page.url
        print(f"Current URL: {current_url}")
        if "onedrive.live.com" in current_url or "powerpoint.office.com" in current_url:
            print("Expected URL reached!")
            break
        new_page.wait_for_timeout(1000)
    else:
        raise TimeoutError(f"New tab did not navigate to expected URL within {max_wait_time}s. Final URL: {new_page.url}")

    # Wait for page load
    new_page.wait_for_load_state("networkidle", timeout=60000)
    new_page.wait_for_load_state("domcontentloaded", timeout=60000)

    # Log page context
    print(f"New page URL: {new_page.url}")
    print(f"Pages in context: {len(page.context.pages)}")
    print(f"All page URLs: {[p.url for p in page.context.pages]}")
    print(f"Active page: {page.context.pages[-1].url}")

    # Save screenshot
    new_page.screenshot(path="new_tab_screenshot.png")
    print("Screenshot saved as 'new_tab_screenshot.png'")

    # Handle "Accept" button
    max_retries = 3
    for attempt in range(max_retries):
        try:
            accept_button = new_page.get_by_role("button", name="Accept")
            if accept_button.count() > 0:
                accept_button.wait_for(state="visible", timeout=30000)
                print(f"Attempt {attempt + 1}: Accept button found on main page, count: {accept_button.count()}")
                accept_button.click()
                print("Clicked Accept button on main page.")
                break
            else:
                iframe = new_page.frame_locator("iframe")
                accept_button = iframe.get_by_role("button", name="Accept")
                if accept_button.count() > 0:
                    accept_button.wait_for(state="visible", timeout=30000)
                    print(f"Attempt {attempt + 1}: Accept button found in iframe, count: {accept_button.count()}")
                    accept_button.click()
                    print("Clicked Accept button in iframe.")
                    new_page.screenshot(path="accept_screenshot.png")
                    break
                else:
                    print(f"Attempt {attempt + 1}: Accept button not found.")
                    if attempt == max_retries - 1:
                        html = new_page.content()
                        with open("page.html", "w", encoding="utf-8") as f:
                            f.write(html)
                        raise Exception("Accept button not found after retries. HTML saved to page.html")
        except Exception as e:
            print(f"Attempt {attempt + 1}: Error with Accept button: {str(e)}")
            if attempt == max_retries - 1:
                raise e

    # Check for dialogs/overlays
    dialogs = new_page.query_selector_all("[role='dialog']")
    if len(dialogs) > 0:
        print(f"Found {len(dialogs)} dialogs. Attempting to dismiss.")
        for dialog in dialogs:
            close_button = dialog.query_selector("button[aria-label*='close' i], button[name*='close' i]")
            if close_button:
                close_button.click()
                print("Closed a dialog.")
                new_page.screenshot(path="dialog_closed_screenshot.png")

    # Wait for editor to load
    new_page.wait_for_load_state("domcontentloaded", timeout=60000)

    # Save HTML
    html = new_page.content()
    with open("page.html", "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Page HTML saved to page.html")

    print(f"Current URL: {new_page.url}\n")
    if "powerpoint.office.com" in new_page.url or "onedrive.live.com" in new_page.url:
        print("Successfully navigated to PowerPoint page!")
        print(new_page.url)
        new_page.screenshot(path="editor_screenshot.png")

        # Wait for PowerPoint editor to load
        new_page.wait_for_selector("iframe", timeout=120000)

        # Log iframes
        iframes = new_page.query_selector_all("iframe")
        print(f"Found {len(iframes)} iframes")
        for i, frame in enumerate(iframes):
            src = frame.get_attribute("src") or "no src"
            print(f"Iframe {i}: src={src}")

        # Try to click Home tab with retries
        max_attempts = 10
        attempt = 0
        home_tab_clicked = False

        while attempt < max_attempts and not home_tab_clicked:
            attempt += 1
            print(f"Attempt {attempt} to find and click Home tab...")

            # Try main page
            home_button = new_page.get_by_role("tab", name="Home")
            if home_button.count() > 0:
                try:
                    home_button.wait_for(state="visible", timeout=10000)
                    home_button.click()
                    print("Clicked Home tab on main page.")
                    home_tab_clicked = True
                    break
                except Exception as e:
                    print(f"Error clicking Home tab on main page: {e}")

            # Try iframe
            iframe = new_page.frame_locator("iframe")
            home_button = iframe.get_by_role("tab", name="Home")
            if home_button.count() > 0:
                try:
                    home_button.wait_for(state="visible", timeout=10000)
                    home_button.click()
                    print("Clicked Home tab in iframe.")
                    home_tab_clicked = True
                    break
                except Exception as e:
                    print(f"Error clicking Home tab in iframe: {e}")

            # Fallback to ID-based locator in iframe
            home_button = iframe.locator("#Home")
            if home_button.count() > 0:
                try:
                    home_button.wait_for(state="visible", timeout=10000)
                    home_button.click()
                    print("Clicked Home tab using ID locator in iframe.")
                    home_tab_clicked = True
                    break
                except Exception as e:
                    print(f"Error clicking Home tab with ID locator: {e}")

            # Wait before retrying
            print(f"Home tab not found on attempt {attempt}. Retrying in 5 seconds...")
            new_page.wait_for_timeout(5000)

        # Check if Home tab was clicked
        if not home_tab_clicked:
            html = new_page.content()
            with open("error_page.html", "w", encoding="utf-8") as f:
                f.write(html)
            new_page.screenshot(path="error_screenshot.png")
            raise Exception(f"Home tab not found after {max_attempts} attempts. HTML saved to error_page.html, screenshot saved to error_screenshot.png")

    else:
        raise Exception(f"Failed to reach PowerPoint page. Current URL: {new_page.url}")

    new_page.pause()