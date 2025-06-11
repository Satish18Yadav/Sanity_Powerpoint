import os,pytest
import time
from dotenv import load_dotenv

from patchright.sync_api import sync_playwright,expect,Playwright,Page
load_dotenv()

@pytest.fixture(scope='session')
def open_powerpoint():
    with sync_playwright() as p:
        yield p

@pytest.fixture(scope="session")
def create_browser(open_powerpoint : Playwright):
    browser = open_powerpoint.chromium.launch(headless=False)
    context = browser.new_context(viewport={'width':1360, 'height':770})
    page = context.new_page()
    yield page

'''
This method is also not working

def handle_alert(dialog):
    print(f"Dialog message: {dialog.message}")
    dialog.accept()
'''


@pytest.mark.sanity
def test_create_presentation(create_browser : Page):
    page = create_browser
    page.goto("https://www.powerpoint.com")

    sign_in = page.get_by_role("textbox", name="Enter your email, phone, or")
    sign_in.wait_for(state='visible',timeout=3000)
    sign_in.fill(os.getenv("TEST_EMAIL")) #type: ignore

    Next_button = page.get_by_role("button", name="Next")
    Next_button.wait_for(state="visible")
    expect(Next_button).to_be_visible()
    Next_button.click()

    password_section = page.get_by_role("textbox", name="Password")
    password_section.wait_for(state='visible')
    expect(password_section).to_be_visible()
    password_section.fill(os.getenv("TEST_PASSWORD")) # type: ignore

    page.wait_for_timeout(2000)

    page.get_by_test_id("primaryButton").click()

    page.get_by_test_id("secondaryButton").click()

    blank_presentation = page.get_by_role("link", name="Create a new blank")
    blank_presentation.wait_for(state='visible',timeout=30000)
    expect(blank_presentation).to_be_visible()

    # capturing the new tab 
   # Capture the new tab and wait for navigation
    with page.context.expect_page() as new_page_info:
        blank_presentation.click()
        new_page = new_page_info.value

    # Bring the new tab to the foreground
    new_page.bring_to_front()

    # Wait for the new tab to navigate to the expected URL
    print("Waiting for new tab to navigate to the correct URL...")
    max_wait_time = 60  # seconds
    start_time = time.time()
    expected_url_pattern = "**/onedrive.live.com/**"  # Adjust based on the URL pattern
    while time.time() - start_time < max_wait_time:
        current_url = new_page.url
        print(f"Current URL: {current_url}")
        if "onedrive.live.com" in current_url or "powerpoint.office.com" in current_url:
            print("Expected URL reached!")
            break
        new_page.wait_for_timeout(1000)  # Wait 1 second before checking again
    else:
        raise TimeoutError(f"New tab did not navigate to the expected URL within {max_wait_time} seconds. Final URL: {new_page.url}")

    # Wait for the page to fully load
    new_page.wait_for_load_state("networkidle", timeout=60000)
    new_page.wait_for_load_state("domcontentloaded", timeout=60000)

    # Close all other tabs except the new one and the original page
    for p in page.context.pages:
        if p != new_page and p != page:
            print(f"Closing extra tab: {p.url}")
            p.close()

    # Log the new tab's URL and context
    print(f"New page URL after navigation: {new_page.url}")
    print(f"Number of pages in context after navigation: {len(page.context.pages)}")
    print(f"All page URLs in context: {[p.url for p in page.context.pages]}")
    print(f"Active page in context: {page.context.pages[-1].url} (last page in context)")

    # Take a screenshot for visual debugging
    new_page.screenshot(path="new_tab_screenshot.png")
    print("Screenshot of new tab saved as 'new_tab_screenshot.png'")

    # Attempt to locate and click the "Accept" button with retries and iframe handling
    max_retries = 3
    for attempt in range(max_retries):
        try:
            # First, try the main page
            accept_button = new_page.get_by_role("button", name="Accept")
            if accept_button.count() > 0:
                accept_button.wait_for(state="visible", timeout=30000)
                print(f"Attempt {attempt + 1}: Accept button found on main page, count: {accept_button.count()}")
                accept_button.click()
                print("Clicked Accept button on main page.")
                break
            else:
                # Check for iframes
                iframe = new_page.frame_locator("iframe")
                accept_button = iframe.get_by_role("button", name="Accept")
                if accept_button.count() > 0:
                    accept_button.wait_for(state="visible", timeout=30000)
                    print(f"Attempt {attempt + 1}: Accept button found inside iframe, count: {accept_button.count()}")
                    accept_button.click()
                    print("Clicked Accept button inside iframe.")
                    break
                else:
                    print(f"Attempt {attempt + 1}: Accept button not found in main page or iframe.")
                    html = new_page.content()
                    print(f"HTML content when Accept button not found:\n{html}")
                    if attempt < max_retries - 1:
                        print("Retrying after a short wait...")
                        new_page.wait_for_timeout(5000)  # Wait before retrying
                    else:
                        raise Exception("Accept button not found after all retries.")
        except Exception as e:
            print(f"Attempt {attempt + 1}: Error interacting with Accept button: {str(e)}")
            if attempt == max_retries - 1:
                html = new_page.content()
                print(f"HTML content after final error:\n{html}")
                raise e

    # Pause for manual inspection
    print("Pausing to allow manual inspection of the new tab...")
    new_page.pause()