from dotenv import load_dotenv
import pytest
import os
import time
from patchright.sync_api import sync_playwright, Page, expect,Playwright 

load_dotenv()

@pytest.fixture(scope="session")
def playwright_instance():
    with sync_playwright() as p:
        yield p 

@pytest.fixture(scope="session")
def open_browser(playwright_instance: Playwright):
        browser = playwright_instance.chromium.launch(headless=False) 
        context = browser.new_context(viewport={"width": 1280, "height": 720})
        page = context.new_page()
        yield page
        # context.close()
        # browser.close()
        

@pytest.mark.sanity
def test_outlook_login_page(open_browser: Page):
    page = open_browser
    page.goto("https://www.outlook.com/")  
    page.wait_for_timeout(5000)
    accept_button = page.get_by_role("button", name="Accept")
    
    if accept_button.count() > 0 and accept_button.is_visible():
        accept_button.click()
        print("Clicked Accept button.")
    else:
        print("Accept button not present or not visible.")

    page.wait_for_timeout(5000)
    sign_in_button = page.locator("#action-oc5b26") 

    with page.context.expect_page() as new_page_info:
     sign_in_button.click()
    
    login_page = new_page_info.value
    login_page.wait_for_load_state()
    print("Navigated to the new tab",login_page.url)

    login_input = login_page.locator("#i0116")
    login_input.wait_for(state='visible',timeout=10000)
    expect(login_input).to_be_visible()

    # Entering the username
    login_input.fill(os.getenv("TEST_EMAIL")) # type: ignore
    login_page.wait_for_timeout(3000)

    #clicking the next button

    login_page.locator("#idSIButton9").click()
    login_page.wait_for_timeout(3000)

    login_password = login_page.locator("#passwordEntry")
    login_password.wait_for(state="visible",timeout=10000)
    expect(login_password).to_be_visible()

    login_password.fill(os.getenv("TEST_PASSWORD")) #type: ignore
    login_page.wait_for_timeout(3000) # type: ignore
    next_button = login_page.locator('[data-testid="primaryButton"]').get_by_text("Next")
    next_button.wait_for(state="visible",timeout=10000)
    expect(next_button).to_be_visible()
    next_button.click()

     # The above part completes the login functionality for Outlook 
  
    login_page.get_by_test_id("secondaryButton").click()
    #clicking on the New Mail buttonn
    login_page.locator("button").filter(has_text="New mailCreate a new email").click()
    email_body = login_page.get_by_role("textbox", name="Message body, press Alt+F10")
    email_body.wait_for(state='visible',timeout=10000)
    expect(email_body).to_be_visible()

    email_body.fill("What is project bluefin?")

    login_page.get_by_role("button", name="Open Copilot").click()

    auto_rewrite = login_page.locator(".Pp26h").first
    auto_rewrite.wait_for(state='visible',timeout=10000)
    auto_rewrite.click()
    login_page.wait_for_timeout(3000)

    # clicking on the proceed button for the Alert pop by Acushield

    pop_acu = login_page.get_by_role("button", name="Proceed")

    if(pop_acu.count() >0 and pop_acu.is_visible(timeout=5000)):
        pop_acu.wait_for(state='visible',timeout=5000)
        expect(pop_acu).to_be_visible()
        pop_acu.click()
    else:
        print("The alert message by the acushield has not come")

    login_page.pause()


