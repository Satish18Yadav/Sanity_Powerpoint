﻿# Sanity_Tests For PowerPoint

## Overview
- This repo contains the code for the testing performed on Powerpoint.


## Location of the test file:
- Powerpoint_testing -> tests -> async_way -> test_async_way.py

## Tasks performed:
- Login using the .env file, which contains the username and password.
- The clicking whetherexpect **Stay Signd in** or **not**
- Performing the clicking operation on the **Blank Presentation** section for new presentation
- New Tab opens and we need to switch the control on the new tab using the **with** keyword and **expect_page()**
- Then PowerPoint presentation opens up and Acushield gives a warning for the AI unsafe page. Here we need to click on the **Accept Button**
- The above step is the most hardest part.
- After Clicking here we will be navigated to the PowerPoint presentation back

## Challenges faced:
After reaching the Powerpoint page again here I was uanble to interact with any element of the page.

## Important part of the code which has the issue:

+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


    # Pause for manual inspection
    print("Pausing to allow manual inspection of the new tab...")
    # new_page.screenshot(path="new_tab_screenshot.png")


    new_page.wait_for_load_state("domcontentloaded", timeout=60000)

    # Get all frames
    for i, frame in enumerate(page.frames):
        print(f"\nFrame {i}:")
        print("URL:", frame.url)
        try:
            # Try to get some known text or element inside
            text = frame.text_content("body")
            print("Text snippet:", text[:100] if text else "No visible text")
        except Exception as e:
            print("Error reading frame:", e)


    # getting all the HTML of the page
    html = new_page.content()

    print(f"Current URL after iframe interaction: {new_page.url} \n")

    
    if "powerpoint.office.com" in new_page.url or "onedrive.live.com" in new_page.url:
        print("Successfully navigated to PowerPoint page!")
        # new_page.bring_to_front()
        print(new_page.url)
        new_page.screenshot(path="new_tab_screenshot.png")

            
        for i, frame in enumerate(page.frames):
            print(f"\nFrame {i}:")
            print("URL:", frame.url)
            try:
                # Try to get some known text or element inside
                text = frame.text_content("body")
                print("Text snippet: inside the last if block", text if text else "No visible text")
            except Exception as e:
                print("Error reading frame:", e)

        home_button = new_page.get_by_role("tab", name="Home")

        print("check something something here")
       # Assuming new_page and home_button are defined in the broader test context
        max_attempts = 10
        for attempt in range(max_attempts):
            print(f"\nAttempt {attempt + 1}/{max_attempts}: Checking for Home tab...")
            try:
                if home_button.count() > 0:
                    home_button.wait_for(state="visible", timeout=5000)  # Wait up to 5 seconds
                    expect(home_button).to_be_visible()
                    print(f"Home button found with count: {home_button.count()}")
                    home_button.click()
                    print("Clicked the Home tab.")
                    new_page.screenshot(path=f"home_tab_clicked_attempt_{attempt + 1}.png")
                    print(f"Screenshot saved as 'home_tab_clicked_attempt_{attempt + 1}.png'")
                    break  # Exit loop on success
                else:
                    print("Home tab not found on main page.")
                    new_page.screenshot(path=f"home_tab_not_found_attempt_{attempt + 1}.png")
                    print(f"Screenshot saved as 'home_tab_not_found_attempt_{attempt + 1}.png'")
                    if attempt < max_attempts - 1:
                        print("Waiting 2 seconds before retrying...")
                        new_page.wait_for_timeout(2000)  # Wait before next attempt
                    else:
                        html = new_page.content()
                        with open("page.html", "w", encoding="utf-8") as f:
                            f.write(html)
                        print("Page HTML saved as 'page.html'")
                        new_page.pause()
                        raise Exception("Home tab button not found on the page after 10 attempts.")
            except Exception as e:
                print(f"Attempt {attempt + 1} failed: {str(e)}")
                new_page.screenshot(path=f"error_attempt_{attempt + 1}.png")
                print(f"Screenshot saved as 'error_attempt_{attempt + 1}.png'")
                if attempt == max_attempts - 1:
                    html = new_page.content()
                    with open("page.html", "w", encoding="utf-8") as f:
                        f.write(html)
                    print("Page HTML saved as 'page.html'")
                    new_page.pause()
                    raise Exception(f"Home tab button not found after {max_attempts} attempts.")
                new_page.wait_for_timeout(2000)  # Wait before retrying
                
    else:
        raise Exception(f"Failed to reach PowerPoint page. Current URL: {new_page.url}")
    

    
    new_page.pause()

+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
