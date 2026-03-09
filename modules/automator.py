import time
import os
import json
import subprocess
import sys
from datetime import datetime
from groq import Groq
from playwright.sync_api import sync_playwright

AUTH_FILE = "auth_state.json"

def maybe_run_auth_setup():
    response = input("Do you want to run the auth setup now? (y/n): ").strip().lower()

    if response not in {"y", "yes"}:
        return

    setup_script = os.path.join(os.path.dirname(__file__), "setup_auth.py")
    print("Launching auth setup...")
    subprocess.run([sys.executable, setup_script], check=False)

def get_ai_answers(client, questions_data, user_data):
    """Feeds the questions and available options to Groq."""
    print("Sending form structure to AI...")
    
    prompt = f"""
    You are an incredibly fast data-extraction assistant.
    I have a list of Microsoft Form questions (with their multiple choice options included).
    I also have raw user data.
    
    User Data:
    {user_data}
    
    Form Questions & Options:
    {json.dumps(questions_data, indent=2)}
    
    Return ONLY a valid, flat JSON object following these EXACT rules:
    1. For the questions, the JSON keys MUST be the exact Question Text.
    2. The values MUST be the extracted answers. For multiple choice/dropdowns, it MUST be the EXACT text of a provided option.
    3. If the user data does NOT contain the answer for a question, set its value to an empty string "".
    4. You MUST include one additional key at the end named "Status". 
       - If you found answers for ALL questions, set "Status": "OK".
       - If you had to leave any answer blank because data was missing, set "Status": "Missing data: [name of the missing field]".
    """

    response = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[{"role": "user", "content": prompt}],
        response_format={"type": "json_object"},
        temperature=0.1
    )
    
    return json.loads(response.choices[0].message.content)

def run_sniper(api_key, form_url, target_time, user_data):
    client = Groq(api_key=api_key)

    with sync_playwright() as p:
        print("Booting headless browser with saved session...")
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(storage_state=AUTH_FILE)
        page = context.new_page()

        # Convert target time to today's datetime object for safe ">=" comparison
        now = datetime.now()
        target_dt = datetime.strptime(target_time, "%H:%M:%S").replace(
            year=now.year, month=now.month, day=now.day
        )

        print(f"Waiting until {target_time} to navigate...")
        while datetime.now() < target_dt:
            time.sleep(0.05) 

        print("Navigating to form...")
        start_time = time.time()
        
        try:
            page.goto(form_url)
            page.wait_for_selector('div[data-automation-id="questionItem"]', timeout=10000)

            # --- 1. SMART SCRAPE ---
            print("Extracting questions and available options...")
            question_elements = page.query_selector_all('div[data-automation-id="questionItem"]')
            
            questions_data = []
            for q in question_elements:
                title_el = q.query_selector('span[data-automation-id="questionTitle"]')
                if not title_el: continue
                
                q_text = title_el.inner_text()
                q_info = {"question": q_text, "type": "Text", "options": []}

                # A. Check for Dropdown (Listbox)
                dropdown_trigger = q.query_selector('div[aria-haspopup="listbox"]')
                if dropdown_trigger:
                    q_info["type"] = "Dropdown"
                    # Force the dropdown open so the options render in the DOM
                    dropdown_trigger.click() 
                    page.wait_for_timeout(150) # Give React a fraction of a second to render
                    
                    # Read the options from the portal that just appeared
                    options = page.query_selector_all('div[role="option"]')
                    q_info["options"] = [opt.inner_text() for opt in options]
                    
                    # Close it by hitting Escape so it doesn't block other elements
                    page.keyboard.press("Escape") 
                    page.wait_for_timeout(100)

                # B. Check for Multiple Choice (Radio) or Checkbox
                elif q.query_selector('input[type="radio"]') or q.query_selector('input[type="checkbox"]'):
                    q_info["type"] = "Multiple Choice / Checkbox"
                    # Microsoft keeps radio/checkbox text grouped inside the question block
                    # inner_text() easily grabs all available choices in a readable format for the AI
                    q_info["options"] = q.inner_text() 
                    
                questions_data.append(q_info)

            # --- 2. AI MAPPING ---
            answers_dict = get_ai_answers(client, questions_data, user_data)
            print(f"AI generated answers:\n{json.dumps(answers_dict, indent=2)}")

            if (answers_dict.get("Status").split()[0] == "Missing"):
                print("DATA MISSING")
                print(f"{answers_dict.get('Status')}")
                new_data = user_data + " " + answers_dict.get("Status").split(": ")[1] + ": " + input("Please provide the missing data and press Enter: ")
                answers_dict = get_ai_answers(client, questions_data, new_data)
                print(f"Updated AI generated answers:\n{json.dumps(answers_dict, indent=2)}")
                

            # --- 3. DYNAMIC INJECTION ---
            print("Injecting answers...")
            for element in question_elements:
                title_el = element.query_selector('span[data-automation-id="questionTitle"]')
                if not title_el: continue
                
                question_text = title_el.inner_text()
                answer = answers_dict.get(question_text, "")
                
                # Skip if the AI returned empty
                if not answer or str(answer).strip() == "": continue

                dropdown_trigger = element.query_selector('div[aria-haspopup="listbox"]')
                
                # A. Inject Dropdown
                if dropdown_trigger:
                    dropdown_trigger.click()
                    page.wait_for_timeout(150)
                    # Use exact=True so it matches the exact string the AI scraped
                    page.get_by_role("option", name=str(answer), exact=True).click()
                
                # B. Inject Multiple Choice / Checkbox
                elif element.query_selector('input[type="radio"]') or element.query_selector('input[type="checkbox"]'):
                    # Use Playwright's get_by_text to click the exact label
                    element.get_by_text(str(answer), exact=True).click()
                
                # C. Inject Text
                else:
                    input_field = element.query_selector('input, textarea')
                    if input_field:
                        input_field.fill(str(answer))

            # --- 4. SUBMIT & VERIFY ---
            print("Submitting form...")
            page.locator('button[data-automation-id="submitButton"]').click()
            
            print("Waiting for confirmation from Microsoft...")
            try:
                # MS Forms usually shows this class on the "Thank You" page
                page.wait_for_selector('div[data-automation-id="thankYouMessage"]', timeout=8000)
                print("Form successfully submitted and VERIFIED!")
                duration = time.time() - start_time
                print(f"Total time: {duration:.2f} seconds.")
            except Exception as e:
                print("Form submitted, but could not verify the success message.")
                page.screenshot(path="verification_timeout.png")
                print("Saved 'verification_timeout.png'. Check this image to see what happened.")

        except Exception as e:
            print(f"\nA fatal error occurred: {e}")
            page.screenshot(path="error_log.png")
            print("Saved 'error_log.png'. Check this image to see where it crashed.")
        
        finally:
            browser.close()

if __name__ == "__main__":
    print("--- Microsoft Forms Sniper ---")

    if not os.path.exists(AUTH_FILE):
        print("auth_state.json not found.")
        maybe_run_auth_setup()

    else:
        print("An existing auth session was found.")
        maybe_run_auth_setup()

    if os.path.exists(AUTH_FILE):
        API_KEY = input("Enter your Groq API Key: ").strip()
        URL = input("Enter the Microsoft Form URL: ").strip()
        TIME = input("Enter the exact time to submit (HH:MM:SS): ").strip()
        DATA = input("Enter the user data (paste text and press Enter): ").strip()
        run_sniper(API_KEY, URL, TIME, DATA)
    else:
        print("No auth session available. Exiting.")