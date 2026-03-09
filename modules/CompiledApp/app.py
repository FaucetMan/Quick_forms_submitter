import time
import os
import json
import subprocess
import sys
from datetime import datetime
from groq import Groq
from playwright.sync_api import sync_playwright

# Tell Playwright to save the browser in a folder next to our .exe 
# instead of deep inside the user's hidden Windows AppData folder.
os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.join(os.getcwd(), "pw-browsers")

def ensure_browser_installed():
    """Checks if the browser is downloaded, and if not, downloads it quietly."""
    browsers_dir = os.environ["PLAYWRIGHT_BROWSERS_PATH"]
    
    # If the folder doesn't exist or is empty, we need to install
    if not os.path.exists(browsers_dir) or len(os.listdir(browsers_dir)) == 0:
        print("First run detected: Downloading the hidden browser engine...")
        print("This might take a minute or two depending on your internet speed.\n")
        
        from playwright.__main__ import main as playwright_main
        
        # We temporarily fake the terminal commands so Playwright thinks 
        # the user typed 'playwright install chromium'
        original_argv = sys.argv
        sys.argv = ['playwright', 'install', 'chromium']
        
        try:
            playwright_main()
        except SystemExit as e:
            # We catch that "Exit" command so our script keeps running
            pass 
        finally:
            # Restore the normal terminal arguments
            sys.argv = original_argv
            
        print("\n✅ Browser engine installed successfully!\n")

AUTH_FILE = "auth_state.json"
PROFILE_DIR = os.path.join(os.getcwd(), "app_browser_profile") 

# --- AUTHENTICATION SETUP FUNCTIONS ---

def get_browser_path():
    paths = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
    ]
    for path in paths:
        if os.path.exists(path):
            return path
    return None

def setup_app_auth():
    browser_path = get_browser_path()
    if not browser_path:
        print("Could not find Chrome or Edge on this computer.")
        return

    print("Launching secure browser environment...")
    browser_process = subprocess.Popen([
        browser_path,
        "--remote-debugging-port=9222",
        f"--user-data-dir={PROFILE_DIR}",
        "https://forms.office.com/"
    ])

    print("\n" + "="*40)
    print("        *** ACTION REQUIRED ***")
    print("1. A browser window has opened.")
    print("2. Log in using your email.")
    print("3. Complete the authentication. SELECT 'STAY SIGNED IN' IF PROMPTED.")
    print("4. Wait until you see the Microsoft Forms dashboard.")
    print("="*40)
    
    input("\nPress ENTER here in the terminal ONLY AFTER you are fully logged in...")

    print("\nExtracting CARNET/Microsoft session...")
    with sync_playwright() as p:
        try:
            browser = p.chromium.connect_over_cdp("http://localhost:9222")
            context = browser.contexts[0]
            context.storage_state(path=AUTH_FILE)
            print(f"Success! Session saved locally to {AUTH_FILE}.")
            browser.close()
        except Exception as e:
            print(f"Failed to extract session: {e}")

    browser_process.terminate()

def verify_authentication():
    if not os.path.exists(AUTH_FILE):
        print("Error: auth_state.json was not created.")
        return False
    try:
        with open(AUTH_FILE, 'r') as f:
            data = json.load(f)
            cookies = data.get("cookies", [])
            if len(cookies) > 0:
                print(f"Success! Found {len(cookies)} saved cookies.")
                return True
            else:
                print("Warning: No cookies found. Login might have failed.")
                return False
    except Exception as e:
        print(f"Error reading the auth file: {e}")
        return False

def maybe_run_auth_setup():
    response = input("Do you want to run the auth setup now? (y/n): ").strip().lower()
    if response in {"y", "yes"}:
        setup_app_auth()
        verify_authentication()

# --- SNIPER FUNCTIONS ---

def get_ai_answers(client, questions_data, user_data):
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

            print("Extracting questions and available options...")
            question_elements = page.query_selector_all('div[data-automation-id="questionItem"]')
            
            questions_data = []
            for q in question_elements:
                title_el = q.query_selector('span[data-automation-id="questionTitle"]')
                if not title_el: continue
                
                q_text = title_el.inner_text()
                q_info = {"question": q_text, "type": "Text", "options": []}

                dropdown_trigger = q.query_selector('div[aria-haspopup="listbox"]')
                if dropdown_trigger:
                    q_info["type"] = "Dropdown"
                    dropdown_trigger.click() 
                    page.wait_for_timeout(150) 
                    options = page.query_selector_all('div[role="option"]')
                    q_info["options"] = [opt.inner_text() for opt in options]
                    page.keyboard.press("Escape") 
                    page.wait_for_timeout(100)
                elif q.query_selector('input[type="radio"]') or q.query_selector('input[type="checkbox"]'):
                    q_info["type"] = "Multiple Choice / Checkbox"
                    q_info["options"] = q.inner_text() 
                    
                questions_data.append(q_info)

            answers_dict = get_ai_answers(client, questions_data, user_data)
            print(f"AI generated answers:\n{json.dumps(answers_dict, indent=2)}")

            if (answers_dict.get("Status", "").split()[0] == "Missing"):
                print("DATA MISSING")
                print(f"{answers_dict.get('Status')}")
                missing_field = answers_dict.get("Status").split(": ")[1]
                new_data = user_data + " " + missing_field + ": " + input("Please provide the missing data and press Enter: ")
                browser.close()
                run_sniper(api_key, form_url, target_time, new_data)
                return

            print("Injecting answers...")
            for element in question_elements:
                title_el = element.query_selector('span[data-automation-id="questionTitle"]')
                if not title_el: continue
                
                question_text = title_el.inner_text()
                answer = answers_dict.get(question_text, "")
                
                if not answer or str(answer).strip() == "": continue

                dropdown_trigger = element.query_selector('div[aria-haspopup="listbox"]')
                
                if dropdown_trigger:
                    dropdown_trigger.click()
                    page.wait_for_timeout(150)
                    page.get_by_role("option", name=str(answer), exact=True).click()
                elif element.query_selector('input[type="radio"]') or element.query_selector('input[type="checkbox"]'):
                    element.get_by_text(str(answer), exact=True).click()
                else:
                    input_field = element.query_selector('input, textarea')
                    if input_field:
                        input_field.fill(str(answer))

            print("Submitting form...")
            page.locator('button[data-automation-id="submitButton"]').click()
            
            print("Waiting for confirmation from Microsoft...")
            try:
                page.wait_for_selector('div[data-automation-id="thankYouMessage"]', timeout=8000)
                print("Form successfully submitted and VERIFIED!")
                duration = time.time() - start_time
                print(f"Total time: {duration:.2f} seconds.")
            except Exception as e:
                print("Form submitted, but could not verify the success message.")
                page.screenshot(path="verification_timeout.png")
                
        except Exception as e:
            print(f"\nA fatal error occurred: {e}")
            page.screenshot(path="error_log.png")
        
        finally:
            if not browser.contexts[0].pages[0].is_closed():
                browser.close()

if __name__ == "__main__":
    print("--- Microsoft Forms Sniper ---")

    ensure_browser_installed()

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