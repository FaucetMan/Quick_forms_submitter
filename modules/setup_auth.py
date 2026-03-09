import os
import subprocess
import time
from playwright.sync_api import sync_playwright
import json

AUTH_FILE = "auth_state.json"
# Creates a temporary folder in your app's directory to store the browser data
PROFILE_DIR = os.path.join(os.getcwd(), "app_browser_profile") 

def get_browser_path():
    """Finds the default installation path for Chrome or Edge on Windows."""
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
    
    # Python launches the browser
    # It opens port 9222 silently in the background
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
            
            # Save the golden ticket
            context.storage_state(path=AUTH_FILE)
            print(f"✅ Success! Session saved locally to {AUTH_FILE}.")
            
            # Close the connection and the browser window cleanly
            browser.close()
            
        except Exception as e:
            print(f"Failed to extract session: {e}")
            print("Make sure you didn't manually close the browser before pressing Enter.")

    # Ensure the background process is killed
    browser_process.terminate()
    print("Setup complete. Your app is ready for high-speed automation.")

def verify_authentication():
    print("\nVerifying saved session...")
    
    if not os.path.exists(AUTH_FILE):
        print("Error: auth_state.json was not created.")
        return

    # Let's read the JSON file to ensure cookies actually saved
    try:
        with open(AUTH_FILE, 'r') as f:
            data = json.load(f)
            cookies = data.get("cookies", [])
            
            if len(cookies) > 0:
                print(f"Success! Found {len(cookies)} saved cookies.")
                print("Your automated script can now use this file to bypass the login screen.")
            else:
                print("Warning: The file was created, but no cookies were found. Login might have failed.")
    except Exception as e:
        print(f"Error reading the auth file: {e}")

if __name__ == "__main__":
    setup_app_auth()
    verify_authentication()