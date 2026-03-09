# Quick Forms Submitter

Quick Forms Submitter is a Windows-first Python automation project for submitting a Microsoft Form at a precise time with a saved Microsoft login session.

It uses:

- Playwright to open and control the form
- a saved Microsoft authentication session so you do not have to log in at the target moment
- the Groq API to map plain user data to the form's questions and answer options

This project is useful when:

- a Microsoft Form opens at an exact time
- speed matters
- you already know the information you want to submit
- you want the script to match your raw data to the form automatically

Important limitations:

- The authentication helper is written for Windows and looks for Google Chrome or Microsoft Edge in standard Windows install paths.
- The project is designed around Microsoft Forms only.
- You are responsible for using this tool legally and according to the form owner's rules.

## What the project does

There are two main scripts in the `modules` folder:

- `modules/setup_auth.py`: opens a real browser window, lets you sign in to Microsoft Forms, then saves your session into `auth_state.json`
- `modules/automator.py`: waits until your chosen time, opens the target form in a headless browser, asks Groq to map your raw data to the form, fills the form, submits it, and tries to verify success

Generated local files:

- `auth_state.json`: saved Microsoft session data used to bypass the login screen
- `app_browser_profile/`: temporary browser profile used during the authentication setup flow
- `error_log.png`: screenshot taken if the automation crashes before submission completes
- `verification_timeout.png`: screenshot taken if the form appears to submit but the success page cannot be confirmed

These files are local runtime artifacts and should not be committed to GitHub.

## Requirements

Minimum practical requirements:

- Windows 10 or Windows 11
- Python 3.11 or newer
- Google Chrome or Microsoft Edge installed locally
- a Groq API key
- a Microsoft account that can access the target Microsoft Form
- internet access during setup and submission

Tested dependency versions in this repository:

- `playwright==1.42.0`
- `groq==0.4.2`

## Quick start

If you want the shortest version:

1. Install Python.
2. Create and activate a virtual environment.
3. Install the requirements.
4. Install Playwright browser binaries.
5. Run `modules/setup_auth.py` and sign in.
6. Run `modules/automator.py` and enter your Groq key, form URL, time, and user data.

Full instructions are below.

## Full installation

### 1. Clone or download the project

Open PowerShell and go to the folder where you want the project:

```powershell
cd C:\Users\User\Documents\GitHubProjects
git clone <your-repository-url>
cd quick_forms_submitter
```

If you downloaded a ZIP file instead, extract it and open PowerShell inside the project folder.

### 2. Create a virtual environment

From the project root:

```powershell
python -m venv .venv
```

If your machine uses the `py` launcher instead of `python`, use:

```powershell
py -3 -m venv .venv
```

### 3. Activate the virtual environment

PowerShell:

```powershell
.\.venv\Scripts\Activate.ps1
```

Command Prompt:

```bat
.venv\Scripts\activate.bat
```

If PowerShell blocks script execution, run this once in the current PowerShell window:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

Then activate the environment again.

### 4. Install Python dependencies

```powershell
pip install -r requirements.txt
```

### 5. Install Playwright browser binaries

This step is required. Without it, the automation script will not be able to launch the headless browser used for submission.

```powershell
python -m playwright install
```

If you only want Chromium for this project:

```powershell
python -m playwright install chromium
```

### 6. Verify your environment

The following commands should work without errors:

```powershell
python --version
pip --version
python -m playwright --version
```

## First-time authentication setup

Before the actual form automation can work, you need to save a Microsoft login session.

Run:

```powershell
python .\modules\setup_auth.py
```

What happens next:

1. The script looks for Chrome or Edge in standard Windows install locations.
2. It launches that browser with a dedicated local profile in `app_browser_profile/`.
3. It opens `https://forms.office.com/`.
4. You sign in manually.
5. After you fully reach the Microsoft Forms dashboard, you return to the terminal and press Enter.
6. The script connects to that browser session through Chrome DevTools Protocol on port `9222`.
7. It saves the authenticated storage state to `auth_state.json`.

If setup is successful, you should see a confirmation that cookies were found and saved.

### Important notes about authentication

- `auth_state.json` contains session data and should be treated as sensitive.
- If Microsoft logs you out later, run `modules/setup_auth.py` again.
- Do not manually close the browser before pressing Enter in the terminal, or session extraction may fail.
- If the script says it cannot find a browser, install Chrome or Edge normally, or adjust the script for your custom install path.

## How to run the form automation

After authentication is set up, run:

```powershell
python .\modules\automator.py
```

The script will ask for four things:

### 1. Groq API key

Enter your Groq API key exactly as provided by Groq.

### 2. Microsoft Form URL

Paste the full Microsoft Form link.

### 3. Exact submit time

Enter the target time in this format:

```text
HH:MM:SS
```

Examples:

- `07:59:58`
- `14:30:00`
- `23:59:59`

Technical detail:

- The script converts the provided time into today's date using your local system clock.
- If the target time has already passed for the current day, the script will not wait and will continue immediately.

### 4. User data

Paste the raw information that should be used to answer the form.

Examples of user data:

```text
Name: Ana Horvat
Class: 3.B
Email: ana@example.com
Phone: 091 123 4567
Meal choice: Vegetarian
Arrival method: Bus
```

The data does not need to be in a perfect format, but clearer data usually gives better results.

## What happens during automation

When you run `modules/automator.py`, the script does this:

1. Checks that `auth_state.json` exists.
2. Starts a headless Playwright Chromium browser.
3. Loads your saved Microsoft session into that browser context.
4. Waits until the target time.
5. Opens the Microsoft Form.
6. Scrapes the question titles and tries to identify input types.
7. Collects available option text for dropdowns and choice questions.
8. Sends the form structure and your raw user data to the Groq API.
9. Receives a JSON mapping of question text to answers.
10. Fills the form dynamically.
11. Clicks submit.
12. Waits for a Microsoft Forms thank-you message.

If the AI reports missing data, the script will pause and ask you for the missing field, then try again.

## Supported form elements

Based on the current code, the automation handles:

- text inputs
- textareas
- dropdowns / listboxes
- radio-button style multiple choice
- checkbox-style choice groups

Technical detail:

- For dropdowns and multiple choice answers, the returned answer must exactly match one of the visible option labels.
- The script uses the exact form question title text as the JSON key when mapping answers.

## Typical workflow

For everyday use, the workflow is:

1. Activate the virtual environment.
2. Run `modules/setup_auth.py` only when you need to create or refresh the saved Microsoft session.
3. Run `modules/automator.py` before the target submit time.
4. Enter the API key, form URL, submit time, and user data.
5. Let the script wait and submit.

## Example session

```powershell
.\.venv\Scripts\Activate.ps1
python .\modules\automator.py
```

Example inputs:

```text
Enter your Groq API Key: gsk_...
Enter the Microsoft Form URL: https://forms.office.com/Pages/ResponsePage.aspx?id=...
Enter the exact time to submit (HH:MM:SS): 07:59:59
Enter the user data (paste text and press Enter): Name: Ana Horvat, Class: 3.B, Email: ana@example.com
```

## Troubleshooting

### `auth_state.json not found! Run setup_auth program first.`

Cause:

- You have not run the authentication setup yet, or the file was deleted.

Fix:

- Run `python .\modules\setup_auth.py` first.

### Browser not found during setup

Cause:

- Chrome or Edge is not installed in a standard Windows location.

Fix:

- Install Chrome or Edge normally.
- Or edit `modules/setup_auth.py` and add your browser path to the `paths` list in `get_browser_path()`.

### Playwright fails to launch

Cause:

- Playwright dependencies or browser binaries are missing.

Fix:

- Re-run:

```powershell
pip install -r requirements.txt
python -m playwright install
```

### The script submits but cannot verify success

Cause:

- Microsoft Forms may have changed the success page layout.
- Network delays may have delayed the thank-you message.

Fix:

- Check `verification_timeout.png`.
- Confirm whether the form actually submitted.
- If needed, update the success selector in `modules/automator.py`.

### The script crashes before submission

Cause:

- A selector may no longer match the current form layout.
- The form may contain an unsupported question type.
- The session may no longer be valid.

Fix:

- Check `error_log.png`.
- Re-run `modules/setup_auth.py` to refresh the login session.
- Inspect and update selectors in `modules/automator.py` if Microsoft changed the page.

### PowerShell says script execution is disabled

Fix for the current terminal session:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

Then activate `.venv` again.

## Security and privacy notes

- Your Groq API key is sensitive.
- `auth_state.json` is sensitive because it may contain active session data.
- The raw user data you enter is sent to Groq so the model can map it to the form fields.
- Do not publish local session files, screenshots, or private form data in a public repository.

## Technical notes for developers

Current script behavior:

- `modules/setup_auth.py` uses a real installed Chrome or Edge browser and attaches through CDP at `http://localhost:9222`.
- `modules/automator.py` launches Playwright Chromium in headless mode.
- The Groq model currently used in code is `llama-3.3-70b-versatile`.
- The automation waits in a loop with `time.sleep(0.05)` until the target time is reached.
- The form success check currently waits for `div[data-automation-id="thankYouMessage"]`.
- On fatal automation errors, the script writes `error_log.png`.

Current repository-relevant files:

- `README.md`: project documentation
- `requirements.txt`: Python dependency pins
- `modules/setup_auth.py`: manual auth capture flow
- `modules/automator.py`: timed AI-assisted submission logic

## Recommended git hygiene

This repository now includes a `.gitignore` for common local artifacts. If you already tracked sensitive or generated files in your own clone, remove them from git history and the git index before publishing.

Common examples of files that should stay local only:

- `auth_state.json`
- `app_browser_profile/`
- `.venv/`
- screenshots generated during failures

## License

This repository includes a `LICENSE` file. Review it before publishing or redistributing the project.