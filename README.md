# Club event automation

Python helper for IIIT Delhi student clubs: generate budget PDFs, fill the institute before/after Google Forms (without submitting), create Gmail drafts, copy an Activity Report spreadsheet, and upload files to Drive.

The form is **never submitted** automatically. Emails are created as **drafts** only. You review and finish those steps yourself.

## What is in this repository

Tracked:

- `foobar_event_automation.py` — main script
- `requirements.txt`
- `event.example.yaml` — template for an event config
- `.gitignore`
- this README

**Not in Git** (you create these locally):

| Path | Why |
|------|-----|
| `credentials.json` | Google OAuth client secret |
| `token.json` | Your club account's login token |
| `events/` | Real event YAML (names, phones, budgets) |
| `browser_profile/` | Playwright Chrome profile / Google session |
| `generated/` | Generated PDFs |
| `.env/` | Python virtual environment |

## 1. Clone and Python environment

```powershell
git clone <this-repo-url>
cd admin_script

python -m venv .env
.env\Scripts\activate
pip install -r requirements.txt
python -m playwright install chromium
```

On macOS/Linux use `source .env/bin/activate`.

## 2. Google Cloud (one-time)

Use the **club** Google account (for example `yourclub@sc.iiitd.ac.in`), not a personal Gmail.

1. Open [Google Cloud Console](https://console.cloud.google.com/) and create or select a project.
2. Enable **Gmail API**, **Google Sheets API**, and **Google Drive API**.
3. Configure the **OAuth consent screen**. If the app is in testing, add the club account as a test user.
4. Create **OAuth client ID** → application type **Desktop app**.
5. Download the JSON and save it in the project root as **`credentials.json`**.

Do not commit `credentials.json`.

## 3. Local Google login

```powershell
python foobar_event_automation.py setup
```

Sign in as the **club** account in the browser that opens. That session is stored in `browser_profile/` (also gitignored).

The first API call (Gmail/Drive/Sheets) will also open a consent window. Approve access. That writes `token.json` (gitignored). Later runs reuse it.

If you change OAuth scopes, delete `token.json` and sign in again.

## 4. Forms, template sheet, and Drive folders

You need:

1. **Before-event** Google Form URL  
2. **After-event** Google Form URL  
3. An **Activity Report spreadsheet template** the club account can copy (File → Make a copy of the council template, or use your own copy). Put that spreadsheet URL in YAML as `drive.activity_report_template`.  
4. **Three Drive folders** the club account can write to:
   - `event_photos` — event photos (form + Activity Report)
   - `cleanup_photos` — timestamped cleanup photos
   - `event_files` — generated expected/actual budget PDFs and the copied Activity Report sheet

Share forms, folders, and the template with the club account if they live in someone else's Drive.

## 5. Create an event YAML

```powershell
mkdir events
copy event.example.yaml events\my_event.yaml
```

Edit `events/my_event.yaml`:

- `club.name` / `club.email`
- event details, organisers, budgets
- the three Drive folder URLs
- `drive.activity_report_template`
- before/after form URLs
- Gmail `to` / `cc` / bodies for **before-event** drafts only

`events/` is gitignored so real data is not pushed.

## 6. Before an event

Fill `expected_budget` and event fields, then:

```powershell
python foobar_event_automation.py before events\my_event.yaml
```

This:

- writes `generated/<event_name>/... Expected Breakdown.pdf`
- uploads that PDF to `drive.event_files`
- fills the before form (including organiser PDF on the Instructions upload) and **does not submit**
- creates Gmail drafts for proposal / room booking

Close the browser when you have reviewed the form.

## 7. After an event

Update `actual_budget` and `after_event` (participants, amount used, cleanup names), then:

```powershell
python foobar_event_automation.py after events\my_event.yaml
```

This:

- writes the actual **Budget Breakdown** PDF and uploads it to `drive.event_files`
- copies the Activity Report template into `drive.event_files` and fills the data row
- fills the after form (organisers PDF, budget PDF, Drive links) and **does not submit**

Filled spreadsheet cells use Space Grotesk, size 9, not bold/italic, black text.

## Commands

| Command | What it does |
|---------|----------------|
| `python foobar_event_automation.py setup` | Save a club Google login in `browser_profile/` |
| `python foobar_event_automation.py before events\....yaml` | Expected PDF + Drive upload + before form + Gmail drafts |
| `python foobar_event_automation.py after events\....yaml` | Actual budget PDF + Activity Report sheet + after form |

## Security

Never commit:

- `credentials.json`
- `token.json`
- `events/`
- `browser_profile/`
- `generated/`

Never put Google passwords or 2FA codes in the repo.

If a secret is committed by mistake, revoke the OAuth client in Cloud Console, delete `token.json`, and clean Git history.

## Test first

Use a dummy event YAML and a test Drive folder before a real event. Confirm the club account, PDFs, form fill, drafts, and spreadsheet copy look right.
