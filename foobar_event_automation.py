#!/usr/bin/env python3
"""
FooBar Event Automation

Usage:
    python run.py before events/valentines_prosort_115.yaml
    python run.py after  events/valentines_prosort_115.yaml

BEFORE:
  - Generates the expected-budget PDF.
  - Opens the BEFORE Google Form in a real Chrome/Chromium profile.
  - Fills the form but DOES NOT submit it.
  - Creates Gmail drafts for the activity/proposal email and room booking email.

AFTER:
  - Generates the actual-budget PDF.
  - Opens the AFTER Google Form.
  - Fills the form but DOES NOT submit it.
  - Appends/updates the event in the Activity Report Google Sheet.
  - Creates an optional Gmail activity-report draft.

Authentication:
  1. Put Google OAuth Desktop credentials at credentials.json.
  2. The first run opens a Google authorization page.
  3. token.json is then reused.
  4. For Google Forms, use a persistent Playwright browser profile so the
     IIITD Google account is already logged in.

This intentionally leaves the final Google Form submission and Gmail sending
to the user for review.
"""

from __future__ import annotations

import argparse
import base64
import json
import mimetypes
import os
import re
import sys
import textwrap
import time
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional

import yaml
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError


# ---------------------------------------------------------------------------
# Paths / constants
# ---------------------------------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent
CREDENTIALS_FILE = BASE_DIR / "credentials.json"
TOKEN_FILE = BASE_DIR / "token.json"
BROWSER_PROFILE = BASE_DIR / "browser_profile"
GENERATED_DIR = BASE_DIR / "generated"

SCOPES = [
    "https://www.googleapis.com/auth/gmail.compose",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

# Exact headers observed in the supplied Activity Report workbook.
ACTIVITY_HEADERS = [
    "Date",
    "Event Title",
    "Description and Details",
    "Type of Event: Open to all/ IIITD only/ Club Members only",
    "No of Participants/ Attendees (Not registrations)",
    "Name of Organisers and Contributers",
    "Event photos link (Optional) A drive folder link would be sufficient",
    "Additional Remarks",
]


# ---------------------------------------------------------------------------
# Generic helpers
# ---------------------------------------------------------------------------

def die(message: str) -> None:
    print(f"\nERROR: {message}\n", file=sys.stderr)
    raise SystemExit(1)


def nested_get(data: Dict[str, Any], *keys: str, default=None):
    cur: Any = data
    for key in keys:
        if not isinstance(cur, dict) or key not in cur:
            return default
        cur = cur[key]
    return cur


def money(value: Any) -> str:
    if value is None:
        return "0"
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return f"{value}"


def require_config(config: Dict[str, Any], *keys: str) -> Any:
    value = nested_get(config, *keys, default=None)
    if value is None or value == "":
        die(f"Missing config value: {' -> '.join(keys)}")
    return value


def load_config(path: Path) -> Dict[str, Any]:
    if not path.exists():
        die(f"Config file does not exist: {path}")

    with path.open("r", encoding="utf-8") as f:
        config = yaml.safe_load(f) or {}

    if not isinstance(config, dict):
        die("YAML config must contain a top-level mapping/object.")

    return config


def event_slug(name: str) -> str:
    return re.sub(r"[^a-zA-Z0-9._-]+", "_", name).strip("_")


def ensure_dirs() -> None:
    GENERATED_DIR.mkdir(parents=True, exist_ok=True)
    (GENERATED_DIR / "budgets").mkdir(parents=True, exist_ok=True)


def parse_date(value: str) -> datetime:
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            pass
    die(f"Could not parse date '{value}'. Use DD-MM-YYYY.")


def display_date(value: str) -> str:
    d = parse_date(value)
    return f"{d.day}/{d.month}/{str(d.year)[-2:]}"


def form_date(value: str) -> str:
    d = parse_date(value)
    return d.strftime("%d-%m-%Y")


def split_time(value: str) -> tuple[str, str, str]:
    """
    Returns hour, minute, AM/PM for values such as:
        05:00 PM
        5:00 PM
        17:00
    """
    value = value.strip().upper()

    for fmt in ("%I:%M %p", "%I %p", "%H:%M"):
        try:
            d = datetime.strptime(value, fmt)
            return d.strftime("%I").lstrip("0") or "12", d.strftime("%M"), d.strftime("%p")
        except ValueError:
            pass

    die(f"Could not parse time '{value}'. Use e.g. '05:00 PM'.")


def organiser_text(config: Dict[str, Any]) -> str:
    lines = []
    for person in config.get("organisers", []):
        lines.append(
            f"{person['name']}, {person['roll_no']}, "
            f"{person.get('cgpa', 'NIL')}, {person.get('phone', '')}, "
            f"{person.get('email', '')}"
        )
    return "\n".join(lines)


def organiser_names(config: Dict[str, Any]) -> str:
    return "\n".join(p["name"] for p in config.get("organisers", []))


def cleanup_text(config: Dict[str, Any]) -> str:
    values = nested_get(config, "after_event", "cleanup_responsibility", default=[])
    if isinstance(values, list):
        return "\n".join(str(x) for x in values)
    return str(values or "")


# ---------------------------------------------------------------------------
# Budget PDF
# ---------------------------------------------------------------------------

def budget_totals(budget: Dict[str, Any]) -> tuple[float, float]:
    inflow = budget.get("inflow", {}) or {}
    outflow = budget.get("outflow", {}) or {}

    total_in = sum(float(v or 0) for v in inflow.values())
    total_out = sum(float(v or 0) for v in outflow.values())
    return total_in, total_out


def make_budget_pdf(
    config: Dict[str, Any],
    which: str,
) -> Path:
    ensure_dirs()

    if which not in ("expected", "actual"):
        raise ValueError(which)

    budget = config.get(f"{which}_budget")
    if not budget:
        die(f"No {which}_budget section found in config.")

    event_name = require_config(config, "event", "name")
    title = (
        f"Expected {event_name} Breakdown"
        if which == "expected"
        else f"{event_name} Budget Breakdown"
    )

    filename = (
        f"{event_slug(event_name)} "
        f"{'Expected Breakdown' if which == 'expected' else 'Budget Breakdown'}.pdf"
    )
    output = GENERATED_DIR / "budgets" / filename

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "BudgetTitle",
        parent=styles["Title"],
        fontSize=16,
        leading=20,
        alignment=TA_CENTER,
        spaceAfter=10,
    )
    heading_style = ParagraphStyle(
        "BudgetHeading",
        parent=styles["Heading2"],
        fontSize=12,
        leading=15,
        spaceBefore=8,
        spaceAfter=5,
    )
    body_style = ParagraphStyle(
        "BudgetBody",
        parent=styles["BodyText"],
        fontSize=10,
        leading=13,
    )

    doc = SimpleDocTemplate(
        str(output),
        pagesize=A4,
        rightMargin=18 * mm,
        leftMargin=18 * mm,
        topMargin=18 * mm,
        bottomMargin=18 * mm,
    )

    story = [
        Paragraph(title, title_style),
        Spacer(1, 3 * mm),
    ]

    inflow_labels = [
        ("institute", "Funds requested from the institute"),
        ("other_income", "Any other sources of income"),
        ("sponsorship", "Sponsorships or donations"),
        ("registration_fees", "Registration fees"),
    ]

    outflow_labels = [
        ("venue_logistics", "Venue and logistics"),
        ("food_hospitality", "Food and hospitality"),
        ("equipment_technical", "Equipment and technical requirements"),
        ("publicity_printing", "Publicity and printing"),
        ("prizes_certificates_merchandise", "Prizes, certificates, merchandise, etc."),
        ("miscellaneous", "Miscellaneous expenses"),
    ]

    def section_table(title_text: str, rows: list[tuple[str, str]], values: Dict[str, Any]):
        data = [["Description", "Amount"]]
        for key, label in rows:
            data.append([label, f"₹{money(values.get(key, 0))}"])

        total = sum(float(values.get(k, 0) or 0) for k, _ in rows)
        data.append(["TOTAL", f"₹{money(total)}"])

        table = Table(data, colWidths=[135 * mm, 35 * mm])
        table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#eeeeee")),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                    ("ALIGN", (1, 1), (1, -1), "RIGHT"),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 6),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                    ("TOPPADDING", (0, 0), (-1, -1), 5),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ]
            )
        )

        story.append(Paragraph(title_text, heading_style))
        story.append(table)
        story.append(Spacer(1, 3 * mm))

    story.append(Paragraph("1. Inflow", heading_style))
    inflow = budget.get("inflow", {}) or {}
    section_table("Sources of Funds", inflow_labels, inflow)

    story.append(Paragraph("2. Outflow", heading_style))
    outflow = budget.get("outflow", {}) or {}
    section_table("Expenditure", outflow_labels, outflow)

    total_in, total_out = budget_totals(budget)

    note = (
        f"<b>Total Inflow:</b> ₹{money(total_in)} &nbsp;&nbsp;&nbsp; "
        f"<b>Total Outflow:</b> ₹{money(total_out)}"
    )
    story.append(Paragraph(note, body_style))

    if total_in < total_out:
        story.append(
            Spacer(1, 2 * mm)
        )
        story.append(
            Paragraph(
                "<b>WARNING:</b> Total inflow is less than total outflow.",
                ParagraphStyle(
                    "Warning",
                    parent=body_style,
                    textColor=colors.red,
                ),
            )
        )

    notes = budget.get("notes", {}) or {}
    if notes:
        story.append(Paragraph("Notes", heading_style))
        for key, value in notes.items():
            story.append(
                Paragraph(
                    f"<b>{key.replace('_', ' ').title()}:</b> {value}",
                    body_style,
                )
            )

    doc.build(story)
    print(f"✓ Generated {output}")
    return output


# ---------------------------------------------------------------------------
# Google OAuth
# ---------------------------------------------------------------------------

def google_credentials() -> Credentials:
    creds = None

    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)

    if creds and creds.valid:
        return creds

    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        if not CREDENTIALS_FILE.exists():
            die(
                f"Missing {CREDENTIALS_FILE}. Download Google OAuth Desktop "
                "credentials and save them as credentials.json."
            )

        flow = InstalledAppFlow.from_client_secrets_file(
            str(CREDENTIALS_FILE),
            SCOPES,
        )
        creds = flow.run_local_server(port=0)

    TOKEN_FILE.write_text(creds.to_json(), encoding="utf-8")
    return creds


# ---------------------------------------------------------------------------
# Gmail drafts
# ---------------------------------------------------------------------------

def gmail_service():
    return build("gmail", "v1", credentials=google_credentials())


def build_email(
    to: Iterable[str],
    cc: Iterable[str],
    subject: str,
    body: str,
    attachments: Optional[List[Path]] = None,
) -> EmailMessage:
    msg = EmailMessage()
    msg["To"] = ", ".join(to)
    if list(cc):
        msg["Cc"] = ", ".join(cc)
    msg["Subject"] = subject
    msg.set_content(body)

    for attachment in attachments or []:
        attachment = Path(attachment)
        if not attachment.exists():
            continue

        mime_type, _ = mimetypes.guess_type(str(attachment))
        if mime_type:
            maintype, subtype = mime_type.split("/", 1)
        else:
            maintype, subtype = "application", "octet-stream"

        msg.add_attachment(
            attachment.read_bytes(),
            maintype=maintype,
            subtype=subtype,
            filename=attachment.name,
        )

    return msg


def create_gmail_draft(
    service,
    to: List[str],
    cc: List[str],
    subject: str,
    body: str,
    attachments: Optional[List[Path]] = None,
) -> str:
    msg = build_email(to, cc, subject, body, attachments)

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    draft = (
        service.users()
        .drafts()
        .create(
            userId="me",
            body={"message": {"raw": raw}},
        )
        .execute()
    )

    return draft["id"]


def render_template(template: str, config: Dict[str, Any]) -> str:
    event = config.get("event", {})
    values = {
        "event.name": event.get("name", ""),
        "event.date": event.get("date", ""),
        "event.start_time": event.get("start_time", ""),
        "event.end_time": event.get("end_time", ""),
        "event.venue": event.get("venue", ""),
        "event.expected_participants": event.get("expected_participants", ""),
        "event.description": event.get("description", "").strip(),
        "event.budget": nested_get(config, "expected_budget", "inflow", "institute", default=0),
    }

    # Support {event.name}-style placeholders.
    for key, value in values.items():
        template = template.replace("{" + key + "}", str(value))

    return template


def create_email_drafts(config: Dict[str, Any], mode: str, budget_pdf: Path) -> None:
    service = gmail_service()
    emails = config.get("emails", {}) or {}
    event = config.get("event", {}) or {}

    if mode == "before":
        activity = emails.get("activity_report", {}) or {}
        room = emails.get("room_booking", {}) or {}

        activity_body = activity.get(
            "body",
            """Greetings,

This is with regards to the event that FooBar is conducting.
Following are the details:

Event Name: {event.name}
Event Date: {event.date}
Event Time: {event.start_time} to {event.end_time}
Proposed Venue: {event.venue}
Expected Number of Participants: {event.expected_participants}
Description of Event: {event.description}
Budget: {event.budget}

Best regards,
Team FooBar""",
        )

        activity_body = render_template(activity_body, config)

        activity_attachments = [
            Path(x) for x in activity.get("attachments", [])
        ]

        draft_id = create_gmail_draft(
            service,
            activity.get("to", []),
            activity.get("cc", []),
            activity.get("subject", event.get("name", "")),
            activity_body,
            activity_attachments,
        )
        print(f"✓ Activity/proposal email draft created: {draft_id}")

        room_body = room.get(
            "body",
            """Good Evening,

This is with regards to the event FooBar is conducting - {event.name}.

We want to book {event.venue} on {event.date} from 4 to 7 pm.
Kindly do the needful.

Best regards,
Team FooBar""",
        )

        room_body = render_template(room_body, config)

        draft_id = create_gmail_draft(
            service,
            room.get("to", []),
            room.get("cc", []),
            room.get("subject", "Room Booking - FooBar"),
            room_body,
            [Path(x) for x in room.get("attachments", [])],
        )
        print(f"✓ Room booking email draft created: {draft_id}")

    else:
        activity = emails.get("after_event_activity_report", {}) or {}
        body = activity.get(
            "body",
            """Greetings,

This is with regards to the event FooBar conducted - {event.name}.

The event was conducted on {event.date} at {event.venue}.

Number of participants attended: {after_event.participants_attended}
Amount used from institute: {after_event.amount_used}

Activity Report:
{activity_report.sheet_url}

Photos:
{after_event.event_photos_link}

Best regards,
Team FooBar""",
        )

        body = body.replace("{event.name}", event.get("name", ""))
        body = body.replace("{event.date}", event.get("date", ""))
        body = body.replace("{event.venue}", event.get("venue", ""))
        body = body.replace(
            "{after_event.participants_attended}",
            str(nested_get(config, "after_event", "participants_attended", default="")),
        )
        body = body.replace(
            "{after_event.amount_used}",
            str(nested_get(config, "after_event", "amount_used", default="")),
        )
        body = body.replace(
            "{activity_report.sheet_url}",
            str(nested_get(config, "activity_report", "sheet_url", default="")),
        )
        body = body.replace(
            "{after_event.event_photos_link}",
            str(nested_get(config, "after_event", "event_photos_link", default="")),
        )

        draft_id = create_gmail_draft(
            service,
            activity.get("to", []),
            activity.get("cc", []),
            activity.get("subject", f"Activity Report - {event.get('name', '')}"),
            body,
            [Path(x) for x in activity.get("attachments", [])],
        )
        print(f"✓ After-event email draft created: {draft_id}")


# ---------------------------------------------------------------------------
# Google Sheets Activity Report
# ---------------------------------------------------------------------------

def sheets_service():
    return build("sheets", "v4", credentials=google_credentials())


def spreadsheet_id_from_url(url: str) -> str:
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", url)
    if not match:
        die(f"Could not extract spreadsheet ID from URL: {url}")
    return match.group(1)


def get_first_sheet_title(service, spreadsheet_id: str) -> str:
    result = (
        service.spreadsheets()
        .get(
            spreadsheetId=spreadsheet_id,
            fields="sheets(properties(sheetId,title))",
        )
        .execute()
    )
    sheets = result.get("sheets", [])
    if not sheets:
        die("The Activity Report spreadsheet contains no sheets.")
    return sheets[0]["properties"]["title"]


def read_activity_headers(service, spreadsheet_id: str, sheet_name: str) -> List[str]:
    result = (
        service.spreadsheets()
        .values()
        .get(
            spreadsheetId=spreadsheet_id,
            range=f"'{sheet_name}'!A1:H5",
        )
        .execute()
    )

    rows = result.get("values", [])
    for row in rows:
        normalized = [str(x).strip() for x in row]
        if "Date" in normalized and "Event Title" in normalized:
            return normalized

    # Fall back to the exact known headers from the supplied workbook.
    return ACTIVITY_HEADERS.copy()


def find_existing_event_row(
    service,
    spreadsheet_id: str,
    sheet_name: str,
    event_name: str,
) -> Optional[int]:
    result = (
        service.spreadsheets()
        .values()
        .get(
            spreadsheetId=spreadsheet_id,
            range=f"'{sheet_name}'!A:H",
        )
        .execute()
    )

    for index, row in enumerate(result.get("values", []), start=1):
        if len(row) >= 2 and str(row[1]).strip() == event_name:
            return index

    return None


def activity_report_values(config: Dict[str, Any]) -> Dict[str, str]:
    event = config["event"]
    after = config.get("after_event", {})
    report = config.get("activity_report", {})

    report_type = report.get("type", "IIITD only")

    return {
        "Date": display_date(event["date"]),
        "Event Title": event["name"],
        "Description and Details": event.get("description", "").strip(),
        "Type of Event: Open to all/ IIITD only/ Club Members only": report_type,
        "No of Participants/ Attendees (Not registrations)": str(
            after.get("participants_attended", event.get("expected_participants", ""))
        ),
        "Name of Organisers and Contributers": organiser_names(config),
        "Event photos link (Optional) A drive folder link would be sufficient": (
            after.get("event_photos_link")
            or report.get("photos_link")
            or ""
        ),
        "Additional Remarks": (
            after.get("remarks")
            or report.get("remarks")
            or "NIL"
        ),
    }


def update_activity_report(config: Dict[str, Any]) -> None:
    url = require_config(config, "activity_report", "sheet_url")
    spreadsheet_id = spreadsheet_id_from_url(url)

    service = sheets_service()
    sheet_name = get_first_sheet_title(service, spreadsheet_id)
    headers = read_activity_headers(service, spreadsheet_id, sheet_name)

    values_by_header = activity_report_values(config)

    # Only write through H, matching the supplied Activity Report workbook.
    row = [
        values_by_header.get(header, "")
        for header in headers[:8]
    ]

    existing_row = find_existing_event_row(
        service,
        spreadsheet_id,
        sheet_name,
        config["event"]["name"],
    )

    if existing_row:
        range_name = f"'{sheet_name}'!A{existing_row}:H{existing_row}"
        (
            service.spreadsheets()
            .values()
            .update(
                spreadsheetId=spreadsheet_id,
                range=range_name,
                valueInputOption="USER_ENTERED",
                body={"values": [row]},
            )
            .execute()
        )
        print(f"✓ Updated existing Activity Report row {existing_row}")
    else:
        (
            service.spreadsheets()
            .values()
            .append(
                spreadsheetId=spreadsheet_id,
                range=f"'{sheet_name}'!A:H",
                valueInputOption="USER_ENTERED",
                insertDataOption="INSERT_ROWS",
                body={"values": [row]},
            )
            .execute()
        )
        print("✓ Added event to Activity Report")


# ---------------------------------------------------------------------------
# Google Forms / Playwright
# ---------------------------------------------------------------------------

def launch_browser():
    """
    Uses a persistent Chromium profile. Log into the IIITD Google account
    once in this profile; subsequent runs reuse the session.
    """
    playwright = sync_playwright().start()
    context = playwright.chromium.launch_persistent_context(
        user_data_dir=str(BROWSER_PROFILE),
        headless=False,
        viewport={"width": 1400, "height": 1000},
    )
    page = context.pages[0] if context.pages else context.new_page()
    return playwright, context, page


def visible_question_container(page, question: str):
    """
    Google Forms changes its internal DOM frequently. This intentionally
    searches visible text first, then walks upward to a likely question
    container.
    """
    locator = page.get_by_text(
        re.compile(rf"^{re.escape(question)}\s*\*?$", re.I)
    ).first

    if locator.count() == 0:
        locator = page.get_by_text(
            re.compile(re.escape(question), re.I)
        ).first

    if locator.count() == 0:
        return None

    # Walk upward several levels until we find a container containing inputs.
    candidate = locator
    for _ in range(7):
        try:
            if candidate.locator("input, textarea, [role=radio], [role=checkbox]").count():
                return candidate
            candidate = candidate.locator("..")
        except Exception:
            break

    return candidate


def fill_text_question(page, question: str, value: str) -> None:
    container = visible_question_container(page, question)
    if container is None:
        die(f"Could not locate Google Form question: {question}")

    inputs = container.locator("input:not([type=radio]):not([type=checkbox]), textarea")
    if inputs.count() == 0:
        die(f"Could not find text input for form question: {question}")

    inputs.first.fill(str(value))


def click_option(page, question: str, option: str, checkbox: bool = False) -> None:
    container = visible_question_container(page, question)
    if container is None:
        die(f"Could not locate Google Form question: {question}")

    option_re = re.compile(rf"^{re.escape(option)}$", re.I)
    label = container.get_by_text(option_re).first

    if label.count() == 0:
        label = page.get_by_text(option_re).first

    if label.count() == 0:
        die(f"Could not locate option '{option}' for '{question}'.")

    label.click()


def fill_date_question(page, question: str, value: str) -> None:
    container = visible_question_container(page, question)
    if container is None:
        die(f"Could not locate date question: {question}")

    inputs = container.locator("input")
    if inputs.count() == 0:
        die(f"Could not find date input for: {question}")

    d = parse_date(value)
    inputs.first.fill(d.strftime("%Y-%m-%d"))


def fill_time_question(page, question: str, value: str) -> None:
    container = visible_question_container(page, question)
    if container is None:
        die(f"Could not locate time question: {question}")

    hour, minute, ampm = split_time(value)
    inputs = container.locator("input")
    if inputs.count() < 2:
        die(f"Could not find time inputs for: {question}")

    inputs.nth(0).fill(hour)
    inputs.nth(1).fill(minute)

    # Google Forms uses a combobox for AM/PM.
    combos = container.locator("[role=combobox]")
    if combos.count():
        combos.first.click()
        page.get_by_text(ampm, exact=True).last.click()


def click_record_email(page) -> None:
    text = re.compile(r"Record .* as the email to be included with my response", re.I)
    label = page.get_by_text(text).first
    if label.count():
        try:
            label.click()
        except Exception:
            pass


def upload_file_question(page, question_text: str, file_path: Path) -> None:
    """
    Finds the file input near the supplied question text.
    """
    if not file_path.exists():
        die(f"Upload file does not exist: {file_path}")

    container = visible_question_container(page, question_text)
    if container is None:
        die(f"Could not locate upload question: {question_text}")

    file_inputs = container.locator("input[type=file]")
    if file_inputs.count() == 0:
        # Sometimes the input is rendered slightly outside the question node.
        file_inputs = page.locator("input[type=file]")

    if file_inputs.count() == 0:
        die(f"Could not find file upload input for: {question_text}")

    file_inputs.last.set_input_files(str(file_path))


def fill_before_form(page, config: Dict[str, Any], budget_pdf: Path) -> None:
    event = config["event"]

    page.goto(config["forms"]["before_event"]["url"], wait_until="domcontentloaded")
    page.wait_for_timeout(2000)

    click_record_email(page)

    click_option(page, "Proposing Body", "Club (Specify Below)")
    fill_text_question(page, 'If "Club", please mention the name of the club:', "FooBar")
    fill_text_question(page, 'If "Other", specify the group\'s name', "NA")

    fill_text_question(page, "Instructions:", organiser_text(config))

    fill_text_question(page, "Event Title", event["name"])
    click_option(page, "Type of Event", event.get("type", "Competition"))
    fill_text_question(page, "Description of the Event", event["description"])

    fill_date_question(page, "Proposed Start Date", event["date"])
    fill_time_question(page, "Proposed Start Time", event["start_time"])
    fill_date_question(page, "End Date", event["date"])
    fill_time_question(page, "End Time", event["end_time"])

    fill_text_question(page, "Proposed Venue", event["venue"])
    fill_text_question(
        page,
        "Expected Number of Participants:",
        event["expected_participants"],
    )

    for audience in event.get("target_audience", ["Institute Students"]):
        click_option(page, "Target Audience. Tick all that are applicable", audience, checkbox=True)

    fill_text_question(
        page,
        "Total Amount Requested from Institute",
        nested_get(config, "expected_budget", "inflow", "institute", default=0),
    )

    upload_file_question(
        page,
        "Budget Breakdown",
        budget_pdf,
    )

    cleanup = cleanup_text(config)
    if cleanup:
        fill_text_question(
            page,
            "Post-Event Clean-Up Responsibility (Mandatory)",
            cleanup,
        )

    print("\n✓ BEFORE form has been filled.")
    print("Review the form in the browser.")
    print("Do NOT close the browser until you are finished reviewing.")
    input("\nPress ENTER after you have manually submitted the form...")


def fill_after_form(page, config: Dict[str, Any], budget_pdf: Path) -> None:
    event = config["event"]
    after = config.get("after_event", {})

    page.goto(config["forms"]["after_event"]["url"], wait_until="domcontentloaded")
    page.wait_for_timeout(2000)

    click_record_email(page)

    click_option(page, "Event Conducted", "Club (Specify Below)")
    fill_text_question(page, 'If "Club", please mention the name of the club:', "FooBar")
    fill_text_question(page, 'If "Other", specify the group\'s name', "NA")

    fill_text_question(page, "Instructions:", organiser_text(config))

    fill_text_question(page, "Event Title", event["name"])
    click_option(page, "Type of Event", event.get("type", "Competition"))
    fill_text_question(page, "Description of the Event", event["description"])

    fill_date_question(page, "Event Start Date", event["date"])
    fill_time_question(page, "Event Start Time", event["start_time"])
    fill_date_question(page, "End Date", event["date"])
    fill_time_question(page, "End Time", event["end_time"])

    fill_text_question(page, "Venue", event["venue"])

    participants = after.get("participants_attended")
    if participants is None:
        participants = input("Number of participants attended: ").strip()
        after["participants_attended"] = participants

    fill_text_question(
        page,
        "Number of Participants attended:",
        participants,
    )

    for audience in event.get("target_audience", ["Institute Students"]):
        click_option(page, "Target Audience. Tick all that are applicable", audience, checkbox=True)

    amount_used = after.get("amount_used")
    if amount_used is None:
        amount_used = input("Total amount used from institute: ").strip()
        after["amount_used"] = amount_used

    fill_text_question(
        page,
        "Total Amount used from Institute",
        amount_used,
    )

    upload_file_question(
        page,
        "Budget Breakdown",
        budget_pdf,
    )

    cleanup = cleanup_text(config)
    if not cleanup:
        cleanup = input("Cleanup responsibility (Name + Roll No.): ").strip()
        after["cleanup_responsibility"] = [cleanup]

    fill_text_question(
        page,
        "Post-Event Clean-Up Responsibility (Mandatory)",
        cleanup,
    )

    timestamped = after.get("timestamped_photos_link")
    if not timestamped:
        timestamped = input("Timestamped photos Drive link: ").strip()
        after["timestamped_photos_link"] = timestamped

    fill_text_question(
        page,
        "Upload the timestamped photograph(s) of the event area",
        timestamped,
    )

    report_link = nested_get(config, "activity_report", "sheet_url", default="")
    fill_text_question(
        page,
        "Link to the Activity Report (Mandatory for clubs)",
        report_link,
    )

    photos = after.get("event_photos_link")
    if not photos:
        photos = input("Photos of Event Drive link: ").strip()
        after["event_photos_link"] = photos

    fill_text_question(
        page,
        "Photos of the Event",
        photos,
    )

    # Persist interactive after-event answers back to the YAML config.
    config["after_event"] = after

    print("\n✓ AFTER form has been filled.")
    print("Review the form in the browser.")
    input("\nPress ENTER after you have manually submitted the form...")


# ---------------------------------------------------------------------------
# Main workflows
# ---------------------------------------------------------------------------

def run_before(config: Dict[str, Any]) -> None:
    ensure_dirs()

    require_config(config, "event", "name")
    require_config(config, "event", "date")
    require_config(config, "event", "start_time")
    require_config(config, "event", "end_time")
    require_config(config, "event", "venue")
    require_config(config, "forms", "before_event", "url")
    require_config(config, "forms", "after_event", "url")

    print("\n========================================")
    print("FOOBAR EVENT AUTOMATION - BEFORE EVENT")
    print("========================================")
    print(f"Event: {config['event']['name']}\n")

    budget_pdf = make_budget_pdf(config, "expected")

    playwright = context = page = None
    try:
        playwright, context, page = launch_browser()
        fill_before_form(page, config, budget_pdf)
    finally:
        if context:
            context.close()
        if playwright:
            playwright.stop()

    create_email_drafts(config, "before", budget_pdf)

    print("\n========================================")
    print("BEFORE EVENT COMPLETE")
    print("========================================")
    print("✓ Expected budget PDF generated")
    print("✓ Google Form reviewed/submitted manually")
    print("✓ Gmail drafts created")
    print("\nRemember to review and send the Gmail drafts manually.")


def run_after(config: Dict[str, Any]) -> None:
    ensure_dirs()

    require_config(config, "event", "name")
    require_config(config, "event", "date")
    require_config(config, "event", "venue")
    require_config(config, "forms", "after_event", "url")
    require_config(config, "activity_report", "sheet_url")

    print("\n========================================")
    print("FOOBAR EVENT AUTOMATION - AFTER EVENT")
    print("========================================")
    print(f"Event: {config['event']['name']}\n")

    budget_pdf = make_budget_pdf(config, "actual")

    playwright = context = page = None
    try:
        playwright, context, page = launch_browser()
        fill_after_form(page, config, budget_pdf)
    finally:
        if context:
            context.close()
        if playwright:
            playwright.stop()

    update_activity_report(config)
    create_email_drafts(config, "after", budget_pdf)

    print("\n========================================")
    print("AFTER EVENT COMPLETE")
    print("========================================")
    print("✓ Actual budget PDF generated")
    print("✓ Google Form reviewed/submitted manually")
    print("✓ Activity Report updated")
    print("✓ Gmail draft created")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="FooBar event automation"
    )
    parser.add_argument(
        "mode",
        choices=["before", "after"],
        help="Run before-event or after-event workflow",
    )
    parser.add_argument(
        "config",
        type=Path,
        help="Path to event YAML config",
    )

    args = parser.parse_args()
    config = load_config(args.config)

    if args.mode == "before":
        run_before(config)
    else:
        run_after(config)


if __name__ == "__main__":
    main()
