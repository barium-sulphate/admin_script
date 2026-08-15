#!/usr/bin/env python3
"""
Club event automation for IIIT Delhi student clubs.

Usage:
    python foobar_event_automation.py setup
    python foobar_event_automation.py before events/your_event.yaml
    python foobar_event_automation.py after  events/your_event.yaml

BEFORE:
  - Generates the expected-budget PDF and uploads it to drive.event_files.
  - Opens the BEFORE Google Form, fills it, does NOT submit.

AFTER:
  - Generates the actual budget-breakdown PDF and uploads it to drive.event_files.
  - Copies the Activity Report template into drive.event_files and fills it.
  - Opens the AFTER Google Form, fills it, does NOT submit.

Drive folders in the event YAML:
    drive.event_photos              — Event Photos (form + Activity Report)
    drive.cleanup_photos            — post-event cleanup / timestamped photos
    drive.event_files               — spreadsheet + generated budget PDFs
    drive.activity_report_template  — spreadsheet to copy for each event

This leaves Google Form submission and Gmail sending to the user for review.
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
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
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
from googleapiclient.http import MediaFileUpload

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError


# ---------------------------------------------------------------------------
# Paths / constants
# ---------------------------------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent
CREDENTIALS_FILE = BASE_DIR / "credentials.json"
TOKEN_FILE = BASE_DIR / "token.json"
BROWSER_PROFILE = BASE_DIR / "browser_profile"
GENERATED_DIR = BASE_DIR / "generated"
# Windows fonts used for Unicode support, including the ₹ symbol.
SEGOE_UI = Path(r"C:\Windows\Fonts\segoeui.ttf")
SEGOE_UI_BOLD = Path(r"C:\Windows\Fonts\segoeuib.ttf")

SCOPES = [
    "https://www.googleapis.com/auth/gmail.compose",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# Used when the YAML has no organisers, or to pad to 2 leads + 1 treasurer.
DUMMY_ORGANISERS = [
    {
        "name": "Dummy Lead Organiser 1",
        "roll_no": "2020001",
        "cgpa": "9.00",
        "phone": "9999990001",
        "email": "dummy.lead1@iiitd.ac.in",
        "role": "Lead Organiser",
    },
    {
        "name": "Dummy Lead Organiser 2",
        "roll_no": "2020002",
        "cgpa": "9.00",
        "phone": "9999990002",
        "email": "dummy.lead2@iiitd.ac.in",
        "role": "Lead Organiser",
    },
    {
        "name": "Dummy Treasurer",
        "roll_no": "2020003",
        "cgpa": "9.00",
        "phone": "9999990003",
        "email": "dummy.treasurer@iiitd.ac.in",
        "role": "Treasurer",
    },
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


def event_output_dir(config: Dict[str, Any]) -> Path:
    """generated/<event_name>/ from event.name in the YAML."""
    name = require_config(config, "event", "name")
    path = GENERATED_DIR / event_slug(name)
    path.mkdir(parents=True, exist_ok=True)
    return path


def ensure_dirs() -> None:
    GENERATED_DIR.mkdir(parents=True, exist_ok=True)


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
            # Google's time boxes are two characters wide (__ : __).
            return d.strftime("%I"), d.strftime("%M"), d.strftime("%p")
        except ValueError:
            pass

    die(f"Could not parse time '{value}'. Use e.g. '05:00 PM'.")


def organiser_people(config: Dict[str, Any]) -> List[Dict[str, Any]]:
    """
    At least 2 lead organisers and 1 treasurer. YAML entries are used
    first; dummy coordinator rows fill any shortfall.
    """
    people = [dict(p) for p in (config.get("organisers") or [])]
    roles = ["Lead Organiser", "Lead Organiser", "Treasurer"]
    while len(people) < 3:
        people.append(dict(DUMMY_ORGANISERS[len(people)]))
    for i, person in enumerate(people):
        person.setdefault("role", roles[i] if i < 3 else "Organiser")
        person.setdefault("name", DUMMY_ORGANISERS[min(i, 2)]["name"])
        person.setdefault("roll_no", "")
        person.setdefault("cgpa", "NIL")
        person.setdefault("phone", "")
        person.setdefault("email", "")
    return people


def organiser_text(config: Dict[str, Any]) -> str:
    lines = []
    for person in organiser_people(config):
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

def register_pdf_fonts() -> None:
    """Register Windows fonts so ReportLab can render Unicode characters."""

    if not SEGOE_UI.exists():
        die(f"Required font not found: {SEGOE_UI}")

    if not SEGOE_UI_BOLD.exists():
        die(f"Required font not found: {SEGOE_UI_BOLD}")

    if "SegoeUI" not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(
            TTFont("SegoeUI", str(SEGOE_UI))
        )

    if "SegoeUI-Bold" not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(
            TTFont("SegoeUI-Bold", str(SEGOE_UI_BOLD))
        )

def make_budget_pdf(
    config: Dict[str, Any],
    which: str,
) -> Path:
    ensure_dirs()
    register_pdf_fonts()

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
    output = event_output_dir(config) / filename

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "BudgetTitle",
        parent=styles["Title"],
        fontName="SegoeUI-Bold",
        fontSize=16,
        leading=20,
        alignment=TA_CENTER,
        spaceAfter=10,
    )
    
    heading_style = ParagraphStyle(
        "BudgetHeading",
        parent=styles["Heading2"],
        fontName="SegoeUI-Bold",
        fontSize=12,
        leading=15,
        spaceBefore=8,
        spaceAfter=5,
    )
    body_style = ParagraphStyle(
        "BudgetBody",
        parent=styles["BodyText"],
        fontName="SegoeUI",
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
                    ("FONTNAME", (0, 0), (-1, 0), "SegoeUI-Bold"),
                    ("FONTNAME", (0, 1), (-1, -2), "SegoeUI"),
                    ("FONTNAME", (0, -1), (-1, -1), "SegoeUI-Bold"),
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


def make_organisers_pdf(config: Dict[str, Any]) -> Path:
    """PDF of coordinator/organiser details for the Instructions file upload."""
    register_pdf_fonts()

    event_name = require_config(config, "event", "name")
    output = event_output_dir(config) / f"{event_slug(event_name)} Organisers.pdf"

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "OrgTitle",
        parent=styles["Title"],
        fontName="SegoeUI-Bold",
        fontSize=16,
        leading=20,
        alignment=TA_CENTER,
        spaceAfter=10,
    )
    body_style = ParagraphStyle(
        "OrgBody",
        parent=styles["BodyText"],
        fontName="SegoeUI",
        fontSize=10,
        leading=13,
    )
    heading_style = ParagraphStyle(
        "OrgHeading",
        parent=styles["Heading2"],
        fontName="SegoeUI-Bold",
        fontSize=11,
        leading=14,
        spaceAfter=6,
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
        Paragraph(f"{event_name} — Organiser Details", title_style),
        Paragraph(
            "At least 2 Lead Organisers and 1 Treasurer. "
            "For each member: Name, Roll No., CGPA, Phone Number, "
            "and Institute Email ID.",
            body_style,
        ),
        Spacer(1, 4 * mm),
        Paragraph("Coordinators / Organisers", heading_style),
    ]

    data = [["Role", "Name", "Roll No.", "CGPA", "Phone", "Email"]]
    for person in organiser_people(config):
        data.append(
            [
                str(person.get("role", "")),
                str(person.get("name", "")),
                str(person.get("roll_no", "")),
                str(person.get("cgpa", "NIL")),
                str(person.get("phone", "")),
                str(person.get("email", "")),
            ]
        )

    table = Table(data, colWidths=[28 * mm, 32 * mm, 22 * mm, 18 * mm, 28 * mm, 48 * mm])
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#eeeeee")),
                ("FONTNAME", (0, 0), (-1, 0), "SegoeUI-Bold"),
                ("FONTNAME", (0, 1), (-1, -1), "SegoeUI"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    story.append(table)
    doc.build(story)
    print(f"✓ Generated {output}")
    return output


# ---------------------------------------------------------------------------
# Google OAuth
# ---------------------------------------------------------------------------

def google_credentials() -> Credentials:
    creds = None

    if TOKEN_FILE.exists():
        try:
            stored = json.loads(TOKEN_FILE.read_text(encoding="utf-8"))
        except Exception:
            stored = {}
        granted = stored.get("scopes") or stored.get("scope") or []
        if isinstance(granted, str):
            granted = granted.split()
        if set(SCOPES).issubset(set(granted)):
            creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)
        else:
            print(
                "Google login needs extra Drive/Sheets permission. "
                "A browser window will open to sign in again."
            )
            creds = None

    if creds and creds.valid:
        return creds

    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
            TOKEN_FILE.write_text(creds.to_json(), encoding="utf-8")
            return creds
        except Exception:
            print(
                "Saved Google login could not be refreshed. "
                "A browser window will open to sign in again."
            )
            creds = None

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


def drive_service():
    return build("drive", "v3", credentials=google_credentials())


def sheets_service():
    return build("sheets", "v4", credentials=google_credentials())


def drive_folder_id(url: str) -> str:
    match = re.search(r"/folders/([a-zA-Z0-9_-]+)", url)
    if not match:
        die(f"Could not extract a Drive folder ID from: {url}")
    return match.group(1)


def drive_link(config: Dict[str, Any], key: str) -> str:
    return str(require_config(config, "drive", key)).strip()


def spreadsheet_id_from_url(url: str) -> str:
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", url)
    if not match:
        die(f"Could not extract a spreadsheet ID from: {url}")
    return match.group(1)


def club_name(config: Dict[str, Any]) -> str:
    return str(nested_get(config, "club", "name", default="") or "Club")


def club_email(config: Dict[str, Any]) -> str:
    return str(nested_get(config, "club", "email", default="") or "")


def _filled_cell_text_format() -> Dict[str, Any]:
    return {
        "fontFamily": "Space Grotesk",
        "fontSize": 9,
        "bold": False,
        "italic": False,
        "foregroundColor": {"red": 0, "green": 0, "blue": 0},
    }


def _format_filled_sheet_cells(sheets, spreadsheet_id: str) -> None:
    """Space Grotesk 9, not bold/italic, on title + data cells we write."""
    meta = (
        sheets.spreadsheets()
        .get(spreadsheetId=spreadsheet_id, fields="sheets.properties")
        .execute()
    )
    sheet_id = meta["sheets"][0]["properties"]["sheetId"]
    text_format = _filled_cell_text_format()
    ranges = [
        {"startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 8},
        {"startRowIndex": 3, "endRowIndex": 4, "startColumnIndex": 0, "endColumnIndex": 8},
    ]
    requests = []
    for cell_range in ranges:
        requests.append(
            {
                "repeatCell": {
                    "range": {"sheetId": sheet_id, **cell_range},
                    "cell": {
                        "userEnteredFormat": {
                            "textFormat": text_format,
                            "wrapStrategy": "WRAP",
                        }
                    },
                    "fields": (
                        "userEnteredFormat.textFormat.fontFamily,"
                        "userEnteredFormat.textFormat.fontSize,"
                        "userEnteredFormat.textFormat.bold,"
                        "userEnteredFormat.textFormat.italic,"
                        "userEnteredFormat.textFormat.foregroundColor,"
                        "userEnteredFormat.wrapStrategy"
                    ),
                }
            }
        )
    sheets.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": requests},
    ).execute()


def upload_file_to_event_folder(config: Dict[str, Any], file_path: Path) -> None:
    """Upload one file into drive.event_files, replacing a same-named file."""
    file_path = Path(file_path)
    if not file_path.exists():
        die(f"Upload file does not exist: {file_path}")

    folder_id = drive_folder_id(drive_link(config, "event_files"))
    drive = drive_service()
    print(f"  Uploading {file_path.name} to Drive...")
    media = MediaFileUpload(str(file_path), resumable=True)
    escaped = file_path.name.replace("\\", "\\\\").replace("'", "\\'")
    query = (
        f"name = '{escaped}' and '{folder_id}' in parents and trashed = false"
    )
    existing = (
        drive.files()
        .list(
            q=query,
            fields="files(id)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        )
        .execute()
        .get("files", [])
    )
    if existing:
        drive.files().update(
            fileId=existing[0]["id"],
            media_body=media,
            supportsAllDrives=True,
        ).execute()
    else:
        drive.files().create(
            body={"name": file_path.name, "parents": [folder_id]},
            media_body=media,
            fields="id",
            supportsAllDrives=True,
        ).execute()
    print(f"  ✓ Uploaded {file_path.name}")


def activity_report_audience_type(config: Dict[str, Any]) -> str:
    """
    Spreadsheet column: Open to all / IIITD only / Club Members only.
    Optional event.audience_scope overrides the guess from target_audience.
    """
    explicit = nested_get(config, "event", "audience_scope", default="")
    if explicit:
        return str(explicit)
    audience = config.get("event", {}).get("target_audience") or []
    joined = " ".join(str(a).lower() for a in audience)
    if "club" in joined:
        return "Club Members only"
    if "external" in joined or "open to all" in joined:
        return "Open to all"
    return "IIITD only"


def create_activity_report_sheet(config: Dict[str, Any]) -> str:
    """
    Copy the Activity Report template into drive.event_files and fill one data row.
    Returns the new spreadsheet URL.
    """
    template_url = drive_link(config, "activity_report_template")
    template_id = spreadsheet_id_from_url(template_url)
    folder_url = drive_link(config, "event_files")
    photos_url = drive_link(config, "event_photos")
    folder_id = drive_folder_id(folder_url)
    event_name = require_config(config, "event", "name")
    event = config["event"]
    after = config.get("after_event", {}) or {}

    drive = drive_service()
    print("  Copying Activity Report template into the event Drive folder...")

    copied = (
        drive.files()
        .copy(
            fileId=template_id,
            body={"name": event_name},
            fields="id,parents,webViewLink",
            supportsAllDrives=True,
        )
        .execute()
    )
    file_id = copied["id"]
    old_parents = ",".join(copied.get("parents") or [])
    (
        drive.files()
        .update(
            fileId=file_id,
            addParents=folder_id,
            removeParents=old_parents,
            fields="id,parents,webViewLink",
            supportsAllDrives=True,
        )
        .execute()
    )

    sheet_url = copied.get("webViewLink") or (
        f"https://docs.google.com/spreadsheets/d/{file_id}/edit"
    )

    organisers = "\n".join(
        str(p.get("name", "")).strip()
        for p in organiser_people(config)
        if str(p.get("name", "")).strip()
        and not str(p.get("name", "")).startswith("Dummy ")
    )
    if not organisers:
        organisers = organiser_names(config)

    participants = after.get("participants_attended", event.get("expected_participants", ""))
    remarks = after.get("remarks") or "NIL"

    sheets = sheets_service()
    sheets.spreadsheets().values().update(
        spreadsheetId=file_id,
        range="A1",
        valueInputOption="USER_ENTERED",
        body={"values": [[event_name]]},
    ).execute()
    sheets.spreadsheets().values().update(
        spreadsheetId=file_id,
        range="A4:H4",
        valueInputOption="USER_ENTERED",
        body={
            "values": [
                [
                    display_date(event["date"]),
                    event_name,
                    (event.get("description") or "").strip(),
                    activity_report_audience_type(config),
                    str(participants),
                    organisers,
                    photos_url,
                    remarks,
                ]
            ]
        },
    ).execute()

    _format_filled_sheet_cells(sheets, file_id)

    print(f"✓ Activity Report spreadsheet: {sheet_url}")
    return sheet_url


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

This is with regards to the event that the club is conducting.
Following are the details:

Event Name: {event.name}
Event Date: {event.date}
Event Time: {event.start_time} to {event.end_time}
Proposed Venue: {event.venue}
Expected Number of Participants: {event.expected_participants}
Description of Event: {event.description}
Budget: {event.budget}

Best regards,
Team""",
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

This is with regards to the event the club is conducting - {event.name}.

We want to book {event.venue} on {event.date} from 4 to 7 pm.
Kindly do the needful.

Best regards,
Team""",
        )

        room_body = render_template(room_body, config)

        draft_id = create_gmail_draft(
            service,
            room.get("to", []),
            room.get("cc", []),
            room.get("subject", "Room Booking"),
            room_body,
            [Path(x) for x in room.get("attachments", [])],
        )
        print(f"✓ Room booking email draft created: {draft_id}")

    # After-event activity-report email is not used.


# ---------------------------------------------------------------------------
# Google Forms / Playwright
# ---------------------------------------------------------------------------

def launch_browser():
    """
    Launch a persistent browser profile for the club Google account.

    The first run may require manual Google login. The same browser
    profile is reused on subsequent runs.
    """
    playwright = sync_playwright().start()

    context = playwright.chromium.launch_persistent_context(
        user_data_dir=str(BROWSER_PROFILE),
        headless=False,
        viewport={"width": 1400, "height": 1000},
        args=[
            "--disable-blink-features=AutomationControlled",
        ],
    )

    page = context.pages[0] if context.pages else context.new_page()

    return playwright, context, page


def wait_until_browser_closed(page, context, *, filled: bool = True) -> None:
    """
    Leave the browser open. Do not submit. Exit only after the user
    closes the window.
    """
    if filled:
        print("\n" + "=" * 60)
        print("FORM IS FILLED AND HAS NOT BEEN SUBMITTED")
        print("=" * 60)
        print("The browser is yours. Review or edit anything you want.")
        print("This script will not click Submit.")
        print("Close the browser window when you are done.")
        print("=" * 60 + "\n")
    else:
        print("\nClose the browser window when you are done.\n")

    if page is None and context is None:
        return

    while True:
        try:
            if page is None or page.is_closed():
                break
            if context is None or len(context.pages) == 0:
                break
            page.wait_for_timeout(500)
        except Exception:
            break


def wait_for_google_login(page, form_url: str, config: Optional[Dict[str, Any]] = None) -> None:
    """
    Open the Google Form and wait for the user to authenticate manually
    if Google redirects to the login/account page.
    """
    account = club_email(config or {}) or "your club Google account"

    print("\nOpening Google Form...")

    try:
        page.goto(
            form_url,
            wait_until="domcontentloaded",
            timeout=30000,
        )
    except Exception as e:
        # Google sometimes redirects /viewform to
        # /viewform?authuser=0 while Playwright is waiting.
        # If the browser is already on the form, this is harmless.
        if "interrupted by another navigation" not in str(e):
            raise

    page.wait_for_timeout(3000)

    page.wait_for_timeout(3000)

    current_url = page.url

    # Google login/account pages.
    if (
        "accounts.google.com" in current_url
        or "signin" in current_url.lower()
    ):
        print("\n" + "=" * 60)
        print("GOOGLE LOGIN REQUIRED")
        print("=" * 60)
        print()
        print("Please log into the club Google account:")
        print()
        print(f"    {account}")
        print()
        print("IMPORTANT:")
        print("Do NOT use your personal Google account.")
        print()
        print("Complete the Google login in the browser.")
        print("After the Google Form becomes visible,")
        print("return to this terminal and press ENTER.")
        print("=" * 60)

        input("\nPress ENTER after the Google Form is visible...")

        # Re-open the form after authentication.

        try:
            page.goto(
                form_url,
                wait_until="domcontentloaded",
                timeout=30000,
            )
        except Exception as e:
            # Google sometimes redirects /viewform to
            # /viewform?authuser=0 while Playwright is waiting.
            # If the browser is already on the form, this is harmless.
            if "interrupted by another navigation" not in str(e):
                raise

        page.wait_for_timeout(3000)

        page.wait_for_timeout(3000)

    # Check whether we are still on a Google login page.
    current_url = page.url

    if (
        "accounts.google.com" in current_url
        or "signin" in current_url.lower()
    ):
        die(
            "Google login was not completed. "
            f"Please log into {account}."
        )

    print("✓ Google Form opened successfully")

    try:
        page.wait_for_selector('[role="listitem"]', timeout=20000)
    except PlaywrightTimeoutError:
        die("The Google Form questions did not load.")


def visible_question_container(page, question: str, occurrence: int = 0):
    """
    Robust Google Forms question locator.

    Google Forms renders questions dynamically and the question container
    usually contains more text than just the question title.

    Example:

        Event Title
        Your answer

    Therefore we search for the question text inside the container rather
    than requiring the entire container text to equal the question.
    """

    question_re = re.compile(
        rf"^\s*{re.escape(question)}\s*\*?\s*$",
        re.I,
    )

    def text_matches_question(text: str) -> bool:
        collapsed = re.sub(r"\s+", " ", text).strip()
        wanted = re.sub(r"\s+", " ", question).strip()
        if question_re.fullmatch(collapsed):
            return True
        if re.search(rf"^\s*{re.escape(wanted)}\s*\*?\s*$", collapsed, re.I):
            return True
        if wanted.lower() in collapsed.lower():
            return True
        # Titles like "Foo (Mandatory)" often wrap before the parentheses.
        without_paren = re.sub(r"\s*\([^)]*\)\s*$", "", wanted).strip()
        if without_paren and without_paren.lower() in collapsed.lower():
            return True
        return False

    # ------------------------------------------------------------
    # Helper: search currently rendered DOM
    # ------------------------------------------------------------

    def find_question():
        # Google Forms normally uses role=listitem for question blocks.
        # After clicking Next, previous-page items often remain in the DOM
        # but hidden — skip those so we only fill the current page.
        matches = []
        containers = page.locator('[role="listitem"]')

        count = containers.count()

        for i in range(count):
            container = containers.nth(i)

            try:
                if not container.is_visible():
                    continue

                text = container.inner_text(timeout=1000)
                matched = text_matches_question(text)
                if not matched:
                    for line in text.splitlines():
                        line = line.strip()
                        if line and (
                            question_re.fullmatch(line)
                            or text_matches_question(line)
                        ):
                            matched = True
                            break
                if matched:
                    matches.append(container)

            except Exception:
                continue

        if occurrence < len(matches):
            return matches[occurrence]

        locator = page.get_by_text(question_re)
        if locator.count() > occurrence:
            try:
                candidate = locator.nth(occurrence)

                for _ in range(10):
                    role = candidate.get_attribute("role")

                    if role == "listitem":
                        return candidate

                    if candidate.locator(
                        "input, textarea, [role=radio], [role=checkbox]"
                    ).count() > 0:
                        return candidate

                    candidate = candidate.locator("xpath=..")

            except Exception:
                pass

        return None

    # ------------------------------------------------------------
    # First search
    # ------------------------------------------------------------

    container = find_question()

    if container is not None:
        try:
            container.scroll_into_view_if_needed()
        except Exception:
            pass

        return container

    # ------------------------------------------------------------
    # Progressive scrolling
    # ------------------------------------------------------------

    print(f"  Searching form for: {question}")

    last_position = -1

    for _ in range(40):

        container = find_question()

        if container is not None:
            try:
                container.scroll_into_view_if_needed()
                page.wait_for_timeout(500)
            except Exception:
                pass

            return container

        # Get current browser scroll position.
        position = page.evaluate(
            "() => window.scrollY"
        )

        # Scroll down using JavaScript rather than mouse wheel.
        page.evaluate(
            """
            () => {
                window.scrollBy({
                    top: Math.max(window.innerHeight * 0.75, 600),
                    left: 0,
                    behavior: "instant"
                });
            }
            """
        )

        page.wait_for_timeout(500)

        new_position = page.evaluate(
            "() => window.scrollY"
        )

        # If scrolling stopped, try one final search.
        if new_position == position or new_position == last_position:
            break

        last_position = new_position

    # ------------------------------------------------------------
    # Final search
    # ------------------------------------------------------------

    container = find_question()

    if container is not None:
        try:
            container.scroll_into_view_if_needed()
        except Exception:
            pass

        return container

    return None


def _visible_match(locator):
    """Return the last currently visible locator match, or None."""
    found = None
    for i in range(locator.count()):
        candidate = locator.nth(i)
        try:
            if candidate.is_visible():
                found = candidate
        except Exception:
            continue
    return found


def _commit_text_value(locator, value: Any) -> None:
    """
    Type into a Google Forms input and blur so the form's JS state
    actually records the value. A plain fill() can look filled in the
    UI while Next still treats the question as empty.
    """
    locator.click()
    locator.fill(str(value))
    locator.dispatch_event("input")
    locator.dispatch_event("change")
    try:
        locator.press("Tab")
    except Exception:
        try:
            locator.blur()
        except Exception:
            pass


def fill_text_question(
    page, question: str, value: str, occurrence: int = 0
) -> None:
    container = visible_question_container(page, question, occurrence=occurrence)

    if container is None:
        die(
            f"Could not locate Google Form question after searching "
            f"the rendered form: {question}"
        )

    inputs = container.locator(
        "input:not([type=radio]):not([type=checkbox]):not([type=file]), textarea"
    )

    if inputs.count() == 0:
        die(
            f"Found question '{question}' but could not find its "
            f"text input."
        )

    inputs.first.scroll_into_view_if_needed()
    _commit_text_value(inputs.first, value)

def click_option(
    page,
    question: str,
    option: str,
    checkbox: bool = False,
) -> None:
    container = visible_question_container(page, question)

    if container is None:
        die(f"Could not locate Google Form question: {question}")

    option_re = re.compile(
        rf"^\s*{re.escape(option)}\s*$",
        re.I,
    )

    # First try the text inside the question container.
    label = container.get_by_text(option_re).first

    if label.count() == 0:
        # Fallback to the actual radio/checkbox element.
        role = "checkbox" if checkbox else "radio"

        options = container.locator(
            f'[role="{role}"]'
        )

        for i in range(options.count()):
            option_el = options.nth(i)

            aria_label = option_el.get_attribute("aria-label") or ""

            if re.fullmatch(option_re, aria_label):
                option_el.click()
                page.wait_for_timeout(300)
                return

    if label.count() == 0:
        die(
            f"Could not locate option '{option}' "
            f"for '{question}'."
        )

    label.click()
    page.wait_for_timeout(300)


def _first_visible_input(container):
    inputs = container.locator("input")
    for i in range(inputs.count()):
        candidate = inputs.nth(i)
        try:
            if candidate.is_visible():
                return candidate
        except Exception:
            continue
    return inputs.first if inputs.count() else None


def _fill_time_box(locator, digits: str) -> None:
    """
    Google Forms hour/minute boxes often ignore fill() and reset if
    Tab is pressed before both parts are set. Type the digits instead.
    """
    locator.scroll_into_view_if_needed()
    locator.click()
    locator.press("Control+A")
    for _ in range(4):
        locator.press("Backspace")
    locator.type(str(digits), delay=50)

    try:
        actual = (locator.input_value() or "").strip()
    except Exception:
        actual = ""

    if actual not in {str(digits), str(digits).lstrip("0") or "0"}:
        locator.evaluate(
            """(el, value) => {
                el.focus();
                const setter = Object.getOwnPropertyDescriptor(
                    window.HTMLInputElement.prototype,
                    "value"
                ).set;
                setter.call(el, value);
                el.dispatchEvent(new Event("input", { bubbles: true }));
                el.dispatchEvent(new Event("change", { bubbles: true }));
            }""",
            str(digits),
        )


def fill_date_question(page, question: str, value: str) -> None:
    container = visible_question_container(page, question)
    if container is None:
        die(f"Could not locate date question: {question}")

    inp = _first_visible_input(container)
    if inp is None:
        die(f"Could not find date input for: {question}")

    d = parse_date(value)
    input_type = (inp.get_attribute("type") or "").lower()
    placeholder = (inp.get_attribute("placeholder") or "").lower()

    inp.scroll_into_view_if_needed()
    inp.click()

    # Native <input type="date"> still wants ISO. The visible
    # placeholder on this form is dd-mm-yyyy.
    if input_type == "date":
        inp.fill(d.strftime("%Y-%m-%d"))
    elif "dd" in placeholder or "mm" in placeholder:
        inp.fill("")
        inp.type(d.strftime("%d-%m-%Y"), delay=30)
    else:
        inp.fill(d.strftime("%d-%m-%Y"))

    inp.dispatch_event("input")
    inp.dispatch_event("change")
    # Blur without Tabbing into the neighbouring time widget.
    try:
        container.click(position={"x": 12, "y": 8})
    except Exception:
        pass


def fill_time_question(page, question: str, value: str) -> None:
    container = visible_question_container(page, question)
    if container is None:
        die(f"Could not locate time question: {question}")

    hour, minute, ampm = split_time(value)

    hour_el = None
    minute_el = None

    labeled_hour = container.locator(
        'input[aria-label="Hour"], input[aria-label="Hour of day"]'
    )
    labeled_minute = container.locator(
        'input[aria-label="Minute"], input[aria-label="Minute of hour"]'
    )
    if labeled_hour.count() and labeled_minute.count():
        hour_el = labeled_hour.first
        minute_el = labeled_minute.first
    else:
        by_hour = container.get_by_label(re.compile(r"^hour", re.I))
        by_minute = container.get_by_label(re.compile(r"^minute", re.I))
        if by_hour.count() and by_minute.count():
            hour_el = by_hour.first
            minute_el = by_minute.first

    if hour_el is None or minute_el is None:
        visible_inputs = []
        text_inputs = container.locator("input:not([type=hidden]):not([type=radio]):not([type=checkbox]):not([type=file])")
        for i in range(text_inputs.count()):
            candidate = text_inputs.nth(i)
            try:
                if candidate.is_visible():
                    visible_inputs.append(candidate)
            except Exception:
                continue
        if len(visible_inputs) >= 2:
            hour_el = visible_inputs[0]
            minute_el = visible_inputs[1]

    if hour_el is None or minute_el is None:
        die(f"Could not find hour/minute inputs for: {question}")

    _fill_time_box(hour_el, hour)
    _fill_time_box(minute_el, minute)
    _select_ampm(page, container, ampm)


def _select_ampm(page, container, ampm: str) -> None:
    """
    Google Forms keeps AM/PM <role=option> nodes in the DOM even when
    the menu is closed. Those nodes are not visible; clicking them
    hangs until Playwright times out. Open the listbox first, then
    choose the visible option (keyboard fallback if the popup is flaky).
    """
    wanted = ampm.strip().upper()
    combo = container.locator('[role="listbox"], [role="combobox"]')
    if combo.count() == 0:
        return

    toggle = combo.first
    toggle.scroll_into_view_if_needed()

    displayed = re.sub(r"\s+", " ", toggle.inner_text() or "").strip().upper()
    if displayed == wanted:
        return

    # Leave the hour/minute inputs so the listbox can receive the click.
    page.keyboard.press("Escape")
    page.wait_for_timeout(150)
    try:
        toggle.click()
    except Exception:
        toggle.click(force=True)

    popup = page.locator(
        f'[role="option"][data-value="{wanted}"], '
        f'[role="option"][data-value="{wanted.lower()}"]'
    )

    opened = False
    deadline = time.time() + 3
    while time.time() < deadline:
        for i in range(popup.count()):
            option = popup.nth(i)
            try:
                if option.is_visible():
                    option.click()
                    page.wait_for_timeout(200)
                    opened = True
                    break
            except Exception:
                continue
        if opened:
            break
        page.wait_for_timeout(100)

    if not opened:
        # Closed-menu options are in the DOM but hidden. Drive the
        # focused listbox with keys instead of clicking a hidden node.
        toggle.focus()
        page.keyboard.press("ArrowDown")
        page.keyboard.press("Enter")
        page.wait_for_timeout(200)

    displayed = re.sub(r"\s+", " ", toggle.inner_text() or "").strip().upper()
    if displayed == wanted:
        return

    page.evaluate(
        """(wanted) => {
            const visible = [...document.querySelectorAll('[role="option"]')]
                .find((el) => {
                    const style = window.getComputedStyle(el);
                    const value = (el.getAttribute("data-value") || el.textContent || "")
                        .trim()
                        .toUpperCase();
                    return value === wanted
                        && style.visibility !== "hidden"
                        && style.display !== "none"
                        && el.offsetParent !== null;
                });
            if (visible) {
                visible.click();
                return;
            }
            const any = [...document.querySelectorAll('[role="option"]')]
                .find((el) =>
                    (el.getAttribute("data-value") || "").toUpperCase() === wanted
                );
            if (any) any.click();
        }""",
        wanted,
    )
    page.wait_for_timeout(200)


def click_record_email(page) -> None:
    """
    The Email question is a required checkbox:

        Record the signed-in address as the email to be included
        with my response

    Click the checkbox control, not the email text — the address can
    steal the click.
    """
    print("  Confirming email checkbox...")

    checkbox = None
    container = visible_question_container(page, "Email")
    if container is not None:
        boxes = container.locator('[role="checkbox"], input[type="checkbox"]')
        if boxes.count():
            checkbox = boxes.first

    if checkbox is None:
        named = page.get_by_role(
            "checkbox",
            name=re.compile(
                r"Record .* as the email to be included with my response",
                re.I,
            ),
        )
        if named.count():
            checkbox = named.first

    if checkbox is None:
        die(
            "Could not find the email confirmation checkbox "
            "('Record ... as the email to be included with my response')."
        )

    checkbox.scroll_into_view_if_needed()

    def is_checked() -> bool:
        aria = (checkbox.get_attribute("aria-checked") or "").lower()
        if aria == "true":
            return True
        try:
            return bool(checkbox.is_checked())
        except Exception:
            return False

    if not is_checked():
        try:
            checkbox.click()
        except Exception:
            checkbox.click(force=True)
        page.wait_for_timeout(400)

    if not is_checked():
        checkbox.click(force=True)
        page.wait_for_timeout(400)

    if not is_checked():
        die(
            "Could not tick the email confirmation checkbox. "
            "Make sure the club Google account is signed in."
        )

    print("  ✓ Email recording confirmed")


GOOGLE_FORM_UPLOAD_MAX_BYTES = 10 * 1024 * 1024


def _file_inputs_in_any_frame(page):
    found = []
    for frame in page.frames:
        try:
            loc = frame.locator('input[type="file"]')
            if loc.count():
                found.append(loc.last)
        except Exception:
            continue
    return found


def _click_in_any_frame(page, builders) -> bool:
    """
    builders: list of callables frame -> locator
    Clicks the first match in any frame.
    """
    for frame in page.frames:
        for builder in builders:
            try:
                loc = builder(frame)
                if loc.count() == 0:
                    continue
                target = loc.last
                if hasattr(target, "is_visible") and not target.is_visible():
                    # Still try — picker tabs can report as not visible.
                    pass
                target.click(timeout=3000, force=True)
                return True
            except Exception:
                continue
    return False


def _wait_until_file_attached(page, file_path: Path, timeout_ms: int = 20000) -> bool:
    """True only if THIS budget PDF's name is visible on the form."""
    name = file_path.name
    stem = file_path.stem
    tokens = [name, stem]
    if " " in stem:
        tokens.append(stem.replace(" ", "_"))
    pattern = re.compile(
        "|".join(re.escape(t) for t in tokens),
        re.I,
    )
    deadline = time.time() + timeout_ms / 1000
    while time.time() < deadline:
        try:
            if page.get_by_text(pattern).count():
                loc = page.get_by_text(pattern).first
                if loc.is_visible():
                    return True
            body = page.inner_text("body")
            if pattern.search(body or ""):
                return True
        except Exception:
            pass
        page.wait_for_timeout(300)
    return False


def _dismiss_file_picker(page) -> None:
    for _ in range(4):
        try:
            page.keyboard.press("Escape")
        except Exception:
            break
        page.wait_for_timeout(200)


def _attach_file_via_picker(page, file_path: Path) -> bool:
    """Drive picker: Upload tab -> file input or Browse -> Insert."""
    page.wait_for_timeout(800)

    _click_in_any_frame(
        page,
        [
            lambda f: f.get_by_role("tab", name=re.compile(r"^Upload", re.I)),
            lambda f: f.get_by_text(re.compile(r"^Upload$", re.I)),
            lambda f: f.locator('[aria-label="Upload"], [data-id="upload"]'),
        ],
    )
    page.wait_for_timeout(800)

    inputs = _file_inputs_in_any_frame(page)
    if inputs:
        inputs[-1].set_input_files(str(file_path))
    else:
        try:
            with page.expect_file_chooser(timeout=8000) as chooser_info:
                clicked = _click_in_any_frame(
                    page,
                    [
                        lambda f: f.get_by_role(
                            "button", name=re.compile(r"Browse", re.I)
                        ),
                        lambda f: f.get_by_text(re.compile(r"^Browse$", re.I)),
                    ],
                )
                if not clicked:
                    raise PlaywrightTimeoutError("Browse not found")
            chooser_info.value.set_files(str(file_path))
        except PlaywrightTimeoutError:
            inputs = _file_inputs_in_any_frame(page)
            if not inputs:
                return False
            inputs[-1].set_input_files(str(file_path))

    page.wait_for_timeout(1500)
    _click_in_any_frame(
        page,
        [
            lambda f: f.get_by_role(
                "button", name=re.compile(r"^(Insert|Upload|Select)$", re.I)
            ),
            lambda f: f.get_by_text(re.compile(r"^(Insert|Upload|Select)$", re.I)),
        ],
    )
    return _wait_until_file_attached(page, file_path)


def _confirm_budget_upload(page, file_path: Path) -> None:
    _dismiss_file_picker(page)
    page.wait_for_timeout(400)
    if not _wait_until_file_attached(page, file_path, timeout_ms=8000):
        die(
            f"Upload finished but '{file_path.name}' is not shown on the form. "
            "Remove whatever file was attached and retry."
        )
    print(f"  ✓ Attached file: {file_path.name}")
    print(f"    {file_path}")


def upload_file_question(page, question_text: str, file_path: Path) -> None:
    """
    Google Forms file questions show an Add file button (max 10 MB on
    the before/after Google Forms). Clicking it opens a Drive picker
    inside iframes, not a file input on the form page.
    """
    file_path = Path(file_path)
    if not file_path.exists():
        die(f"Upload file does not exist: {file_path}")

    size = file_path.stat().st_size
    if size > GOOGLE_FORM_UPLOAD_MAX_BYTES:
        die(
            f"Budget PDF is {size} bytes; both Google Forms only accept "
            f"files up to 10 MB."
        )

    print(f"  Uploading file:\n    {file_path}")

    if _wait_until_file_attached(page, file_path, timeout_ms=1000):
        _confirm_budget_upload(page, file_path)
        return

    container = visible_question_container(page, question_text)

    def find_add_file(root):
        locators = [
            root.get_by_role("button", name=re.compile(r"Add file", re.I)),
            root.locator('[role="button"]').filter(
                has_text=re.compile(r"^\s*Add file\s*$", re.I)
            ),
            root.get_by_text(re.compile(r"^\s*Add file\s*$", re.I)),
            root.locator('[aria-label="Add file"], [aria-label="Add File"]'),
        ]
        for loc in locators:
            try:
                found = _visible_match(loc)
                if found is not None:
                    return found
            except Exception:
                continue
        return None

    add_btn = find_add_file(container) if container is not None else None
    if add_btn is None:
        add_btn = find_add_file(page)

    if add_btn is None:
        die(
            f"Could not find a visible Add file button for: {question_text}."
        )

    try:
        add_btn.scroll_into_view_if_needed(timeout=5000)
    except Exception:
        pass
    page.wait_for_timeout(300)

    clicked = False
    try:
        with page.expect_file_chooser(timeout=4000) as chooser_info:
            add_btn.click(force=True)
            clicked = True
        chooser_info.value.set_files(str(file_path))
        if _wait_until_file_attached(page, file_path, timeout_ms=12000):
            _confirm_budget_upload(page, file_path)
            return
    except PlaywrightTimeoutError:
        pass

    if not clicked:
        add_btn.click(force=True, timeout=5000)

    page.wait_for_timeout(1000)

    inputs = _file_inputs_in_any_frame(page)
    if inputs:
        inputs[-1].set_input_files(str(file_path))
        page.wait_for_timeout(1000)
        _click_in_any_frame(
            page,
            [
                lambda f: f.get_by_role(
                    "button", name=re.compile(r"^(Insert|Upload|Select)$", re.I)
                ),
            ],
        )
        if _wait_until_file_attached(page, file_path):
            _confirm_budget_upload(page, file_path)
            return

    if _attach_file_via_picker(page, file_path):
        _confirm_budget_upload(page, file_path)
        return

    die(
        f"Could not attach the file for: {question_text}. "
        "The Drive picker did not accept the file."
    )


def click_google_form_next(page, wait_for: str) -> None:
    """
    Click Next on a multi-section Google Form and wait for the next page.

    Google Forms does not use a native <button>. The control is a
    div[role=button] with jsname=OCpkoe. Page 2 is confirmed when the
    Back control (jsname=GeGHKb) appears, then we wait for wait_for.
    """
    print("  Moving to the next form page...")

    # Commit any focused field and close leftover menus.
    page.keyboard.press("Escape")
    page.wait_for_timeout(400)

    next_button = (
        _visible_match(page.locator('[jsname="OCpkoe"]'))
        or _visible_match(
            page.get_by_role(
                "button",
                name=re.compile(r"(Next|Continue)", re.I),
            )
        )
        or _visible_match(
            page.locator('[role="button"]').filter(
                has_text=re.compile(r"^\s*(Next|Continue)\s*$", re.I)
            )
        )
    )

    if next_button is None:
        die("Could not locate the Google Form Next button.")

    next_button.scroll_into_view_if_needed()
    try:
        next_button.click(timeout=5000)
    except Exception:
        try:
            next_button.click(force=True, timeout=5000)
        except Exception:
            page.evaluate(
                """() => {
                    const el = document.querySelector('[jsname="OCpkoe"]');
                    if (el) el.click();
                }"""
            )

    back = page.locator('[jsname="GeGHKb"]')
    marker = page.get_by_text(
        re.compile(rf"{re.escape(wait_for)}", re.I)
    ).first

    def page_advanced() -> bool:
        try:
            if back.count() and back.first.is_visible():
                return True
        except Exception:
            pass
        history = page.evaluate(
            """() => {
                const el = document.querySelector('input[name="pageHistory"]');
                return el ? el.value : '';
            }"""
        )
        return "," in (history or "")

    try:
        back.first.wait_for(state="visible", timeout=8000)
    except PlaywrightTimeoutError:
        if not page_advanced():
            page.evaluate(
                """() => {
                    const el = document.querySelector('[jsname="OCpkoe"]');
                    if (el) el.click();
                }"""
            )
            try:
                back.first.wait_for(state="visible", timeout=8000)
            except PlaywrightTimeoutError:
                extra = ""
                required = page.get_by_text(
                    re.compile(r"This is a required question", re.I)
                )
                if required.count():
                    extra = (
                        " Required-question errors are visible, so Google "
                        "did not accept one or more answers on this page."
                    )
                die(
                    "The form did not advance after clicking Next "
                    f"(looking for '{wait_for}').{extra}"
                )

    try:
        marker.wait_for(state="visible", timeout=10000)
    except PlaywrightTimeoutError:
        die(
            f"Moved past page 1, but '{wait_for}' did not become visible."
        )

    print(f"  ✓ Next page opened ({wait_for})")
    page.wait_for_timeout(500)


def fill_organiser_instructions(
    page, config: Dict[str, Any], organisers_pdf: Path
) -> None:
    """
    Page 1 'Instructions:' is a required PDF upload (organisers), not a
    text box. Fall back to typing only if no Add file control is present.
    """
    container = visible_question_container(page, "Instructions:")
    if container is None:
        die("Could not locate Google Form question: Instructions:")

    looks_like_upload = (
        container.get_by_text(
            re.compile(r"Add file|Upload 1 supported file", re.I)
        ).count()
        > 0
        or container.locator("input[type=file]").count() > 0
    )
    if looks_like_upload:
        upload_file_question(page, "Instructions:", organisers_pdf)
        return

    fill_text_question(page, "Instructions:", organiser_text(config))


def fill_before_form(
    page, config: Dict[str, Any], budget_pdf: Path, organisers_pdf: Path
) -> None:
    event = config["event"]

    form_url = config["forms"]["before_event"]["url"]

    wait_for_google_login(page, form_url, config)

    click_record_email(page)

    click_option(page, "Proposing Body", "Club (Specify Below)")
    fill_text_question(page, 'If "Club", please mention the name of the club:', club_name(config))
    fill_text_question(page, 'If "Other", specify the group\'s name', "NA")

    fill_organiser_instructions(page, config, organisers_pdf)

    click_google_form_next(page, "Event Title")

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
            "Post-Event Clean-Up Responsibility",
            cleanup,
        )

    print("\n✓ BEFORE form has been filled. It has NOT been submitted.")


def fill_after_form(
    page,
    config: Dict[str, Any],
    budget_pdf: Path,
    organisers_pdf: Path,
    activity_report_url: str,
) -> None:
    event = config["event"]

    form_url = config["forms"]["after_event"]["url"]

    wait_for_google_login(page, form_url, config)

    click_record_email(page)

    click_option(page, "Event Conducted", "Club (Specify Below)")
    fill_text_question(page, 'If "Club", please mention the name of the club:', club_name(config))
    fill_text_question(page, 'If "Other", specify the group\'s name', "NA")

    fill_organiser_instructions(page, config, organisers_pdf)

    click_google_form_next(page, "Event Title")

    fill_text_question(page, "Event Title", event["name"])
    click_option(page, "Type of Event", event.get("type", "Competition"))
    fill_text_question(page, "Description of the Event", event["description"])

    fill_date_question(page, "Event Start Date", event["date"])
    fill_time_question(page, "Event Start Time", event["start_time"])
    fill_date_question(page, "End Date", event["date"])
    fill_time_question(page, "End Time", event["end_time"])

    fill_text_question(page, "Venue", event["venue"])

    participants = nested_get(config, "after_event", "participants_attended", default=None)
    if participants is None or participants == "":
        die(
            "Missing after_event.participants_attended in the event YAML."
        )
    fill_text_question(
        page,
        "Number of Participants attended:",
        participants,
    )

    for audience in event.get("target_audience", ["Institute Students"]):
        click_option(page, "Target Audience. Tick all that are applicable", audience, checkbox=True)

    amount_used = nested_get(config, "after_event", "amount_used", default=None)
    if amount_used is None or amount_used == "":
        die("Missing after_event.amount_used in the event YAML.")
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
        die(
            "Missing after_event.cleanup_responsibility in the event YAML "
            '(e.g. cleanup_responsibility: ["Name, RollNo"]).'
        )
    fill_text_question(
        page,
        "Post-Event Clean-Up Responsibility",
        cleanup,
        occurrence=0,
    )

    fill_text_question(
        page,
        "Post-Event Clean-Up Responsibility",
        drive_link(config, "cleanup_photos"),
        occurrence=1,
    )

    fill_text_question(
        page,
        "Link to the Activity Report (Mandatory for clubs)",
        activity_report_url,
    )

    fill_text_question(
        page,
        "Photos of the Event",
        drive_link(config, "event_photos"),
    )

    print("\n✓ AFTER form has been filled. It has NOT been submitted.")


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
    require_config(config, "drive", "event_files")

    print("\n========================================")
    print("FOOBAR EVENT AUTOMATION - BEFORE EVENT")
    print("========================================")
    print(f"Event: {config['event']['name']}\n")

    budget_pdf = make_budget_pdf(config, "expected")
    organisers_pdf = make_organisers_pdf(config)
    upload_file_to_event_folder(config, budget_pdf)

    playwright = context = page = None

    try:
        playwright, context, page = launch_browser()
        fill_before_form(page, config, budget_pdf, organisers_pdf)
        create_email_drafts(config, "before", budget_pdf)
        wait_until_browser_closed(page, context)

    except Exception as e:
        print("\n" + "=" * 60)
        print("ERROR WHILE PROCESSING GOOGLE FORM")
        print("=" * 60)
        print(e)
        print()
        print("The browser will remain open. Close the window when you are done.")
        wait_until_browser_closed(page, context, filled=False)

    finally:
        try:
            if context:
                context.close()
        except Exception:
            pass
        if playwright:
            playwright.stop()

    print("\n========================================")
    print("BEFORE EVENT COMPLETE")
    print("========================================")
    print("✓ Expected budget PDF generated and uploaded to Drive")
    print("✓ Google Form left open for you (not submitted by the script)")
    print("✓ Gmail drafts created")
    print("\nRemember to review and send the Gmail drafts manually.")


def run_after(config: Dict[str, Any]) -> None:
    ensure_dirs()

    require_config(config, "event", "name")
    require_config(config, "event", "date")
    require_config(config, "event", "venue")
    require_config(config, "forms", "after_event", "url")
    require_config(config, "drive", "event_files")
    require_config(config, "drive", "event_photos")
    require_config(config, "drive", "cleanup_photos")
    require_config(config, "drive", "activity_report_template")

    print("\n========================================")
    print("FOOBAR EVENT AUTOMATION - AFTER EVENT")
    print("========================================")
    print(f"Event: {config['event']['name']}\n")

    budget_pdf = make_budget_pdf(config, "actual")
    organisers_pdf = make_organisers_pdf(config)
    activity_report_url = create_activity_report_sheet(config)
    config["_activity_report_url"] = activity_report_url
    upload_file_to_event_folder(config, budget_pdf)

    playwright = context = page = None
    try:
        playwright, context, page = launch_browser()
        fill_after_form(
            page, config, budget_pdf, organisers_pdf, activity_report_url
        )
        wait_until_browser_closed(page, context)
    except Exception as e:
        print("\n" + "=" * 60)
        print("ERROR WHILE PROCESSING GOOGLE FORM")
        print("=" * 60)
        print(e)
        print()
        print("The browser will remain open. Close the window when you are done.")
        wait_until_browser_closed(page, context, filled=False)
    finally:
        try:
            if context:
                context.close()
        except Exception:
            pass
        if playwright:
            playwright.stop()

    print("\n========================================")
    print("AFTER EVENT COMPLETE")
    print("========================================")
    print("✓ Actual budget PDF generated and uploaded to Drive")
    print("✓ Activity Report spreadsheet created in the event Drive folder")
    print("✓ Google Form left open for you (not submitted by the script)")


def run_setup():
    print("\n" + "=" * 60)
    print("GOOGLE ACCOUNT SETUP")
    print("=" * 60)

    ensure_dirs()

    playwright = None
    context = None

    try:
        playwright, context, page = launch_browser()

        print("\nOpening Google...")
        page.goto(
            "https://accounts.google.com/",
            wait_until="domcontentloaded",
            timeout=30000,
        )

        page.wait_for_timeout(3000)

        print("\n" + "=" * 60)
        print("LOGIN REQUIRED")
        print("=" * 60)
        print()
        print("In the browser, log into your club Google account")
        print("(for IIIT Delhi student clubs, typically club@sc.iiitd.ac.in).")
        print()
        print("Do NOT use your personal Google account.")
        print()
        print("Complete the entire Google login process.")
        print("Once you can see that you are logged into Google,")
        print("come back here.")
        print("=" * 60)

        input("\nPress ENTER after login is complete...")

        print("\n✓ Browser profile setup complete.")
        print()
        print("The Google login has been saved in:")
        print(f"  {BROWSER_PROFILE}")
        print()
        print("You can now close the browser.")

        input("\nPress ENTER to close the browser...")

    finally:
        if context:
            context.close()

        if playwright:
            playwright.stop()

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Club event automation"
    )

    parser.add_argument(
        "mode",
        choices=["setup", "before", "after"],
        help=(
            "setup = configure Google browser profile, "
            "before = before-event workflow, "
            "after = after-event workflow"
        ),
    )

    parser.add_argument(
        "config",
        type=Path,
        nargs="?",
        help="Path to event YAML config",
    )

    args = parser.parse_args()

    # --------------------------------------------------------
    # SETUP MODE
    # --------------------------------------------------------
    # Does not require a YAML config.
    # Only establishes the persistent Google browser session.
    if args.mode == "setup":
        run_setup()
        return

    # --------------------------------------------------------
    # BEFORE / AFTER MODE
    # --------------------------------------------------------

    if args.config is None:
        parser.error(
            f"the following arguments are required for "
            f"'{args.mode}': config"
        )

    config = load_config(args.config)

    if args.mode == "before":
        run_before(config)

    elif args.mode == "after":
        run_after(config)


if __name__ == "__main__":
    main()
