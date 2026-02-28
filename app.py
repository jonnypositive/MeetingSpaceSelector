#!/usr/bin/env python3
import datetime as dt
import io
import json
import math
import os
import re
import html
from email import policy
from email.parser import BytesParser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import parse_qs, urlparse
from pypdf import PdfReader

ROOT = Path(__file__).resolve().parent
ROOMS_PATH = ROOT / "room_catalog.json"
LOGO_PATH = ROOT / "HotelPolaris_TM_Primary_1c_BLUE.png"
INTERIOR_PATH = ROOT / "Hotel Polaris_Chisholm_2_25_Generations Ballroom-1009.jpg"
LAST_REPORT: Dict[str, Any] = {}

SETUP_TO_CAPACITY_KEY = {
    "classroom": "classroom",
    "theater": "theater",
    "conference": "conference",
    "u-shape": "u_shape",
    "u shape": "u_shape",
    "ushape": "u_shape",
    "hollow": "hollow",
    "reception": "reception",
    "crescent": "banquet_10",
    "rounds": "banquet_10",
    "buffet": "banquet_10",
    "banquet": "banquet_10",
}

DATE_FORMAT = "%a, %b %d, %Y"


def load_rooms() -> List[Dict[str, Any]]:
    rooms = json.loads(ROOMS_PATH.read_text())
    excluded_tokens = {"pre-function", "total indoor space"}
    return [
        r
        for r in rooms
        if not any(token in r["name"].lower() for token in excluded_tokens)
    ]


ROOMS = load_rooms()


def extract_pdf_text(pdf_bytes: bytes) -> str:
    try:
        reader = PdfReader(io.BytesIO(pdf_bytes))
    except Exception as err:
        raise RuntimeError(f"PDF extraction failed: {err}") from err

    pages: List[str] = []
    for idx, page in enumerate(reader.pages, start=1):
        try:
            text = page.extract_text() or ""
        except Exception:
            text = ""
        if text.strip():
            pages.append(f"===PAGE {idx}===\n{text}")
    return "\n".join(pages)


def clean_inline_whitespace(value: str) -> str:
    return " ".join(value.split())


def extract_line_value(text: str, label: str) -> Optional[str]:
    pattern = re.compile(rf"{re.escape(label)}\s+(.+)")
    for line in text.splitlines():
        m = pattern.search(line)
        if m:
            return clean_inline_whitespace(m.group(1))
    return None


def parse_date_range(event_dates: str) -> Tuple[Optional[str], Optional[str]]:
    if not event_dates:
        return None, None

    m = re.search(r"([A-Za-z]{3},\s+[A-Za-z]{3}\s+\d{1,2},\s+\d{4})\s+-\s+([A-Za-z]{3},\s+[A-Za-z]{3}\s+\d{1,2},\s+\d{4})", event_dates)
    if not m:
        return None, None

    try:
        arrival = dt.datetime.strptime(m.group(1), DATE_FORMAT).date().isoformat()
        departure = dt.datetime.strptime(m.group(2), DATE_FORMAT).date().isoformat()
        return arrival, departure
    except ValueError:
        return m.group(1), m.group(2)


def parse_rfp_header(text: str) -> Dict[str, Optional[str]]:
    rfp_name = extract_line_value(text, "RFP Name")
    event_dates = extract_line_value(text, "Event Dates") or ""
    response_due_date = extract_line_value(text, "Response Due Date")
    rfp_type = extract_line_value(text, "RFP Type")
    organization_name = extract_line_value(text, "Organization Name")
    total_room_nights = extract_line_value(text, "Total Room Nights")
    peak_room_nights = extract_line_value(text, "Peak Room Nights")
    arrival, departure = parse_date_range(event_dates)

    key_contact_name = None
    key_contact_organization = None

    contact_line = None
    for line in text.splitlines():
        if "Contact Name" in line:
            contact_line = clean_inline_whitespace(line)
            break

    if contact_line:
        line_match = re.search(r"Contact Name\s+(.+?)\s+Organization\s+(.+?)\s+Address", contact_line)
        if line_match:
            key_contact_name = clean_inline_whitespace(line_match.group(1))
            key_contact_organization = clean_inline_whitespace(line_match.group(2))
        else:
            loose_match = re.search(r"Contact Name\s+(.+?)\s+Email Address", contact_line)
            if loose_match:
                key_contact_name = clean_inline_whitespace(loose_match.group(1))

    if not key_contact_organization:
        org_match = re.search(r"\nOrganization\s+(.+?)\s+Address", text)
        if org_match:
            key_contact_organization = clean_inline_whitespace(org_match.group(1))

    return {
        "rfp_name": rfp_name,
        "event_dates": event_dates,
        "response_due_date": response_due_date,
        "rfp_type": rfp_type,
        "key_contact_name": key_contact_name,
        "key_contact_organization": key_contact_organization,
        "organization_name": organization_name,
        "total_room_nights": total_room_nights,
        "peak_room_nights": peak_room_nights,
        "arrival_date": arrival,
        "departure_date": departure,
    }


def infer_setup_type(raw_setup: str) -> str:
    s = raw_setup.lower()
    for key in SETUP_TO_CAPACITY_KEY:
        if key in s:
            return key
    return "rounds"


def setup_to_capacity_key(setup_type: str) -> str:
    return SETUP_TO_CAPACITY_KEY.get(setup_type.lower(), "banquet_10")


def parse_event_date(date_text: str) -> Optional[str]:
    try:
        return dt.datetime.strptime(date_text, DATE_FORMAT).date().isoformat()
    except ValueError:
        return None


def is_outdoor_season(date_iso: Optional[str]) -> bool:
    if not date_iso:
        return False
    try:
        d = dt.date.fromisoformat(date_iso)
    except ValueError:
        return False
    # Late May through early October.
    return (d.month, d.day) >= (5, 25) and (d.month, d.day) <= (10, 7)


def parse_time_bounds(date_iso: Optional[str], time_range: str) -> Tuple[Optional[dt.datetime], Optional[dt.datetime]]:
    if not date_iso or not time_range:
        return None, None
    parts = [p.strip() for p in time_range.split("-")]
    if len(parts) != 2:
        return None, None
    try:
        base_date = dt.date.fromisoformat(date_iso)
        start_t = dt.datetime.strptime(parts[0], "%I:%M %p").time()
        end_t = dt.datetime.strptime(parts[1], "%I:%M %p").time()
    except ValueError:
        return None, None
    start_dt = dt.datetime.combine(base_date, start_t)
    end_dt = dt.datetime.combine(base_date, end_t)
    if end_dt <= start_dt:
        end_dt = end_dt + dt.timedelta(days=1)
    return start_dt, end_dt


def overlaps(a_start: Optional[dt.datetime], a_end: Optional[dt.datetime], b_start: Optional[dt.datetime], b_end: Optional[dt.datetime]) -> bool:
    if not a_start or not a_end or not b_start or not b_end:
        return False
    return a_start < b_end and b_start < a_end


def detect_room_name(text_line: str) -> Optional[str]:
    line_lower = text_line.lower()
    room_names = sorted((r["name"] for r in ROOMS), key=len, reverse=True)
    for name in room_names:
        if name.lower() in line_lower:
            return name
    return None


def canonicalize_room_name(name: str) -> str:
    raw = clean_inline_whitespace(name or "").lower()
    raw = raw.replace("&amp;", "&")
    raw = raw.replace(" and ", " & ")
    raw = re.sub(r"\s+", " ", raw)
    aliases = {
        "generations b & c": "generations bc",
        "generations c & b": "generations bc",
        "generations a & b": "generations ab",
        "generations b & a": "generations ab",
        "generations a & b & c": "generations ballroom",
        "eagles peak lawn": "eagles peak event lawn",
        "generations ballroom pre-function": "generations ballroom pre-function",
    }
    return aliases.get(raw, raw)


def resolve_room_name(name: str) -> Optional[str]:
    canonical = canonicalize_room_name(name)
    for room in ROOMS:
        if canonical == canonicalize_room_name(room["name"]):
            return room["name"]
    return detect_room_name(name)


def parse_function_diary_html(text: str) -> List[Dict[str, Any]]:
    entries: List[Dict[str, Any]] = []
    rows = re.findall(r"<tr>(.*?)</tr>", text, flags=re.IGNORECASE | re.DOTALL)
    if not rows:
        return entries

    def cell_values(row_html: str, tag: str) -> List[str]:
        vals = re.findall(rf"<{tag}[^>]*>(.*?)</{tag}>", row_html, flags=re.IGNORECASE | re.DOTALL)
        out = []
        for v in vals:
            stripped = re.sub(r"<[^>]+>", "", v)
            out.append(clean_inline_whitespace(html.unescape(stripped)))
        return out

    headers = cell_values(rows[0], "th")
    if not headers:
        headers = cell_values(rows[0], "td")
    idx = {h.lower(): i for i, h in enumerate(headers)}
    needed = [
        "function room",
        "start date",
        "start time 12 hour",
        "end date",
        "end time 12 hour",
        "booking: owner name",
        "booking: booking post as",
    ]
    if not all(col in idx for col in needed):
        return entries

    for row in rows[1:]:
        cols = cell_values(row, "td")
        if not cols:
            continue
        def get(col: str) -> str:
            i = idx[col]
            return cols[i] if i < len(cols) else ""

        room_raw = get("function room")
        room_name = resolve_room_name(room_raw)
        if not room_name:
            continue

        start_date = get("start date")
        start_time = get("start time 12 hour").upper()
        end_date = get("end date")
        end_time = get("end time 12 hour").upper()
        owner = get("booking: owner name") or "Unknown Salesperson"
        booking_post_as = get("booking: booking post as") or "Unknown Group"

        date_iso = None
        date_display = ""
        try:
            d = dt.datetime.strptime(start_date, "%m/%d/%Y").date()
            date_iso = d.isoformat()
            date_display = d.strftime("%a, %b %d, %Y").replace(" 0", " ")
        except Exception:
            pass

        time_range = f"{start_time}-{end_time}" if start_time and end_time else ""
        start_dt = end_dt = None
        try:
            sd = dt.datetime.strptime(start_date, "%m/%d/%Y").date()
            ed = dt.datetime.strptime(end_date, "%m/%d/%Y").date()
            st = dt.datetime.strptime(start_time, "%I:%M %p").time()
            et = dt.datetime.strptime(end_time, "%I:%M %p").time()
            start_dt = dt.datetime.combine(sd, st)
            end_dt = dt.datetime.combine(ed, et)
            if end_dt <= start_dt:
                end_dt += dt.timedelta(days=1)
        except Exception:
            start_dt, end_dt = parse_time_bounds(date_iso, time_range)

        entries.append(
            {
                "room_name": room_name,
                "date_iso": date_iso,
                "date_display": date_display or start_date,
                "time_range": time_range,
                "start_dt": start_dt,
                "end_dt": end_dt,
                "group_name": booking_post_as,
                "salesperson": owner,
            }
        )
    return entries


def parse_function_diary(text: str) -> List[Dict[str, Any]]:
    entries: List[Dict[str, Any]] = []
    current_group = ""
    current_salesperson = ""
    row_re = re.compile(
        r"(Mon|Tue|Wed|Thu|Fri|Sat|Sun),\s+([A-Za-z]{3}\s+\d{1,2},\s+\d{4}).*?(\d{1,2}:\d{2}\s*[AP]M-\d{1,2}:\d{2}\s*[AP]M)",
        re.IGNORECASE,
    )
    group_re = re.compile(r"(Group Name|Group|Event Name|Account Name)\s*[:\-]?\s*(.+)", re.IGNORECASE)
    salesperson_re = re.compile(r"(Salesperson|Sales Manager|Booked By|Catering Sales)\s*[:\-]?\s*(.+)", re.IGNORECASE)

    lines = [clean_inline_whitespace(line) for line in text.splitlines() if clean_inline_whitespace(line)]
    for i, line in enumerate(lines):
        g = group_re.search(line)
        if g:
            current_group = clean_inline_whitespace(g.group(2))
        s = salesperson_re.search(line)
        if s:
            current_salesperson = clean_inline_whitespace(s.group(2))

        m = row_re.search(line)
        if not m:
            continue
        date_display = f"{m.group(1).title()}, {m.group(2)}"
        date_iso = parse_event_date(date_display)
        time_range = m.group(3).upper().replace("  ", " ")
        room_name = resolve_room_name(line)
        if not room_name and i + 1 < len(lines):
            room_name = resolve_room_name(lines[i + 1])
        if not room_name:
            continue

        line_group = current_group
        line_sales = current_salesperson
        g2 = group_re.search(line)
        if g2:
            line_group = clean_inline_whitespace(g2.group(2))
        s2 = salesperson_re.search(line)
        if s2:
            line_sales = clean_inline_whitespace(s2.group(2))

        start_dt, end_dt = parse_time_bounds(date_iso, time_range)
        entries.append(
            {
                "room_name": room_name,
                "date_iso": date_iso,
                "date_display": date_display,
                "time_range": time_range,
                "start_dt": start_dt,
                "end_dt": end_dt,
                "group_name": line_group or "Unknown Group",
                "salesperson": line_sales or "Unknown Salesperson",
            }
        )
    return entries


def parse_diary_upload(diary_bytes: bytes, filename: str = "") -> List[Dict[str, Any]]:
    # This "xls" export is HTML table content with consistent columns.
    raw = diary_bytes.decode("latin-1", errors="ignore")
    lower = raw.lower()
    if "<table" in lower and "booking event: name" in lower and "function room" in lower:
        entries = parse_function_diary_html(raw)
        if entries:
            return entries
    if filename.lower().endswith(".pdf"):
        text = extract_pdf_text(diary_bytes)
        return parse_function_diary(text)
    return parse_function_diary(raw)


def find_room_conflict(room_name: str, requirement: Dict[str, Any], diary_entries: List[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    req_date = requirement.get("event_date_iso")
    req_start, req_end = parse_time_bounds(req_date, requirement.get("time_range", ""))
    for entry in diary_entries:
        if entry.get("room_name", "").lower() != room_name.lower():
            continue
        if req_date and entry.get("date_iso") and req_date != entry.get("date_iso"):
            continue
        if req_start and req_end and entry.get("start_dt") and entry.get("end_dt"):
            if overlaps(req_start, req_end, entry["start_dt"], entry["end_dt"]):
                return entry
        elif req_date == entry.get("date_iso"):
            return entry
    return None


def extract_meeting_requirements_section(text: str) -> str:
    start = text.find("Meeting Room Requirements")
    if start < 0:
        return text
    section = text[start:]
    for end_marker in ["AV Requirements", "Additional Questions"]:
        idx = section.find(end_marker)
        if idx >= 0:
            section = section[:idx]
            break
    return section


def infer_attendees_from_text(raw: str) -> Optional[int]:
    text = raw or ""
    range_match = re.search(r"(\d{1,4})\s*-\s*(\d{1,4})\s*people", text, re.IGNORECASE)
    if range_match:
        return max(int(range_match.group(1)), int(range_match.group(2)))
    single = re.search(r"(\d{1,4})\s*people", text, re.IGNORECASE)
    if single:
        return int(single.group(1))
    return None


def parse_agenda_blocks(text: str) -> List[Dict[str, Any]]:
    section = extract_meeting_requirements_section(text)
    lines = [clean_inline_whitespace(line) for line in section.splitlines() if clean_inline_whitespace(line)]

    row_re = re.compile(
        r"^(Mon|Tue|Wed|Thu|Fri|Sat|Sun),\s+([A-Za-z]{3}\s+\d{1,2},\s+\d{4})\s+(\d{1,2}:\d{2}\s*[AP]M-\d{1,2}:\d{2}\s*[AP]M)\s+(.+)$"
    )
    setup_line_re = re.compile(
        r"^(Crescent rounds|Classroom|Reception|Rounds for 8|Rounds|Buffet|Theater|U-Shape|U Shape|Conference|Hollow)(?:\s*\(Meeting Room Required\))?$",
        re.IGNORECASE,
    )
    people_line_re = re.compile(r"^(\d{1,4})(?:\s*-\s*(\d{1,4}))?\s*people$", re.IGNORECASE)

    blocks: List[Dict[str, Any]] = []
    current: Optional[Dict[str, Any]] = None
    collecting_notes = False

    for line in lines:
        m = row_re.match(line)
        if m:
            if current:
                blocks.append(current)
            weekday, date_text, time_range, agenda_item = m.groups()
            date_display = f"{weekday}, {date_text}"
            current = {
                "date_display": date_display,
                "date_iso": parse_event_date(date_display),
                "time_range": time_range,
                "agenda_item": agenda_item.strip(),
                "setup_requested": "",
                "attendees": None,
                "notes_or_exceptions": "",
            }
            collecting_notes = False
            continue
        if not current:
            continue

        if line.lower().startswith("notes or exceptions:"):
            note = clean_inline_whitespace(line.split(":", 1)[1] if ":" in line else "")
            if note:
                current["notes_or_exceptions"] = (current["notes_or_exceptions"] + " " + note).strip()
            collecting_notes = True
            continue

        setup_match = setup_line_re.match(line.replace("(Meeting Room Required)", "").strip())
        if setup_match:
            current["setup_requested"] = setup_match.group(1)
            collecting_notes = False
            continue

        if "(Meeting Room Required)" in line:
            setup_text = clean_inline_whitespace(line.split("(Meeting Room Required)", 1)[0])
            if setup_text:
                current["setup_requested"] = setup_text
            collecting_notes = False
            continue

        ppl = people_line_re.match(line)
        if ppl:
            hi = ppl.group(2)
            current["attendees"] = int(hi if hi else ppl.group(1))
            collecting_notes = False
            continue

        if collecting_notes:
            if current["notes_or_exceptions"]:
                current["notes_or_exceptions"] += " "
            current["notes_or_exceptions"] += line

    if current:
        blocks.append(current)
    return blocks


def infer_purpose(agenda_item: str, setup_requested: str) -> str:
    base = f"{agenda_item} {setup_requested}".lower()
    if "breakout" in base:
        return "Breakout Session"
    if "breakfast" in base:
        return "Breakfast"
    if "lunch" in base:
        return "Lunch"
    if "reception" in base:
        return "Reception"
    if "dinner" in base:
        return "Dinner"
    return "Meeting"


def parse_meeting_requirements(text: str, header: Optional[Dict[str, Optional[str]]] = None) -> List[Dict[str, Any]]:
    blocks = parse_agenda_blocks(text)
    av_global = "av requirements" in text.lower() or "audio visual" in text.lower() or " a/v " in text.lower()
    total_attendees_text = extract_line_value(text, "Total Attendees") or "0"
    total_attendees = int(re.search(r"\d+", total_attendees_text).group(0)) if re.search(r"\d+", total_attendees_text) else 0

    event_days: Dict[str, int] = {}
    ordered_dates = []
    for row in blocks:
        d = row.get("date_iso")
        if d and d not in event_days:
            ordered_dates.append(d)
            event_days[d] = len(ordered_dates)

    requirements: List[Dict[str, Any]] = []
    for block in blocks:
        setup_requested = block.get("setup_requested", "").strip()
        notes = block.get("notes_or_exceptions", "")
        attendees = block.get("attendees")
        if attendees is None:
            attendees = infer_attendees_from_text(notes or "")
        purpose_preview = infer_purpose(block.get("agenda_item", ""), setup_requested or block.get("agenda_item", ""))
        if attendees is None and purpose_preview in {"Breakfast", "Lunch", "Dinner", "Reception"}:
            attendees = total_attendees if total_attendees > 0 else None
        if attendees is None:
            attendees = infer_attendees_from_text(block.get("agenda_item", ""))
        attendees = int(attendees or 0)
        if attendees <= 0:
            continue
        if not setup_requested:
            if purpose_preview == "Reception":
                setup_requested = "Reception"
            elif purpose_preview in {"Breakfast", "Lunch", "Dinner"}:
                setup_requested = "Rounds"
            elif purpose_preview == "Breakout Session":
                setup_requested = "Classroom"
            else:
                setup_requested = "Classroom"

        purpose = infer_purpose(block.get("agenda_item", ""), setup_requested)
        setup_type = infer_setup_type(setup_requested)
        date_iso = block.get("date_iso") or (header or {}).get("arrival_date")
        av_buffer = 0.0
        if av_global:
            av_buffer = 0.15
        elif purpose in {"Meeting", "Breakout Session", "Dinner"}:
            av_buffer = 0.10
        notes_l = (notes or "").lower()
        if any(token in notes_l for token in ["stage", "band", "dj", "entertainment"]):
            av_buffer = max(av_buffer, 0.20)
        recommended_capacity_need = int(math.ceil(attendees * (1 + av_buffer)))

        day_no = event_days.get(date_iso) if date_iso else None
        requirements.append(
            {
                "purpose": purpose,
                "agenda_item": block.get("agenda_item", purpose),
                "setup_requested": setup_requested,
                "setup_type": setup_type,
                "attendees": attendees,
                "recommended_capacity_need": recommended_capacity_need,
                "capacity_key": setup_to_capacity_key(setup_type),
                "event_date_iso": date_iso,
                "event_date_display": block.get("date_display", date_iso),
                "time_range": block.get("time_range", ""),
                "day_number": day_no,
                "day_label": f"Day {day_no}" if day_no else "Day ?",
                "av_buffer_pct": int(av_buffer * 100),
                "notes_or_exceptions": notes,
            }
        )

    deduped: List[Dict[str, Any]] = []
    seen = set()
    for req in requirements:
        key = (
            req["purpose"],
            req.get("agenda_item", ""),
            req["setup_requested"].lower(),
            req["attendees"],
            req.get("event_date_iso"),
            req.get("time_range"),
        )
        if key in seen:
            continue
        seen.add(key)
        deduped.append(req)
    return deduped


def room_preference_penalty(room_name: str) -> float:
    name = room_name.lower()
    if "constellation" in name:
        return 0.0
    if any(token in name for token in ["santa fe", "homestead", "skyline", "rock", "noctua", "sagitta"]):
        return 16.0
    if "generations" in name:
        return 34.0
    return 20.0


def adjacent_outdoor_option(room_name: str) -> Optional[str]:
    name = room_name.lower()
    if "constellation" in name:
        return "Flat Iron Plaza"
    if "generations" in name or "skyline" in name:
        return "Eagles Peak Event Lawn"
    return None


def appropriateness_label(score: float) -> str:
    if score <= 20:
        return "Excellent"
    if score <= 45:
        return "Strong"
    if score <= 75:
        return "Fair"
    return "Limited"


def rank_rooms(
    requirement: Dict[str, Any],
    rooms: List[Dict[str, Any]],
    diary_entries: Optional[List[Dict[str, Any]]] = None,
    same_rfp_tokens: Optional[List[str]] = None,
) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    attendees = requirement["attendees"]
    needed = requirement.get("recommended_capacity_need", attendees)
    capacity_key = requirement["capacity_key"]
    purpose = requirement.get("purpose", "Meeting")
    outdoor_ok = is_outdoor_season(requirement.get("event_date_iso"))

    candidates = []
    conflicts: List[Dict[str, Any]] = []
    for room in rooms:
        room_name = room["name"]
        cap = room.get(capacity_key)
        if cap is None:
            continue
        if "pre-function" in room_name.lower():
            continue
        if room.get("area") == "Outdoor" and not outdoor_ok:
            continue
        if room.get("area") == "Outdoor" and purpose in {"Meeting", "Breakout Session"}:
            continue
        if cap < needed:
            continue
        if diary_entries:
            conflict = find_room_conflict(room_name, requirement, diary_entries)
            if conflict:
                group_name = conflict.get("group_name", "Unknown Group")
                conflicts.append(
                    {
                        "room_name": room_name,
                        "group_name": group_name,
                        "salesperson": conflict.get("salesperson", "Unknown Salesperson"),
                        "time_range": conflict.get("time_range", ""),
                        "date_display": conflict.get("date_display") or conflict.get("date_iso") or "",
                    }
                )
                continue

        extra = cap - needed
        fit_penalty = (extra / max(needed, 1)) * 100
        score = room_preference_penalty(room_name) + fit_penalty
        candidates.append(
            {
                "room_name": room_name,
                "area": room["area"],
                "capacity_key": capacity_key,
                "capacity": cap,
                "sq_ft": room.get("sq_ft"),
                "extra_capacity": extra,
                "overall_score": round(score, 2),
                "appropriateness_label": appropriateness_label(score),
            }
        )

    candidates.sort(key=lambda c: (c["overall_score"], c["extra_capacity"], c["room_name"]))
    conflicts.sort(key=lambda c: c["room_name"])
    return candidates[:3], conflicts


def add_sequence_suggestions(recommendations: List[Dict[str, Any]]) -> None:
    for i, item in enumerate(recommendations):
        req = item["requirement"]
        purpose = req.get("purpose", "")
        notes = item.setdefault("notes", [])
        if purpose not in {"Breakfast", "Lunch"}:
            continue
        next_item = None
        for j in range(i + 1, len(recommendations)):
            candidate_req = recommendations[j]["requirement"]
            if req.get("event_date_iso") != candidate_req.get("event_date_iso"):
                continue
            if candidate_req.get("purpose") in {"Meeting", "Breakout Session"}:
                next_item = recommendations[j]
                break
        if not next_item:
            continue
        next_req = next_item["requirement"]
        next_top = next_item.get("recommendations", [])
        if next_top:
            notes.append(
                f"Use {next_top[0]['room_name']} for both {purpose.lower()} and the following {next_req.get('purpose', 'meeting').lower()} to reduce room resets."
            )
        if is_outdoor_season(req.get("event_date_iso")) and next_top:
            adjacent = adjacent_outdoor_option(next_top[0]["room_name"])
            if adjacent:
                notes.append(f"Warm-weather meal option: consider adjacent outdoor space at {adjacent}.")


def build_recommendations(
    meeting_requirements: List[Dict[str, Any]],
    diary_entries: Optional[List[Dict[str, Any]]] = None,
    same_rfp_tokens: Optional[List[str]] = None,
) -> List[Dict[str, Any]]:
    output: List[Dict[str, Any]] = []
    for req in meeting_requirements:
        ranked, conflicts = rank_rooms(
            req,
            ROOMS,
            diary_entries=diary_entries,
            same_rfp_tokens=same_rfp_tokens,
        )
        notes: List[str] = []
        if req.get("notes_or_exceptions"):
            notes.append(f"Notes or Exceptions: {req['notes_or_exceptions']}")
        if req.get("av_buffer_pct", 0) > 0:
            notes.append(
                f"Capacity adjusted by {req['av_buffer_pct']}% to account for potential AV/stage footprint."
            )
        if conflicts:
            top_conflicts = conflicts[:3]
            for c in top_conflicts:
                notes.append(
                    f"Conflict: {c['room_name']} is booked by {c['group_name']} (Salesperson: {c['salesperson']}) at {c['date_display']} {c['time_range']}."
                )
        output.append(
            {
                "requirement": req,
                "recommendations": ranked,
                "conflicts": conflicts,
                "notes": notes,
            }
        )
    add_sequence_suggestions(output)
    return output


def calculate_food_beverage(meeting_requirements: List[Dict[str, Any]]) -> Dict[str, Any]:
    rates = {
        "Breakfast": 50,
        "Lunch": 55,
        "Dinner": 120,
        "Reception": 70,
    }
    events: List[Dict[str, Any]] = []
    total = 0
    for req in meeting_requirements:
        purpose = req.get("purpose", "")
        if purpose not in rates:
            continue
        attendees = int(req.get("attendees", 0))
        amount = attendees * rates[purpose]
        total += amount
        events.append(
            {
                "purpose": purpose,
                "attendees": attendees,
                "rate_per_person": rates[purpose],
                "estimated_total": amount,
                "day_label": req.get("day_label", "Day ?"),
                "event_date_display": req.get("event_date_display") or req.get("event_date_iso"),
                "time_range": req.get("time_range", ""),
            }
        )
    return {
        "events": events,
        "total_suggested_fnb_minimum": total,
    }


def looks_like_cvent_rfp(text: str) -> bool:
    required_tokens = [
        "Request for Proposal (RFP)",
        "RFP Details",
        "Meeting Room Requirements",
    ]
    lowered = text.lower()
    return all(token.lower() in lowered for token in required_tokens)


def build_report_lines(report: Dict[str, Any]) -> List[str]:
    def dollars(amount: Any) -> str:
        try:
            return f"${int(amount):,}"
        except Exception:
            return f"${amount}"

    lines: List[str] = []
    header = report.get("header", {})
    lines.append("Hotel Polaris Space Suggester - Recommendation Report")
    lines.append("")
    lines.append(f"RFP Name: {header.get('rfp_name') or '-'}")
    lines.append(f"Event Dates: {header.get('event_dates') or '-'}")
    lines.append(f"Response Due Date: {header.get('response_due_date') or '-'}")
    lines.append(f"RFP Type: {header.get('rfp_type') or '-'}")
    lines.append(f"Key Contact Name: {header.get('key_contact_name') or '-'}")
    lines.append(f"Key Contact Organization: {header.get('key_contact_organization') or '-'}")
    lines.append(f"Organization Name: {header.get('organization_name') or '-'}")
    lines.append(f"Total Room Nights: {header.get('total_room_nights') or '-'}")
    lines.append(f"Peak Room Nights: {header.get('peak_room_nights') or '-'}")
    lines.append("")

    fnb = report.get("food_beverage", {})
    events = fnb.get("events", [])
    if events:
        lines.append("Food And Beverage Events")
        for e in events:
            lines.append(
                f"  {e.get('day_label')} | {e.get('event_date_display')} | {e.get('time_range')} | {e.get('purpose')} | attendees {e.get('attendees')} | {dollars(e.get('estimated_total'))}"
            )
        lines.append(f"Total Suggested Food And Beverage Minimum: {dollars(fnb.get('total_suggested_fnb_minimum', 0))}")
        lines.append("")

    for item in report.get("recommendations", []):
        req = item.get("requirement", {})
        title = f"{req.get('day_label', 'Day ?')} | {req.get('event_date_display') or req.get('event_date_iso') or '-'} | {req.get('time_range') or '-'} | {req.get('purpose') or 'Meeting'}"
        lines.append(title)
        lines.append(
            f"Requested Setup: {req.get('setup_requested', '-')} | Attendees: {req.get('attendees', '-')} | Capacity Target (w/ AV): {req.get('recommended_capacity_need', '-')}"
        )
        for rank, rec in enumerate(item.get("recommendations", []), start=1):
            rank_label = "Best Choice" if rank == 1 else ("2nd Choice" if rank == 2 else ("3rd Choice" if rank == 3 else f"{rank}th Choice"))
            lines.append(
                f"  {rank_label}: {rec.get('room_name')} ({rec.get('area')}) - capacity {rec.get('capacity')}"
            )
        for c in item.get("conflicts", [])[:5]:
            lines.append(
                f"  Conflict: {c.get('room_name')} booked by {c.get('group_name')} (Salesperson: {c.get('salesperson')}) at {c.get('date_display')} {c.get('time_range')}"
            )
        for note in item.get("notes", []):
            lines.append(f"  Note: {note}")
        lines.append("")
    return lines


def render_report_doc(report: Dict[str, Any]) -> bytes:
    lines = build_report_lines(report)
    html_lines = "".join(f"<p>{line}</p>" for line in lines)
    html = f"""<html><head><meta charset="utf-8"></head><body>{html_lines}</body></html>"""
    return html.encode("utf-8")


def render_report_pdf(report: Dict[str, Any]) -> bytes:
    lines = build_report_lines(report)
    max_lines = 48
    pages = [lines[i:i + max_lines] for i in range(0, len(lines), max_lines)] or [[]]

    def esc(s: str) -> str:
        return (
            s.replace("\\", "\\\\")
            .replace("(", "\\(")
            .replace(")", "\\)")
            .encode("latin-1", "replace")
            .decode("latin-1")
        )

    objects: List[bytes] = []
    page_ids: List[int] = []
    content_ids: List[int] = []
    font_id = 3
    next_id = 4
    for _ in pages:
        page_ids.append(next_id)
        content_ids.append(next_id + 1)
        next_id += 2

    objects.append(b"<< /Type /Catalog /Pages 2 0 R >>")
    kids = " ".join(f"{pid} 0 R" for pid in page_ids)
    objects.append(f"<< /Type /Pages /Kids [ {kids} ] /Count {len(page_ids)} >>".encode("ascii"))
    objects.append(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")

    for idx, page_lines in enumerate(pages):
        content = ["BT", "/F1 11 Tf", "50 760 Td", "14 TL"]
        for line in page_lines:
            content.append(f"({esc(line[:120])}) Tj")
            content.append("T*")
        content.append("ET")
        content_stream = "\n".join(content).encode("latin-1", "replace")
        page_obj = (
            f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] "
            f"/Resources << /Font << /F1 {font_id} 0 R >> >> "
            f"/Contents {content_ids[idx]} 0 R >>"
        ).encode("ascii")
        content_obj = b"<< /Length " + str(len(content_stream)).encode("ascii") + b" >>\nstream\n" + content_stream + b"\nendstream"
        objects.append(page_obj)
        objects.append(content_obj)

    pdf = b"%PDF-1.4\n"
    offsets = [0]
    for idx, obj in enumerate(objects, start=1):
        offsets.append(len(pdf))
        pdf += f"{idx} 0 obj\n".encode("ascii") + obj + b"\nendobj\n"
    xref_pos = len(pdf)
    pdf += f"xref\n0 {len(objects) + 1}\n".encode("ascii")
    pdf += b"0000000000 65535 f \n"
    for off in offsets[1:]:
        pdf += f"{off:010d} 00000 n \n".encode("ascii")
    pdf += (
        b"trailer\n<< /Size "
        + str(len(objects) + 1).encode("ascii")
        + b" /Root 1 0 R >>\nstartxref\n"
        + str(xref_pos).encode("ascii")
        + b"\n%%EOF\n"
    )
    return pdf


def html_template() -> str:
    return """<!doctype html>
<html lang=\"en\">
<head>
  <meta charset=\"utf-8\" />
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\" />
  <title>Space Suggester</title>
  <style>
    :root {
      --ink: #0d2236;
      --ink-soft: #2d4359;
      --brand: #1e658f;
      --brand-2: #4f90b7;
      --sand: #eef3f7;
      --paper: #ffffff;
      --glow: #a6c5db;
      --ok: #0a6a58;
      --warn: #8b5e11;
      --radius: 16px;
      --shadow: 0 20px 50px rgba(9, 26, 42, 0.12);
    }

    * { box-sizing: border-box; }

    body {
      margin: 0;
      font-family: "Avenir Next", "Segoe UI", sans-serif;
      color: var(--ink);
      background:
        radial-gradient(circle at 15% 10%, #d6e8f3 0%, transparent 35%),
        radial-gradient(circle at 80% 20%, #dfeef8 0%, transparent 35%),
        linear-gradient(165deg, #f7fbff 0%, #ecf2f7 45%, #f5f8fb 100%);
      min-height: 100vh;
    }

    .hero {
      padding: 2rem 1.25rem 1rem;
      max-width: 1100px;
      margin: 0 auto;
      animation: rise 600ms ease-out;
    }

    .brand-bar {
      display: flex;
      align-items: center;
      gap: 0.9rem;
      margin-bottom: 0.6rem;
    }

    .brand-logo {
      width: 230px;
      max-width: 60vw;
      height: auto;
      display: block;
    }

    h1 {
      margin: 0.7rem 0 0.25rem;
      font-size: clamp(1.8rem, 2.4vw, 2.8rem);
      line-height: 1.15;
    }

    .sub {
      margin: 0;
      color: var(--ink-soft);
      max-width: 60ch;
      font-size: 1.03rem;
    }

    .hero-photo {
      max-width: 1100px;
      margin: 0.75rem auto 0;
      padding: 0 1.25rem;
      animation: rise 700ms ease-out;
    }

    .hero-photo img {
      width: 100%;
      max-height: 260px;
      object-fit: cover;
      border-radius: 18px;
      box-shadow: var(--shadow);
      display: block;
    }

    .grid {
      max-width: 1100px;
      margin: 1rem auto 2rem;
      padding: 0 1.25rem;
      display: grid;
      grid-template-columns: 1.1fr 0.9fr;
      gap: 1rem;
    }

    .panel {
      background: var(--paper);
      border-radius: var(--radius);
      box-shadow: var(--shadow);
      padding: 1rem;
    }

    .panel h2 {
      margin: 0.25rem 0 0.75rem;
      font-size: 1.05rem;
      letter-spacing: 0.02em;
    }

    .field { margin-bottom: 0.8rem; }
    .field label {
      display: block;
      font-size: 0.87rem;
      color: var(--ink-soft);
      margin-bottom: 0.35rem;
      font-weight: 600;
    }

    input[type="text"],
    input[type="password"],
    input[type="url"],
    input[type="file"] {
      width: 100%;
      border: 1px solid #cfdae4;
      border-radius: 10px;
      padding: 0.7rem 0.75rem;
      font-size: 0.95rem;
      background: #fff;
    }

    .row {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 0.7rem;
    }

    button {
      border: 0;
      background: linear-gradient(125deg, var(--brand), #0d4f75);
      color: #fff;
      border-radius: 10px;
      padding: 0.75rem 1rem;
      font-size: 0.95rem;
      font-weight: 700;
      cursor: pointer;
      transition: transform 120ms ease;
    }

    button:hover { transform: translateY(-1px); }
    .export-row {
      display: flex;
      gap: 0.6rem;
      margin-top: 0.75rem;
    }
    .export-row button {
      background: linear-gradient(125deg, #2f566f, #1a3f57);
      padding: 0.55rem 0.8rem;
      font-size: 0.88rem;
    }

    .muted {
      color: var(--ink-soft);
      font-size: 0.87rem;
      margin: 0.35rem 0 0;
    }

    .status {
      margin-top: 0.6rem;
      font-size: 0.9rem;
      color: var(--ok);
      min-height: 1.25rem;
    }

    .warn { color: var(--warn); }

    .results {
      margin-top: 1rem;
      display: grid;
      gap: 0.85rem;
    }

    .card {
      border: 1px solid #d6e1ea;
      border-radius: 12px;
      padding: 0.8rem;
      background: linear-gradient(180deg, #fff, #fbfdff);
    }

    .card h3 {
      margin: 0 0 0.35rem;
      font-size: 0.98rem;
    }

    .chips {
      display: flex;
      gap: 0.4rem;
      flex-wrap: wrap;
      margin-bottom: 0.45rem;
    }

    .chip {
      background: #eaf3fa;
      border: 1px solid #d6e8f6;
      border-radius: 999px;
      padding: 0.2rem 0.55rem;
      font-size: 0.78rem;
      color: #1c4f73;
      font-weight: 600;
    }

    ul {
      margin: 0.4rem 0 0;
      padding-left: 1rem;
    }

    @keyframes rise {
      from { opacity: 0; transform: translateY(8px); }
      to { opacity: 1; transform: translateY(0); }
    }

    @media (max-width: 900px) {
      .grid { grid-template-columns: 1fr; }
      .row { grid-template-columns: 1fr; }
    }
  </style>
</head>
<body>
  <section class=\"hero\">
    <div class=\"brand-bar\">
      <img src=\"/assets/logo\" alt=\"Hotel Polaris\" class=\"brand-logo\" />
    </div>
    <h1>Space Suggester</h1>
    <p class=\"sub\">Upload a Cvent RFP and get best-fit room recommendations based on attendee counts and requested setup style.</p>
  </section>
  <section class=\"hero-photo\">
    <img src=\"/assets/interior\" alt=\"Hotel Polaris Interior\" />
  </section>

  <section class=\"grid\">
    <div class=\"panel\">
      <h2>Cvent RFP Upload</h2>
      <form id=\"rfpForm\">
        <div class=\"field\">
          <label for=\"rfpFile\">RFP PDF</label>
          <input id=\"rfpFile\" name=\"rfp\" type=\"file\" accept=\".pdf\" required />
        </div>
        <div class=\"field\">
          <label for=\"diaryFile\">Tentative/Definite Function Diary (optional)</label>
          <input id=\"diaryFile\" name=\"diary\" type=\"file\" accept=\".pdf,.xls,.html\" />
        </div>
        <button type=\"submit\">Analyze And Recommend Rooms</button>
        <p id=\"status\" class=\"status\"></p>
      </form>
      <div id=\"exports\" class=\"export-row\" style=\"display:none;\">
        <button type=\"button\" id=\"wordBtn\">Download Word</button>
        <button type=\"button\" id=\"pdfBtn\">Download PDF</button>
      </div>
      <div id=\"summary\" class=\"results\"></div>
      <div id=\"recommendations\" class=\"results\"></div>
    </div>

    <div class=\"panel\">
      <h2>Availability + F&B Summary</h2>
      <p class=\"muted\">Upload a Tentative/Definite Function Diary to detect room conflicts and improve recommendation accuracy.</p>
      <hr style=\"border:0;border-top:1px solid #e0e8ef;margin:1rem 0\" />
      <h2>Knowledge Base</h2>
      <p class=\"muted\">Loaded from capacity chart: <strong>33</strong> meeting spaces (indoor + outdoor) with style-specific capacity limits.</p>
      <div id=\"fnbSummary\" class=\"results\"></div>
    </div>
  </section>

  <script>
    const form = document.getElementById("rfpForm");
    const statusEl = document.getElementById("status");
    const summaryEl = document.getElementById("summary");
    const recsEl = document.getElementById("recommendations");
    const fnbEl = document.getElementById("fnbSummary");
    const exportsEl = document.getElementById("exports");
    const wordBtn = document.getElementById("wordBtn");
    const pdfBtn = document.getElementById("pdfBtn");

    function escapeHtml(value) {
      return String(value)
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
    }

    function renderHeader(data) {
      const header = data.header || {};
      const rows = [
        ["RFP Name", header.rfp_name || "-"],
        ["Event Dates", header.event_dates || "-"],
        ["Response Due Date", header.response_due_date || "-"],
        ["RFP Type", header.rfp_type || "-"],
        ["Key Contact Name", header.key_contact_name || "-"],
        ["Key Contact Organization", header.key_contact_organization || "-"],
        ["Organization Name", header.organization_name || "-"],
        ["Total Room Nights", header.total_room_nights || "-"],
        ["Peak Room Nights", header.peak_room_nights || "-"],
      ];

      summaryEl.innerHTML = `
        <div class=\"card\">
          <h3>RFP Details</h3>
          <ul>
            ${rows.map(([k, v]) => `<li><strong>${escapeHtml(k)}:</strong> ${escapeHtml(v)}</li>`).join("")}
          </ul>
          <p class=\"muted\"><strong>Diary rows parsed:</strong> ${escapeHtml(data.diary_entries_parsed || 0)}</p>
        </div>
      `;
    }

    function renderFnb(data) {
      const fnb = data.food_beverage || {};
      const events = fnb.events || [];
      const usd = new Intl.NumberFormat("en-US", { style: "currency", currency: "USD", maximumFractionDigits: 0 });
      if (!events.length) {
        fnbEl.innerHTML = `<div class=\"card\"><h3>Food And Beverage Events</h3><p class=\"muted\">No breakfast/lunch/dinner/reception events detected.</p></div>`;
        return;
      }
      fnbEl.innerHTML = `
        <div class=\"card\">
          <h3>Food And Beverage Events</h3>
          <ul>
            ${events.map(e => `<li><strong>${escapeHtml(e.day_label || "")} ${escapeHtml(e.purpose || "")}</strong> (${escapeHtml(e.event_date_display || "-")} ${escapeHtml(e.time_range || "")}) · ${escapeHtml(e.attendees)} attendees · ${escapeHtml(usd.format(Number(e.estimated_total || 0)))}</li>`).join("")}
          </ul>
          <p><strong>Total Suggested Food And Beverage Minimum: ${escapeHtml(usd.format(Number(fnb.total_suggested_fnb_minimum || 0)))}</strong></p>
        </div>
      `;
    }

    function renderRecommendations(data) {
      const items = data.recommendations || [];
      if (!items.length) {
        recsEl.innerHTML = `<div class=\"card\"><h3>No meeting-room requirements detected</h3><p class=\"muted\">Try another Cvent RFP sample with explicit \"(Meeting Room Required)\" lines.</p></div>`;
        return;
      }

      recsEl.innerHTML = items.map((item) => {
        const req = item.requirement;
        const recs = item.recommendations || [];
        const title = `${req.day_label || "Day ?"} · ${req.event_date_display || req.event_date_iso || "-"} · ${req.time_range || "-"} · ${req.purpose || "Meeting"}`;
        const chips = `
          <div class=\"chips\">
            <span class=\"chip\">Setup: ${escapeHtml(req.setup_requested)}</span>
            <span class=\"chip\">Attendees: ${escapeHtml(req.attendees)}</span>
            <span class=\"chip\">Capacity Target: ${escapeHtml(req.recommended_capacity_need || req.attendees)}</span>
          </div>
        `;

        const rankLabel = (idx) => idx === 0 ? "Best Choice" : idx === 1 ? "2nd Choice" : idx === 2 ? "3rd Choice" : `${idx + 1}th Choice`;
        const list = recs.length
          ? `<ul>${recs.map((r, i) => `<li><strong>${escapeHtml(rankLabel(i))}:</strong> <strong>${escapeHtml(r.room_name)}</strong> (${escapeHtml(r.area)}) · cap ${escapeHtml(r.capacity)}</li>`).join("")}</ul>`
          : `<p class=\"muted warn\">No room in the current catalog can hold ${escapeHtml(req.attendees)} for this setup style.</p>`;

        const notes = (item.notes || []).length
          ? `<ul>${item.notes.map(n => `<li>${escapeHtml(n)}</li>`).join("")}</ul>`
          : "";
        const conflicts = (item.conflicts || []).length
          ? `<ul>${item.conflicts.slice(0, 4).map(c => `<li><strong>Conflict:</strong> ${escapeHtml(c.room_name)} booked by ${escapeHtml(c.group_name)} (Salesperson: ${escapeHtml(c.salesperson)}) at ${escapeHtml(c.date_display || "")} ${escapeHtml(c.time_range || "")}</li>`).join("")}</ul>`
          : `<p class=\"muted\">No diary conflicts detected for this request.</p>`;
        return `<div class=\"card\"><h3>${escapeHtml(title)}</h3>${chips}${list}${conflicts}${notes}</div>`;
      }).join("");
    }

    form.addEventListener("submit", async (event) => {
      event.preventDefault();
      statusEl.textContent = "Analyzing RFP...";
      statusEl.classList.remove("warn");
      summaryEl.innerHTML = "";
      recsEl.innerHTML = "";
      fnbEl.innerHTML = "";
      exportsEl.style.display = "none";

      const file = document.getElementById("rfpFile").files[0];
      if (!file) {
        statusEl.textContent = "Select a PDF first.";
        statusEl.classList.add("warn");
        return;
      }

      const formData = new FormData();
      formData.append("rfp", file);
      const diaryFile = document.getElementById("diaryFile").files[0];
      if (diaryFile) {
        formData.append("diary", diaryFile);
      }

      try {
        const res = await fetch("/api/parse-rfp", {
          method: "POST",
          body: formData,
        });

        const data = await res.json();
        if (!res.ok) {
          throw new Error(data.error || "Upload failed");
        }

        statusEl.textContent = "Done. Recommendations generated.";
        renderHeader(data);
        renderFnb(data);
        renderRecommendations(data);
        exportsEl.style.display = "flex";
      } catch (err) {
        statusEl.textContent = err.message || "Could not parse this file.";
        statusEl.classList.add("warn");
      }
    });

    wordBtn.addEventListener("click", () => {
      window.location.href = "/api/export?format=word";
    });
    pdfBtn.addEventListener("click", () => {
      window.location.href = "/api/export?format=pdf";
    });
  </script>
</body>
</html>
"""


class AppHandler(BaseHTTPRequestHandler):
    def _multipart_fields(self) -> Dict[str, Dict[str, Any]]:
        content_type = self.headers.get("Content-Type", "")
        content_length = int(self.headers.get("Content-Length", "0") or "0")
        if "multipart/form-data" not in content_type or content_length <= 0:
            return {}

        body = self.rfile.read(content_length)
        raw = (
            f"Content-Type: {content_type}\r\n"
            "MIME-Version: 1.0\r\n"
            "\r\n"
        ).encode("utf-8") + body
        message = BytesParser(policy=policy.default).parsebytes(raw)
        if not message.is_multipart():
            return {}

        fields: Dict[str, Dict[str, Any]] = {}
        for part in message.iter_parts():
            name = part.get_param("name", header="content-disposition")
            if not name:
                continue
            fields[name] = {
                "filename": part.get_filename() or "",
                "data": part.get_payload(decode=True) or b"",
            }
        return fields

    def _file(self, path: Path, content_type: str) -> None:
        if not path.exists() or not path.is_file():
            self._json({"error": "Not found"}, status=404)
            return
        body = path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _json(self, payload: Dict[str, Any], status: int = 200) -> None:
        body = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _binary(self, body: bytes, content_type: str, filename: str) -> None:
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/assets/logo":
            self._file(LOGO_PATH, "image/png")
            return

        if parsed.path == "/assets/interior":
            self._file(INTERIOR_PATH, "image/jpeg")
            return

        if parsed.path == "/":
            page = html_template().encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(page)))
            self.end_headers()
            self.wfile.write(page)
            return

        if parsed.path == "/api/export":
            if not LAST_REPORT:
                self._json({"error": "No analysis available yet. Upload an RFP first."}, status=400)
                return
            fmt = parse_qs(parsed.query).get("format", ["word"])[0].lower()
            if fmt == "word":
                body = render_report_doc(LAST_REPORT)
                self._binary(body, "application/msword", "hotel-polaris-recommendations.doc")
                return
            if fmt == "pdf":
                body = render_report_pdf(LAST_REPORT)
                self._binary(body, "application/pdf", "hotel-polaris-recommendations.pdf")
                return
            self._json({"error": "Unsupported export format"}, status=400)
            return

        if parsed.path == "/api/rooms":
            self._json({"rooms": ROOMS})
            return

        self._json({"error": "Not found"}, status=404)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path != "/api/parse-rfp":
            self._json({"error": "Not found"}, status=404)
            return

        if "multipart/form-data" not in (self.headers.get("Content-Type", "")):
            self._json({"error": "Expected multipart/form-data"}, status=400)
            return

        fields = self._multipart_fields()
        if not fields:
            self._json({"error": "Invalid multipart/form-data payload"}, status=400)
            return

        filefield = fields.get("rfp")
        if filefield is None:
            self._json({"error": "Missing RFP file upload"}, status=400)
            return
        diaryfield = fields.get("diary")

        pdf_bytes = filefield.get("data", b"")
        if not pdf_bytes:
            self._json({"error": "Uploaded file was empty"}, status=400)
            return

        try:
            text = extract_pdf_text(pdf_bytes)
        except RuntimeError as err:
            self._json({"error": str(err)}, status=500)
            return

        if not looks_like_cvent_rfp(text):
            self._json(
                {
                    "error": "This does not look like a Cvent RFP format. Phase 1 is limited to Cvent templates.",
                },
                status=400,
            )
            return

        diary_entries: List[Dict[str, Any]] = []
        if diaryfield is not None:
            diary_bytes = diaryfield.get("data", b"")
            if diary_bytes:
                try:
                    diary_entries = parse_diary_upload(diary_bytes, diaryfield.get("filename", ""))
                except RuntimeError:
                    diary_entries = []

        header = parse_rfp_header(text)
        requirements = parse_meeting_requirements(text, header=header)
        same_rfp_tokens = [
            (header.get("rfp_name") or "").lower(),
            (header.get("organization_name") or "").lower(),
        ]
        recommendations = build_recommendations(
            requirements,
            diary_entries=diary_entries,
            same_rfp_tokens=same_rfp_tokens,
        )
        food_beverage = calculate_food_beverage(requirements)
        global LAST_REPORT
        LAST_REPORT = {
            "header": header,
            "requirements_count": len(requirements),
            "recommendations": recommendations,
            "food_beverage": food_beverage,
            "diary_entries_parsed": len(diary_entries),
        }

        self._json(
            LAST_REPORT
        )


def run() -> None:
    host = os.getenv("HOST", "0.0.0.0")
    port = int(os.getenv("PORT", "8080"))
    server = ThreadingHTTPServer((host, port), AppHandler)
    print(f"Hotel Polaris Space Suggester running at http://{host}:{port}")
    server.serve_forever()


if __name__ == "__main__":
    run()
