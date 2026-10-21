"""
Meeting Briefs

Hourly job that:
  1. Looks at the next 26 hours of calendar events
  2. Finds outside attendees (anyone not on an internal domain) who haven't been briefed yet
  3. Labels each guest (client / contractor / contact) to shape the brief. Every new
     outside guest gets a brief; the label never blocks one.
  4. Counts their past meetings with the firm, checks Notion for an existing record
     (which turns the brief into a refresher), and researches them on the web
  5. Saves a Google Doc brief titled "Meeting Brief | <name or email> | <date, time>"
  6. Emails the doc link to the bosses who are in that meeting

Because it runs every hour over a 26 hour window, normal meetings get briefed about
a day ahead and last-minute bookings get picked up within the hour. Postgres tracks
who's been briefed so nobody gets researched twice.

Setup:
  pip install -r requirements.txt
  python meeting_briefs.py --auth        # one time, prints GOOGLE_REFRESH_TOKEN
  python meeting_briefs.py --check       # test every connection and permission, writes nothing
  python meeting_briefs.py --dry-run     # test: logs what it would do, writes nothing
  python meeting_briefs.py               # real run (schedule hourly: 0 * * * *)

Env vars:
  DATABASE_URL           Postgres connection string
  OPENROUTER_API_KEY
  LLM_MODEL              OpenRouter model slug, e.g. openai/gpt-4o-mini
  GOOGLE_CLIENT_ID
  GOOGLE_CLIENT_SECRET
  GOOGLE_REFRESH_TOKEN
  CALENDAR_IDS           comma list of calendars to watch (default: primary)
  INTERNAL_DOMAINS       comma list, e.g. aletapartners.com (these people are never researched)
  INTERNAL_EMAILS        comma list of specific addresses never researched (team's personal gmails,
                         existing contractors). BOSS_EMAILS are always treated as internal too.
  REBRIEF_DAYS           don't re-research someone briefed within this many days (default 180)
  NEW_CONTACT_MAX_MEETINGS  only brief people with this many or fewer meetings with the firm in the
                         last 90 days (default 2; 0 = first meetings only)
  MAX_OUTSIDE_ATTENDEES  skip events with more outside guests than this (default 8, webinars/group invites)
  DEBUG_REASONS          set to 1 to log why each guest got the label they did (testing)
  BOSS_EMAILS            comma list of who gets the briefs
  BOSS_FOLDERS           each boss's own Drive folder, e.g.
                         sarah@aletapartners.com:<folder-id>,alina@aletapartners.com:<folder-id>
  BRIEFS_FOLDER_ID       fallback Drive folder for any boss not listed in BOSS_FOLDERS
  NOTION_API_KEY         optional: Notion integration secret (same as the Notion pipeline)
  NOTION_DATABASE_ID     optional: Notion database of people the firm already knows
  BOSS_TZ                default America/Los_Angeles
  LOOKAHEAD_HOURS        default 26

Requires Python 3.11+.
"""

import argparse
import base64
import json
import logging
import os
import re
from datetime import datetime, timedelta, timezone
from email.mime.text import MIMEText
from html import escape
from zoneinfo import ZoneInfo

import psycopg
import requests
from dotenv import load_dotenv
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaInMemoryUpload
from openai import OpenAI

# Load .env from the script's own folder. Real env vars (like Railway's) win if both exist,
# and on Railway there's no .env file, so this quietly does nothing there.
load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("meeting_briefs")


# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

def env_list(name, default=""):
    return [x.strip() for x in os.environ.get(name, default).split(",") if x.strip()]

SCOPES = [
    "https://www.googleapis.com/auth/calendar.readonly",
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/drive",
]

CALENDAR_IDS = env_list("CALENDAR_IDS", "primary")
INTERNAL_DOMAINS = {d.lower() for d in env_list("INTERNAL_DOMAINS")}
INTERNAL_EMAILS = {e.lower() for e in env_list("INTERNAL_EMAILS")}
REBRIEF_DAYS = int(os.environ.get("REBRIEF_DAYS", "180"))
BOSS_EMAILS = env_list("BOSS_EMAILS")
BRIEFS_FOLDER_ID = os.environ.get("BRIEFS_FOLDER_ID", "")
# boss email -> their Drive folder ID
BOSS_FOLDERS = {
    k.strip().lower(): v.strip()
    for k, v in (pair.split(":", 1) for pair in env_list("BOSS_FOLDERS"))
}
NOTION_API_KEY = os.environ.get("NOTION_API_KEY", "")
NOTION_DATABASE_ID = os.environ.get("NOTION_DATABASE_ID", "")
BOSS_TZ = ZoneInfo(os.environ.get("BOSS_TZ", "America/Los_Angeles"))
LOOKAHEAD = timedelta(hours=int(os.environ.get("LOOKAHEAD_HOURS", "26")))
LLM_MODEL = os.environ.get("LLM_MODEL", "openai/gpt-4o-mini")
DEBUG_REASONS = os.environ.get("DEBUG_REASONS", "").lower() in {"1", "true", "yes"}

PERSONAL_DOMAINS = {
    "gmail.com", "googlemail.com", "yahoo.com", "outlook.com", "hotmail.com",
    "live.com", "icloud.com", "me.com", "aol.com", "proton.me", "protonmail.com",
}

KIND_LABEL = {"client": "Potential Client", "contractor": "Potential Contractor", "contact": "Contact"}
MAX_OUTSIDE_ATTENDEES = int(os.environ.get("MAX_OUTSIDE_ATTENDEES", "8"))
HISTORY_DAYS = 90
NEW_CONTACT_MAX_MEETINGS = int(os.environ.get("NEW_CONTACT_MAX_MEETINGS", "2"))

# Firm shorthand the bosses use. Add more lines anytime: "ABBR": "what it means"
GLOSSARY = {
    "MBB": "McKinsey & Company, Bain & Company, and Boston Consulting Group (BCG), the top three strategy consulting firms",
    "Big 4": "Deloitte, PwC, EY, and KPMG, the four largest accounting and professional services firms",
}
GLOSSARY_TEXT = "\n".join(f"- {k}: {v}" for k, v in GLOSSARY.items())


# ---------------------------------------------------------------------------
# Google auth
# ---------------------------------------------------------------------------

def print_refresh_token():
    """One-time: opens a browser to sign in, then prints the refresh token to store as an env var."""
    from google_auth_oauthlib.flow import InstalledAppFlow

    flow = InstalledAppFlow.from_client_config(
        {
            "installed": {
                "client_id": os.environ["GOOGLE_CLIENT_ID"],
                "client_secret": os.environ["GOOGLE_CLIENT_SECRET"],
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
            }
        },
        SCOPES,
    )
    creds = flow.run_local_server(port=0, access_type="offline", prompt="consent")
    print("\nGOOGLE_REFRESH_TOKEN=" + creds.refresh_token)


def google_services():
    creds = Credentials(
        token=None,
        refresh_token=os.environ["GOOGLE_REFRESH_TOKEN"],
        token_uri="https://oauth2.googleapis.com/token",
        client_id=os.environ["GOOGLE_CLIENT_ID"],
        client_secret=os.environ["GOOGLE_CLIENT_SECRET"],
        scopes=SCOPES,
    )
    creds.refresh(Request())
    calendar = build("calendar", "v3", credentials=creds, cache_discovery=False)
    gmail = build("gmail", "v1", credentials=creds, cache_discovery=False)
    drive = build("drive", "v3", credentials=creds, cache_discovery=False)
    return calendar, gmail, drive


# ---------------------------------------------------------------------------
# Database (tracks who's already been briefed)
# ---------------------------------------------------------------------------

DDL = """
CREATE TABLE IF NOT EXISTS briefed_attendees (
    meeting_key  TEXT NOT NULL,
    email        TEXT NOT NULL,
    kind         TEXT,
    doc_url      TEXT,
    processed_at TIMESTAMPTZ NOT NULL DEFAULT now(),
    PRIMARY KEY (meeting_key, email)
)
"""

def already_processed(conn, key, email):
    row = conn.execute(
        "SELECT 1 FROM briefed_attendees WHERE meeting_key = %s AND email = %s",
        (key, email),
    ).fetchone()
    return row is not None


def recently_briefed(conn, email):
    """True if this person already got a brief in another meeting recently,
    so recurring meetings and follow-ups don't re-research the same people."""
    row = conn.execute(
        """SELECT 1 FROM briefed_attendees
           WHERE email = %s AND doc_url IS NOT NULL
             AND processed_at > now() - make_interval(days => %s)""",
        (email, REBRIEF_DAYS),
    ).fetchone()
    return row is not None


def mark_processed(conn, key, email, kind, doc_url=None):
    conn.execute(
        """INSERT INTO briefed_attendees (meeting_key, email, kind, doc_url)
           VALUES (%s, %s, %s, %s) ON CONFLICT DO NOTHING""",
        (key, email, kind, doc_url),
    )


# ---------------------------------------------------------------------------
# Notion (people the firm already knows -> refresher briefs)
# ---------------------------------------------------------------------------

def _notion_prop_text(prop):
    """Turns any Notion property into plain text."""
    t = prop.get("type")
    v = prop.get(t)
    if v is None:
        return ""
    if t in ("title", "rich_text"):
        return "".join(x.get("plain_text", "") for x in v)
    if t in ("select", "status"):
        return v.get("name", "")
    if t == "multi_select":
        return ", ".join(x.get("name", "") for x in v)
    if t == "date":
        return v.get("start", "") or ""
    if t == "people":
        return ", ".join(p.get("name", "") or "" for p in v)
    if t == "formula":
        return str(v.get(v.get("type")) or "")
    if t in ("email", "phone_number", "url", "number", "checkbox"):
        return str(v)
    return ""


def load_notion_rows():
    """Reads the whole Notion database once per run. Returns [] if Notion isn't set up."""
    if not (NOTION_API_KEY and NOTION_DATABASE_ID):
        return []
    headers = {
        "Authorization": f"Bearer {NOTION_API_KEY}",
        "Notion-Version": "2022-06-28",
        "Content-Type": "application/json",
    }
    rows, cursor = [], None
    try:
        while True:
            body = {"page_size": 100}
            if cursor:
                body["start_cursor"] = cursor
            r = requests.post(
                f"https://api.notion.com/v1/databases/{NOTION_DATABASE_ID}/query",
                headers=headers, json=body, timeout=30,
            )
            r.raise_for_status()
            data = r.json()
            for page in data.get("results", []):
                props = page.get("properties", {})
                title = next((_notion_prop_text(p) for p in props.values() if p.get("type") == "title"), "")
                fields = {k: _notion_prop_text(p) for k, p in props.items()}
                rows.append({"title": title, "fields": {k: v for k, v in fields.items() if v}})
            if not data.get("has_more"):
                break
            cursor = data.get("next_cursor")
    except Exception:
        log.warning("Couldn't read Notion, briefs will go out without Notion info", exc_info=True)
        return []
    log.info("Loaded %d rows from Notion", len(rows))
    return rows


def notion_match(rows, email, name):
    """Finds this person in Notion: first by email in any column, then by exact full name."""
    email = email.lower()
    for r in rows:
        if any(email in v.lower() for v in r["fields"].values()):
            return r
    if name and "@" not in name:
        n = name.lower().strip()
        for r in rows:
            if r["title"].lower().strip() == n:
                return r
    return None


def notion_text(row):
    return "\n".join(f"{k}: {v}" for k, v in row["fields"].items()) if row else ""


# ---------------------------------------------------------------------------
# Calendar
# ---------------------------------------------------------------------------

def meeting_key(ev):
    """Stable ID for a meeting. Same across calendars, survives reschedules,
    and stays distinct per occurrence for recurring meetings."""
    uid = ev.get("iCalUID", ev["id"])
    original = ev.get("originalStartTime", {}).get("dateTime")
    return f"{uid}|{original}" if original else uid


def upcoming_events(calendar, now):
    seen, events = {}, []
    for cal_id in CALENDAR_IDS:
        page_token = None
        while True:
            resp = calendar.events().list(
                calendarId=cal_id,
                timeMin=now.isoformat(),
                timeMax=(now + LOOKAHEAD).isoformat(),
                singleEvents=True,
                orderBy="startTime",
                maxResults=250,
                pageToken=page_token,
            ).execute()
            for ev in resp.get("items", []):
                if ev.get("status") == "cancelled":
                    continue
                if "dateTime" not in ev.get("start", {}):  # all-day events
                    continue
                if ev.get("eventType", "default") != "default":  # focus time, out of office, etc.
                    continue
                key = meeting_key(ev)
                if key in seen:  # same meeting on two bosses' calendars
                    seen[key]["_calendars"].append(cal_id.lower())
                    continue
                ev["_calendars"] = [cal_id.lower()]
                seen[key] = ev
                events.append(ev)
            page_token = resp.get("nextPageToken")
            if not page_token:
                break
    return events


def is_internal(email):
    email = email.lower()
    if email in INTERNAL_EMAILS or email in {b.lower() for b in BOSS_EMAILS}:
        return True
    domain = email.split("@")[-1]
    return any(domain == d or domain.endswith("." + d) for d in INTERNAL_DOMAINS)


def outside_attendees(ev):
    people = []
    for a in ev.get("attendees", []):
        email = a.get("email", "").lower()
        if not email or a.get("resource"):           # meeting rooms
            continue
        if a.get("responseStatus") == "declined":
            continue
        if email.endswith("calendar.google.com"):    # group/resource calendars
            continue
        if is_internal(email):                       # never research our own people
            continue
        people.append({**a, "email": email})
    return people


def meeting_when(ev):
    start = datetime.fromisoformat(ev["start"]["dateTime"]).astimezone(BOSS_TZ)
    return start.strftime("%a %b %-d, %Y, %-I:%M %p %Z")


def bosses_for(ev):
    """Bosses who are in this meeting: invited, organizing, or it's on their calendar."""
    involved = {a.get("email", "").lower() for a in ev.get("attendees", [])}
    involved.add(ev.get("organizer", {}).get("email", "").lower())
    involved.update(ev.get("_calendars", []))
    return [b for b in BOSS_EMAILS if b.lower() in involved] or BOSS_EMAILS


# ---------------------------------------------------------------------------
# Gmail context (helps classify the meeting and find names)
# ---------------------------------------------------------------------------

def email_context(gmail, email, limit=5):
    try:
        resp = gmail.users().messages().list(
            userId="me", q=f"from:{email} OR to:{email} newer_than:90d", maxResults=limit
        ).execute()
    except Exception:
        log.warning("Gmail search failed for %s", email)
        return ""

    lines = []
    for m in resp.get("messages", []):
        msg = gmail.users().messages().get(
            userId="me", id=m["id"], format="metadata", metadataHeaders=["From", "Subject"]
        ).execute()
        headers = {h["name"]: h["value"] for h in msg.get("payload", {}).get("headers", [])}
        lines.append(
            f"From: {headers.get('From', '')} | Subject: {headers.get('Subject', '')} | {msg.get('snippet', '')}"
        )
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# LLM
# ---------------------------------------------------------------------------

_llm = None

def llm():
    global _llm
    if _llm is None:
        _llm = OpenAI(base_url="https://openrouter.ai/api/v1", api_key=os.environ["OPENROUTER_API_KEY"])
    return _llm


def ask(model, prompt, json_mode=False):
    kwargs = {"response_format": {"type": "json_object"}} if json_mode else {}
    resp = llm().chat.completions.create(
        model=model,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.2,
        **kwargs,
    )
    return resp.choices[0].message.content or ""


def strip_fences(text):
    return re.sub(r"^```[a-zA-Z]*\s*|\s*```$", "", text.strip())


CLASSIFY_PROMPT = """You label outside guests for a small consulting firm's meeting briefs. The firm takes on clients and hands the work to contractors, and it also brings on contractors as consultants. Every guest gets a brief no matter what you answer; your label only decides which sections the brief includes.

Meeting title: {title}
Meeting description: {description}
Outside attendees: {emails}

Recent emails with these people:
{context}

Reply with JSON only, in this shape:
{{"kind": "client" | "contractor" | "contact", "reason": "<one short sentence on why>", "names": {{"<email>": "<full name or null>"}}}}

- "client": they look like a client or prospective client (company buying services, kickoff, discovery, pitch)
- "contractor": they look like a consultant/contractor or someone being considered as one
- "contact": anything else, or you can't tell (this is a fine answer)

Only give a name if it actually appears in the title, description, or emails (for example an email signature or a From header). Never guess a name."""


def classify(ev, people, contexts):
    """Labels the meeting for the brief. Never blocks a brief: any failure falls back to 'contact'."""
    try:
        prompt = CLASSIFY_PROMPT.format(
            title=ev.get("summary", "(no title)"),
            description=(ev.get("description") or "(none)")[:3000],
            emails=", ".join(p["email"] for p in people),
            context="\n\n".join(f"[{e}]\n{c or '(no emails found)'}" for e, c in contexts.items()),
        )
        data = json.loads(strip_fences(ask(LLM_MODEL, prompt, json_mode=True)))
    except Exception:
        log.warning("Labeling failed for %r, using 'contact'", ev.get("summary"))
        return "contact", {}
    kind = data.get("kind") if data.get("kind") in KIND_LABEL else "contact"
    names = {k.lower(): v for k, v in (data.get("names") or {}).items() if v}
    if DEBUG_REASONS:
        log.info("Labeled %r as %s: %s", ev.get("summary"), kind, data.get("reason", "(no reason given)"))
    return kind, names


def past_meetings(calendar, email, now):
    """How many meetings this person had with the firm in the last HISTORY_DAYS (code, not AI)."""
    seen = set()
    for cal_id in CALENDAR_IDS:
        try:
            resp = calendar.events().list(
                calendarId=cal_id, q=email, singleEvents=True, maxResults=250,
                timeMin=(now - timedelta(days=HISTORY_DAYS)).isoformat(), timeMax=now.isoformat(),
            ).execute()
        except Exception:
            continue
        for ev in resp.get("items", []):
            emails = {a.get("email", "").lower() for a in ev.get("attendees", [])}
            if email in emails and ev.get("status") != "cancelled":
                seen.add(meeting_key(ev))
    return len(seen)


def history_line(count):
    if count == 0:
        return f"First meeting with the firm (no meetings in the last {HISTORY_DAYS} days)"
    return f"Met with the firm {count} time{'s' if count != 1 else ''} in the last {HISTORY_DAYS} days"


RESEARCH_PROMPT = """Research this person ahead of a meeting with a small consulting firm.

Name: {name}
Email: {email}
Company website (from their email domain): {site}
Meeting type: {kind_label}
Meeting title: {title}
History with the firm: {history}

Context from recent emails:
{context}

What the firm already has on file about them (from its Notion database):
{notion}
If there is something on file, this person is already known to the firm and this brief is a REFRESHER. In that case, start the Snapshot with one line on how they're connected to the firm (based only on the Notion info), and focus on what has changed or is new since then (new roles, companies, launches, news). Talking points should build on the existing relationship.

Search the web for their professional background. Try "{name} {company_hint} LinkedIn", their company website, articles, talks, GitHub, and similar public sources.

Firm shorthand (the partners use these terms, so use them in the brief too):
{glossary}
If the person has worked at any of these firms, flag it near the top of the Snapshot using the shorthand, e.g. "Ex-MBB (McKinsey, 2015-2019)" or "Big 4 background (Deloitte)".

Rules:
- Only include facts you found in a source, with the source URL linked right next to each fact. Facts from the Notion info can be cited as "(firm records)".
- If you can't confirm something, list it under "Not found" instead of guessing.
- If results might belong to a different person with the same name, say so clearly.
- Professional info only: roles, employers, skills, education, public work. Nothing personal.

Output an HTML fragment only (no <html>/<body> tags, no markdown, no code fences).
The first element must be: <p><b>Confidence this is the right person:</b> High, Medium, or Low, plus one short reason</p>
Then these sections, each with an <h2> heading:
Snapshot (2-3 sentences: who they are and their current role)
Background (past roles, education, skills)
{third_section}
Talking points (3-5 suggested questions for the meeting)
Not found
Sources (list of every URL used)"""


def research(name, email, kind, ev, context, history, notion=""):
    domain = email.split("@")[-1]
    personal = domain in PERSONAL_DOMAINS
    prompt = RESEARCH_PROMPT.format(
        name=name if name != email else "(unknown, search using the email and domain)",
        email=email,
        site="none (personal email address)" if personal else domain,
        kind_label=KIND_LABEL[kind],
        title=ev.get("summary", "(no title)"),
        context=context or "(none)",
        company_hint="" if personal else domain.split(".")[0],
        history=history,
        notion=notion or "(nothing on file, treat them as new)",
        glossary=GLOSSARY_TEXT,
        third_section={
            "client": "Company (what their company does, size, industry, likely needs)",
            "contractor": "Skills and fit (what they could bring to the firm as a contractor)",
        }.get(kind, "Organization (what their company or organization does, and their role in it)"),
    )
    return strip_fences(ask(f"{LLM_MODEL}:online", prompt))  # :online = OpenRouter web search


# ---------------------------------------------------------------------------
# Output: Google Doc + email
# ---------------------------------------------------------------------------

def build_doc_html(name, email, kind, ev, when, body, history, notion_row=None):
    on_file = ""
    if notion_row:
        items = "".join(
            f"<li><b>{escape(k)}:</b> {escape(v)}</li>" for k, v in notion_row["fields"].items()
        )
        on_file = f"<h2>On file at the firm (Notion)</h2>\n<ul>{items}</ul>\n<hr>\n"
    return f"""<html><body>
<h1>{escape(name)}</h1>
<p><b>{KIND_LABEL[kind]}{" | Refresher" if notion_row else ""}</b></p>
<p><b>Meeting:</b> {escape(when)} | {escape(ev.get('summary', '(no title)'))}<br>
<b>Email:</b> {escape(email)}<br>
<b>History:</b> {escape(history)}</p>
<hr>
{on_file}{body}
<hr>
<p><i>Auto-generated from public sources and firm records. Double-check anything important before the meeting.</i></p>
</body></html>"""


def save_doc(drive, title, html, folder_id):
    doc = drive.files().create(
        body={
            "name": title,
            "mimeType": "application/vnd.google-apps.document",
            "parents": [folder_id],
        },
        media_body=MediaInMemoryUpload(html.encode("utf-8"), mimetype="text/html"),
        fields="id, webViewLink",
        supportsAllDrives=True,
    ).execute()
    return doc["webViewLink"]


def send_email(gmail, to, subject, html):
    msg = MIMEText(html, "html")
    msg["To"] = ", ".join(to)
    msg["Subject"] = subject
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    gmail.users().messages().send(userId="me", body={"raw": raw}).execute()


def email_html(name, kind, ev, when, link):
    return f"""<p>Prep brief for your upcoming meeting:</p>
<p><b>{escape(name)}</b> ({KIND_LABEL[kind]})<br>
{escape(when)} | {escape(ev.get('summary', '(no title)'))}</p>
<p><a href="{link}">Open the brief</a></p>"""


# ---------------------------------------------------------------------------
# Main loop
# ---------------------------------------------------------------------------

def handle_event(conn, calendar, gmail, drive, ev, key, people, dry_run, notion_rows=None):
    now = datetime.now(timezone.utc)
    notion_rows = notion_rows or []

    # Code decides who's new: skip anyone who's met the firm more than NEW_CONTACT_MAX_MEETINGS times
    counts = {p["email"]: past_meetings(calendar, p["email"], now) for p in people}
    for p in people:
        c = counts[p["email"]]
        if c > NEW_CONTACT_MAX_MEETINGS:
            log.info("Skipping %s: already met the firm %d times in the last %d days", p["email"], c, HISTORY_DAYS)
            if not dry_run:
                mark_processed(conn, key, p["email"], "known")
    people = [p for p in people if counts[p["email"]] <= NEW_CONTACT_MAX_MEETINGS]
    if not people:
        return

    contexts = {p["email"]: email_context(gmail, p["email"]) for p in people}
    kind, names = classify(ev, people, contexts)
    when = meeting_when(ev)

    for p in people:
        email = p["email"]
        name = p.get("displayName") or names.get(email) or email
        try:
            known = notion_match(notion_rows, email, name)
            if known and name == email and known["title"]:
                name = known["title"]  # use the name from Notion when the invite has none
            label = "Meeting Brief (Refresher)" if known else "Meeting Brief"
            title = f"{label} | {name} | {when}"

            history = history_line(counts[email])
            body = research(name, email, kind, ev, contexts[email], history, notion_text(known))
            html = build_doc_html(name, email, kind, ev, when, body, history, known)

            bosses = bosses_for(ev)

            if dry_run:
                log.info("[dry run] would create %r for %s (%s; %s%s)\n%s",
                         title, ", ".join(bosses), KIND_LABEL[kind], history,
                         "; found in Notion" if known else "", body[:1500])
                continue

            links = []
            for boss in bosses:
                folder = BOSS_FOLDERS.get(boss.lower(), BRIEFS_FOLDER_ID)
                link = save_doc(drive, title, html, folder)
                send_email(gmail, [boss], title, email_html(name, kind, ev, when, link))
                links.append(link)
            mark_processed(conn, key, email, kind, links[0])
            log.info("Briefed: %s -> %s", title, ", ".join(bosses))

        except Exception:
            log.exception("Failed on %s, will retry next run", email)


def run(dry_run=False):
    now = datetime.now(timezone.utc)
    calendar, gmail, drive = google_services()
    notion_rows = load_notion_rows()

    with psycopg.connect(os.environ["DATABASE_URL"], autocommit=True) as conn:
        conn.execute(DDL)
        events = upcoming_events(calendar, now)
        log.info("Found %d events in the next %s", len(events), LOOKAHEAD)

        for ev in events:
            key = meeting_key(ev)
            outsiders = outside_attendees(ev)
            if len(outsiders) > MAX_OUTSIDE_ATTENDEES:
                log.info("Skipping %r: %d outside guests looks like a webinar or group invite", ev.get("summary"), len(outsiders))
                continue
            people = [p for p in outsiders if not already_processed(conn, key, p["email"])]

            # Already briefed in an earlier meeting: skip the research, just record it
            repeats = [p for p in people if recently_briefed(conn, p["email"])]
            for p in repeats:
                log.info("Skipping %s: already briefed in the last %d days", p["email"], REBRIEF_DAYS)
                if not dry_run:
                    mark_processed(conn, key, p["email"], "repeat")
            people = [p for p in people if p not in repeats]

            if not people:
                continue
            try:
                handle_event(conn, calendar, gmail, drive, ev, key, people, dry_run, notion_rows)
            except Exception:
                log.exception("Failed on event %r, will retry next run", ev.get("summary"))


# ---------------------------------------------------------------------------
# Preflight check (--check): verifies every connection without writing anything
# ---------------------------------------------------------------------------

def check():
    results = []

    def report(ok, label, detail=""):
        results.append(ok)
        mark = "PASS" if ok is True else ("WARN" if ok is None else "FAIL")
        print(f"[{mark}] {label}" + (f": {detail}" if detail else ""))

    # 1. Env vars
    required = [
        "DATABASE_URL", "OPENROUTER_API_KEY", "GOOGLE_CLIENT_ID", "GOOGLE_CLIENT_SECRET",
        "GOOGLE_REFRESH_TOKEN", "INTERNAL_DOMAINS", "BOSS_EMAILS", "BRIEFS_FOLDER_ID",
    ]
    missing = [v for v in required if not os.environ.get(v)]
    report(not missing, "Env vars", f"missing {', '.join(missing)}" if missing else "all set")
    if missing:
        return False

    for b in BOSS_EMAILS:
        if not is_internal(b):
            report(None, f"Boss email {b}", "not on an INTERNAL_DOMAINS domain, so they'd get researched if invited")

    # 2. Google sign-in + scopes
    try:
        creds = Credentials(
            token=None,
            refresh_token=os.environ["GOOGLE_REFRESH_TOKEN"],
            token_uri="https://oauth2.googleapis.com/token",
            client_id=os.environ["GOOGLE_CLIENT_ID"],
            client_secret=os.environ["GOOGLE_CLIENT_SECRET"],
            scopes=SCOPES,
        )
        creds.refresh(Request())
        report(True, "Google sign-in", "refresh token works")
    except Exception as e:
        report(False, "Google sign-in", f"{e} (rerun --auth; if it says invalid_grant, publish the app first)")
        return False

    granted = set(getattr(creds, "granted_scopes", None) or [])
    if granted:
        for s in SCOPES:
            name = s.rsplit("/", 1)[-1]
            report(s in granted, f"Scope {name}", "granted" if s in granted else "not granted, rerun --auth and approve it")

    calendar = build("calendar", "v3", credentials=creds, cache_discovery=False)
    drive = build("drive", "v3", credentials=creds, cache_discovery=False)

    # 3. Calendars: need reader/writer/owner, free/busy hides attendees
    now = datetime.now(timezone.utc)
    for cal_id in CALENDAR_IDS:
        try:
            resp = calendar.events().list(
                calendarId=cal_id, timeMin=now.isoformat(),
                timeMax=(now + LOOKAHEAD).isoformat(), singleEvents=True, maxResults=50,
            ).execute()
            role = resp.get("accessRole", "unknown")
            n = len(resp.get("items", []))
            if role == "freeBusyReader":
                report(False, f"Calendar {cal_id}", "shared as free/busy only, ask for 'See all event details'")
            else:
                report(True, f"Calendar {cal_id}", f"access: {role}, {n} event(s) in the next {LOOKAHEAD}")
        except Exception as e:
            status = getattr(getattr(e, "resp", None), "status", "")
            hint = " (not shared with this account yet, or a typo)" if str(status) == "404" else ""
            report(False, f"Calendar {cal_id}", f"{status} {e}{hint}")

    # 4. Drive folders: the fallback plus each boss's own folder
    folders = [("Fallback folder", BRIEFS_FOLDER_ID)] + [(f"Folder for {b}", f) for b, f in BOSS_FOLDERS.items()]
    for label, folder_id in folders:
        try:
            f = drive.files().get(
                fileId=folder_id, supportsAllDrives=True,
                fields="id, name, capabilities(canAddChildren)",
            ).execute()
            can_add = f.get("capabilities", {}).get("canAddChildren", False)
            report(can_add, f"{label} '{f.get('name')}'",
                   "can create docs" if can_add else "can see it but can't add files, need Contributor or higher")
        except Exception as e:
            report(False, label, f"{e} (check the folder ID and that your account has access)")

    # 5. Postgres
    try:
        with psycopg.connect(os.environ["DATABASE_URL"], connect_timeout=10) as conn:
            conn.execute("SELECT 1")
        report(True, "Postgres", "connected")
    except Exception as e:
        hint = " (on your laptop, use the PUBLIC url)" if "railway.internal" in os.environ["DATABASE_URL"] else ""
        report(False, "Postgres", f"{e}{hint}")

    # 6. Notion (optional)
    if NOTION_API_KEY and NOTION_DATABASE_ID:
        rows = load_notion_rows()
        report(bool(rows), "Notion", f"{len(rows)} rows loaded" if rows
               else "no rows, check the key, database ID, and that the integration is connected to the database")
    else:
        report(None, "Notion", "not set up, briefs won't include firm records")

    # 7. OpenRouter + model ID
    try:
        out = ask(LLM_MODEL, "Reply with just the word OK.")
        report(True, f"OpenRouter model {LLM_MODEL}", f"replied: {out.strip()[:20]}")
    except Exception as e:
        report(False, f"OpenRouter model {LLM_MODEL}", f"{e} (check the API key, credits, and exact model ID)")

    failed = results.count(False)
    print("\nAll good, ready to run." if not failed else f"\n{failed} check(s) failed.")
    return not failed


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--auth", action="store_true", help="one-time Google sign-in, prints refresh token")
    parser.add_argument("--dry-run", action="store_true", help="log what would happen without writing anything")
    parser.add_argument("--check", action="store_true", help="test every connection and permission, writes nothing")
    args = parser.parse_args()

    if args.auth:
        print_refresh_token()
    elif args.check:
        raise SystemExit(0 if check() else 1)
    else:
        run(dry_run=args.dry_run)