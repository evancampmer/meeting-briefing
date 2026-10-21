"""
Meeting Briefs

Hourly job that:
  1. Looks at the next 26 hours of calendar events
  2. Finds outside attendees (anyone not on an internal domain) who haven't been briefed yet
  3. Classifies the meeting: potential client, potential contractor, or other (skipped)
  4. Researches each new outside person on the web (LLM + web search)
  5. Saves a Google Doc brief titled "Meeting Brief | <name or email> | <date, time>"
  6. Emails the doc link to the bosses

Because it runs every hour over a 26 hour window, normal meetings get briefed about
a day ahead and last-minute bookings get picked up within the hour. Postgres tracks
who's been briefed so nobody gets researched twice.

Setup:
  pip install -r requirements.txt
  python meeting_briefs.py --auth        # one time, prints GOOGLE_REFRESH_TOKEN
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
  BOSS_EMAILS            comma list of who gets the briefs
  BRIEFS_FOLDER_ID       Drive folder shared with the bosses
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
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaInMemoryUpload
from openai import OpenAI

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
BOSS_EMAILS = env_list("BOSS_EMAILS")
BRIEFS_FOLDER_ID = os.environ.get("BRIEFS_FOLDER_ID", "")
BOSS_TZ = ZoneInfo(os.environ.get("BOSS_TZ", "America/Los_Angeles"))
LOOKAHEAD = timedelta(hours=int(os.environ.get("LOOKAHEAD_HOURS", "26")))
LLM_MODEL = os.environ.get("LLM_MODEL", "openai/gpt-4o-mini")

PERSONAL_DOMAINS = {
    "gmail.com", "googlemail.com", "yahoo.com", "outlook.com", "hotmail.com",
    "live.com", "icloud.com", "me.com", "aol.com", "proton.me", "protonmail.com",
}

KIND_LABEL = {"client": "Potential Client", "contractor": "Potential Contractor"}


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


def mark_processed(conn, key, email, kind, doc_url=None):
    conn.execute(
        """INSERT INTO briefed_attendees (meeting_key, email, kind, doc_url)
           VALUES (%s, %s, %s, %s) ON CONFLICT DO NOTHING""",
        (key, email, kind, doc_url),
    )


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
    seen, events = set(), []
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
                key = meeting_key(ev)
                if key in seen:  # same meeting on two bosses' calendars
                    continue
                seen.add(key)
                events.append(ev)
            page_token = resp.get("nextPageToken")
            if not page_token:
                break
    return events


def is_internal(email):
    domain = email.split("@")[-1].lower()
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


CLASSIFY_PROMPT = """You sort meetings for a small consulting firm. The firm takes on clients and hands the work to contractors, and it also interviews contractors who want to join as consultants.

Meeting title: {title}
Meeting description: {description}
Outside attendees: {emails}

Recent emails with these people:
{context}

Reply with JSON only, in this shape:
{{"kind": "client" | "contractor" | "other", "names": {{"<email>": "<full name or null>"}}}}

- "client": a potential client the firm may take on
- "contractor": someone interviewing to join the firm as a consultant/contractor
- "other": anything else (vendors, personal, internal, not an interview)

Only give a name if it actually appears in the title, description, or emails (for example an email signature or a From header). Never guess a name."""


def classify(ev, people, contexts):
    prompt = CLASSIFY_PROMPT.format(
        title=ev.get("summary", "(no title)"),
        description=(ev.get("description") or "(none)")[:3000],
        emails=", ".join(p["email"] for p in people),
        context="\n\n".join(f"[{e}]\n{c or '(no emails found)'}" for e, c in contexts.items()),
    )
    data = json.loads(strip_fences(ask(LLM_MODEL, prompt, json_mode=True)))  # raises -> retried next run
    kind = data.get("kind") if data.get("kind") in KIND_LABEL else "other"
    names = {k.lower(): v for k, v in (data.get("names") or {}).items() if v}
    return kind, names


RESEARCH_PROMPT = """Research this person ahead of a meeting with a small consulting firm.

Name: {name}
Email: {email}
Company website (from their email domain): {site}
Meeting type: {kind_label}
Meeting title: {title}

Context from recent emails:
{context}

Search the web for their professional background. Try "{name} {company_hint} LinkedIn", their company website, articles, talks, GitHub, and similar public sources.

Rules:
- Only include facts you found in a source, with the source URL linked right next to each fact.
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


def research(name, email, kind, ev, context):
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
        third_section=(
            "Company (what their company does, size, industry, likely needs)"
            if kind == "client"
            else "Skills and fit (what they could bring to the firm as a contractor)"
        ),
    )
    return strip_fences(ask(f"{LLM_MODEL}:online", prompt))  # :online = OpenRouter web search


# ---------------------------------------------------------------------------
# Output: Google Doc + email
# ---------------------------------------------------------------------------

def build_doc_html(name, email, kind, ev, when, body):
    return f"""<html><body>
<h1>{escape(name)}</h1>
<p><b>{KIND_LABEL[kind]}</b></p>
<p><b>Meeting:</b> {escape(when)} | {escape(ev.get('summary', '(no title)'))}<br>
<b>Email:</b> {escape(email)}</p>
<hr>
{body}
<hr>
<p><i>Auto-generated from public sources. Double-check anything important before the meeting.</i></p>
</body></html>"""


def save_doc(drive, title, html):
    """Uploads HTML and lets Drive convert it into a real Google Doc."""
    doc = drive.files().create(
        body={
            "name": title,
            "mimeType": "application/vnd.google-apps.document",
            "parents": [BRIEFS_FOLDER_ID],
        },
        media_body=MediaInMemoryUpload(html.encode("utf-8"), mimetype="text/html"),
        fields="id, webViewLink",
        supportsAllDrives=True,
    ).execute()
    return doc["webViewLink"]


def send_email(gmail, subject, html):
    msg = MIMEText(html, "html")
    msg["To"] = ", ".join(BOSS_EMAILS)
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

def handle_event(conn, gmail, drive, ev, key, people, dry_run):
    contexts = {p["email"]: email_context(gmail, p["email"]) for p in people}
    kind, names = classify(ev, people, contexts)
    when = meeting_when(ev)

    if kind == "other":
        log.info("Skipping %r: not a client/contractor interview", ev.get("summary"))
        if not dry_run:
            for p in people:
                mark_processed(conn, key, p["email"], kind)
        return

    for p in people:
        email = p["email"]
        name = p.get("displayName") or names.get(email) or email
        title = f"Meeting Brief | {name} | {when}"
        try:
            body = research(name, email, kind, ev, contexts[email])
            html = build_doc_html(name, email, kind, ev, when, body)

            if dry_run:
                log.info("[dry run] would create %r\n%s", title, body[:1500])
                continue

            link = save_doc(drive, title, html)
            send_email(gmail, title, email_html(name, kind, ev, when, link))
            mark_processed(conn, key, email, kind, link)
            log.info("Briefed: %s", title)
        except Exception:
            log.exception("Failed on %s, will retry next run", email)


def run(dry_run=False):
    now = datetime.now(timezone.utc)
    calendar, gmail, drive = google_services()

    with psycopg.connect(os.environ["DATABASE_URL"], autocommit=True) as conn:
        conn.execute(DDL)
        events = upcoming_events(calendar, now)
        log.info("Found %d events in the next %s", len(events), LOOKAHEAD)

        for ev in events:
            key = meeting_key(ev)
            people = [p for p in outside_attendees(ev) if not already_processed(conn, key, p["email"])]
            if not people:
                continue
            try:
                handle_event(conn, gmail, drive, ev, key, people, dry_run)
            except Exception:
                log.exception("Failed on event %r, will retry next run", ev.get("summary"))


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--auth", action="store_true", help="one-time Google sign-in, prints refresh token")
    parser.add_argument("--dry-run", action="store_true", help="log what would happen without writing anything")
    args = parser.parse_args()

    if args.auth:
        print_refresh_token()
    else:
        run(dry_run=args.dry_run)