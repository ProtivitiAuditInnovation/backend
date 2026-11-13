#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
User Access Review (UAR) testing engine - used by the web backend and as a command-line tool
============================================================================================
Originally built for TRAX (LA-CP4); application name, control owner, labels and codes are options.
Replicates the IA testing walked through on the 26-Oct-2026 call:

  1. IPE            User Report xlsx vs PDF - row count and "Created on" timestamp
  2. Completeness   every row of the User Report is in the review file (and vice versa)
  3. Responses      manager approval e-mails (.msg / .eml) -> response date (MM-DD-YYYY),
                    timeliness against the due date
  4. Status         Approved / Rejected per access row (review file), cross-checked
                    against what the manager actually wrote in the e-mail
  5. Self-approval  user name = manager name
  6. Post-review    rejected access must be gone from the After Listing report;
                    reconciled to the client's "Analysis of Removed Accounts"
  7. Movement       rows removed / added between the start and after listings
  8. Summary        formula-driven, recalculates when the workbook is opened

Usage
-----
  python trax_uar_automation.py --input "C:\\...\\LA-CP4"
  python trax_uar_automation.py --input "C:\\...\\LA-CP4" --due 2026-04-16 --output "UAR result.xlsx"
  python trax_uar_automation.py --input "C:\\...\\LA-CP4" --no-ai        (keyword rules only)

Azure OpenAI (optional) - reads the manager e-mail replies
---------------------------------------------------------
Put these in a file called .env next to this script (never in the script itself):
  AZURE_OPENAI_ENDPOINT=https://<resource>.openai.azure.com/
  AZURE_OPENAI_API_KEY=<key>
  AZURE_OPENAI_DEPLOYMENT=<deployment name>
  AZURE_OPENAI_API_VERSION=2024-12-01-preview
Only e-mail reply text and the manager's user list are sent. Every AI answer must quote the e-mail
word for word; quotes are verified in Python and unverifiable removal claims are not counted.
Answers are cached in <input>\\.uar_ai_cache.json so re-runs give identical results.

Requirements:  pip install pandas openpyxl pdfplumber extract-msg
"""
import argparse
import glob
import hashlib
import html
import io
import json
import os
import re
import sys
import time
import unicodedata
from html.parser import HTMLParser
import urllib.error
import urllib.request
from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties

# =============================================================================
# CONFIG - edit here if file names / labels change between review periods
# =============================================================================
CONFIG = {
    # glob patterns, searched in the input folder ([0-9] keeps "After Listing" out)
    "patterns": {
        "user_report_xlsx": ["*User Report [0-9]*.xlsx"],
        "user_report_pdf":  ["*User Report [0-9]*.pdf"],
        "after_xlsx":       ["*After Listing*.xlsx"],
        "after_pdf":        ["*After Listing*.pdf"],
        "review_xlsx":      ["*Periodic Access Review*.xlsx"],
        "removed_xlsx":     ["*Removed Accounts*.xlsx"],
        "steps_doc":        ["*Steps to Create User Report*.docx"],
    },
    "app_name": "TRAX",                # used in the AI prompt and report titles
    "emails_dir": "emails",            # searched recursively for .msg / .eml
    "removed_sheet": "English",        # sheet of the removed-accounts analysis
    "due_date": None,                  # informational: read "Due Thursday, April 16" from e-mail subjects
    "response_days": 30,               # a reply within this many days of the request e-mail is on time
    "dormant_days": 90,
    "elevated_codes": ["AC", "SA", "UA", "UAO"],
    "status_approve": "Approved",
    "status_remove": "Rejected",       # IA analysis file used "Remove" - change here if preferred
    "control_owner": "Crystal Hood",   # sends the requests/reminders - never counted as a response
    "request_markers": [               # text that identifies the original request / reminders
        "as part of the trax it general controls periodic access review",
        "i need you to review the list of your direct reports",
    ],
    "output": "LA-CP4 TRAX UAR Testing Result.xlsx",
}

# keyword fallback (used when Azure OpenAI is off or a call fails); matched at word starts, accents ignored
REMOVE_KW = ["disable", "remove", "revoke", "terminate", "delete", "deactivate", "offboard",
             "no longer", "not required", "not needed", "left the company", "has left"]
APPROVE_KW = ["approve", "approval", "retain", "appropriate", "confirm", "comfirm", "no change", "looks good",
              "all good", "are correct", "is correct", "keep", "same access", "fine", "ok to"]


def _kw_re(words):
    return re.compile(r"\b(?:" + "|".join(re.escape(w) for w in words) + r")", re.I)


REMOVE_RE, APPROVE_RE = _kw_re(REMOVE_KW), _kw_re(APPROVE_KW)
COMMON_TOKENS = {"the", "and", "admin", "support", "user", "test", "contractor", "team", "service"}

CANON = ["User code", "User name", "User email", "Active user", "Manager", "Access Confirmed",
         "Access Removed", "Last login", "Type of membership", "Name", "Code"]


# =============================================================================
# Generic helpers
# =============================================================================
_LOG = [print]          # the web backend swaps this for a per-run logger


def log(msg):
    _LOG[0](msg)


def warn(msg):
    log(f"  [WARN] {msg}")


def norm_col(c):
    return re.sub(r"\s+", " ", str(c if c is not None else "")).strip()


def sval(v):
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return ""
    if isinstance(v, (pd.Timestamp, datetime)):
        return v.strftime("%Y-%m-%d %H:%M")
    return str(v).strip()


def to_naive(ts):
    if ts is None:
        return None
    try:
        ts = pd.Timestamp(ts)
    except Exception:
        return None
    if pd.isna(ts):
        return None
    if ts.tzinfo is not None:
        ts = ts.tz_localize(None)          # keep the sender's local wall-clock time
    return ts


def parse_dt(s):
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return None
    if isinstance(s, (pd.Timestamp, datetime)):
        return to_naive(s)
    if isinstance(s, (int, float)):                       # Excel serial
        return to_naive(pd.Timestamp("1899-12-30") + pd.to_timedelta(float(s), unit="D"))
    s = str(s).strip()
    if not s:
        return None
    for fmt in ("%b %d, %Y %I:%M %p", "%b %d, %Y %I:%M:%S %p", "%m-%d-%Y", "%m/%d/%Y"):
        try:
            return pd.Timestamp(datetime.strptime(s, fmt))
        except ValueError:
            pass
    try:
        from dateutil import parser as dp
        return to_naive(dp.parse(s, fuzzy=True))
    except Exception:
        return None


def date_from_name(path):
    """'... 260401.xlsx' -> 2026-04-01 (yymmdd suffix used in the evidence file names)."""
    if not path:
        return None
    m = re.findall(r"(\d{6})", os.path.basename(path))
    if not m:
        return None
    try:
        return pd.Timestamp(datetime.strptime(m[-1], "%y%m%d"))
    except ValueError:
        return None


def find_file(folder, key, required=True):
    for pat in CONFIG["patterns"][key]:
        hits = sorted(h for h in glob.glob(os.path.join(folder, pat))
                      if not os.path.basename(h).startswith("~$"))
        if hits:
            if len(hits) > 1:
                warn(f"{key}: {len(hits)} files match, using {os.path.basename(hits[0])}")
            return hits[0]
    if required:
        sys.exit(f"ERROR: no file for '{key}' (patterns {CONFIG['patterns'][key]}) in {folder}")
    warn(f"{key}: not found (optional)")
    return None


def fold(s):
    """'Séverine' -> 'Severine' (accents ignored when matching names)."""
    return unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode("ascii")


def tokens(s, min_len=2):
    return [t for t in re.split(r"[^a-z]+", fold(s).lower()) if len(t) >= min_len]


def person_tokens(s):
    """'Young, Carrie <Carrie.Young@adm.com>' -> {'carrie', 'young'}"""
    s = str(s or "")
    emails = re.findall(r"[\w.\-']+@[\w\-.]+", s)
    disp = re.sub(r"\S+@\S+|mailto:|[<>\[\]\"']", " ", s).strip()
    if "," in disp:
        last, first = disp.split(",", 1)
        disp = f"{first} {last}"
    out = set(tokens(disp))
    for e in emails:
        out |= set(tokens(e.split("@")[0]))
    return out


def _edit1(a, b):
    """True when a and b differ by at most one inserted, deleted or changed letter."""
    la, lb = len(a), len(b)
    if abs(la - lb) > 1:
        return False
    i = j = d = 0
    while i < la and j < lb:
        if a[i] == b[j]:
            i += 1
            j += 1
            continue
        d += 1
        if d > 1:
            return False
        if la > lb:
            i += 1
        elif lb > la:
            j += 1
        else:
            i += 1
            j += 1
    return d + (la - i) + (lb - j) <= 1


def _close(a, b):
    """Same name part: equal, a short form of 3-4 letters (Dan/Daniel, Tom/Tomasz), or a one-letter typo in a
    name of 6+ letters (Johnathan/Jonathan). Gabriel/Gabriela and Marie/Mario stay different people."""
    if a == b:
        return True
    if a.startswith(b) or b.startswith(a):
        return 3 <= min(len(a), len(b)) <= 4
    return min(len(a), len(b)) >= 6 and _edit1(a, b)


def same_person(name, toks):
    """Does a person name match a set of name tokens (from a display name / e-mail address)?"""
    t = tokens(name)
    if not t or not toks:
        return False
    first, last = t[0], t[-1]
    if {first, last} <= toks:
        return True
    last_hits = {x for x in toks if _close(last, x)}
    return bool(last_hits) and any(_close(first, y) for y in toks if y not in last_hits or first == last)


def same_surname(name, toks):
    """Lenient check used only to decide whether a reply came from the manager themselves (so typos in the
    review file, e.g. Smyth/Smith, don't raise false 'sent by someone else' flags): surname within one
    letter + same first initial."""
    t = tokens(name)
    if not t or not toks:
        return False
    if same_person(name, toks):
        return True
    return any(x == t[-1] or _edit1(t[-1], x) for x in toks) and any(y[:1] == t[0][:1] for y in toks)


# =============================================================================
# Readers
# =============================================================================
def read_table(path, must_have, sheet=None, scan_rows=15):
    """Find the header row (row containing `must_have`) and return the table as text."""
    xl = pd.ExcelFile(path)
    sheets = [sheet] if sheet and sheet in xl.sheet_names else xl.sheet_names
    for sh in sheets:
        raw = pd.read_excel(xl, sheet_name=sh, header=None, nrows=scan_rows,
                            dtype=object, keep_default_na=False)
        for i, row in raw.iterrows():
            if any(norm_col(v).lower() == must_have.lower() for v in row.tolist()):
                # keep_default_na=False -> the region code "NA" stays "NA"
                df = pd.read_excel(xl, sheet_name=sh, header=i, dtype=object,
                                   keep_default_na=False, na_values=[""])
                df.columns = [norm_col(c) for c in df.columns]
                return df.dropna(how="all"), sh
    raise ValueError(f"No header row with '{must_have}' found in {os.path.basename(path)} (first {scan_rows} rows checked)")


FOOTER_RE = re.compile(r"^\s*(created\s+on\b|©|\(c\)|copyright\b)|\bsungard\b|\btrademark", re.I)


def standardise_report(df, label, info=None):
    """info (optional dict) receives the report's own 'Created on' timestamp from its footer lines."""
    rename = {}
    for canon in CANON:
        for c in df.columns:
            lc = c.lower()
            if lc == canon.lower() or (len(canon) > 8 and lc.startswith(canon.lower()[:12])):
                rename[c] = canon
                break
    df = df.rename(columns=rename)
    missing = [c for c in ["User code", "User name", "Type of membership", "Code"] if c not in df.columns]
    if missing:
        raise ValueError(f"{label} is missing columns {missing}. Found: {list(df.columns)}")
    for c in df.columns:
        if c != "Last login":
            df[c] = df[c].map(sval)
    df = df[df["User code"] != ""].copy()
    # report exports end with footer lines ("Created on: ...", "(c)2014 SunGard ...") in the User code column
    foot = df["User code"].map(lambda v: bool(FOOTER_RE.search(v)))
    for v in df.loc[foot, "User code"]:
        m = re.search(r"created\s+on:?\s*(.+)$", v, re.I)
        if m and info is not None and "created" not in info:
            info["created_text"] = m.group(1).strip()
            info["created"] = parse_dt(m.group(1).strip())
    blank = ~foot & (df["Type of membership"] == "") & (df["Code"] == "")
    if info is not None:
        info["footer_rows"], info["blank_rows"] = int(foot.sum()), int(blank.sum())
    if foot.any() or blank.any():
        log(f"  {label}: ignored {int(foot.sum())} footer line(s)"
            + (f" and {int(blank.sum())} row(s) with no access type/code" if blank.any() else ""))
    df = df[~foot & ~blank].copy()
    if "Last login" in df.columns:
        df["Last login"] = df["Last login"].map(parse_dt)
    df.insert(0, "Key", (df["User code"].str.upper() + "|" + df["Type of membership"].str.upper()
                         + "|" + df["Code"].str.upper()))
    dups = df["Key"].duplicated().sum()
    if dups:
        warn(f"{label}: {dups} duplicate user/type/code rows")
    return df.reset_index(drop=True)


def html_to_text(h):
    h = re.sub(r"(?is)<(script|style).*?</\1>", " ", h)
    h = re.sub(r"(?i)<br\s*/?>|</p>|</div>|</tr>", "\n", h)
    return html.unescape(re.sub(r"<[^>]+>", " ", h))


def read_one_email(path):
    """-> dict(sender, to, sent, subject, body (plain text), html)."""
    if path.lower().endswith(".msg"):
        import extract_msg
        m = extract_msg.Message(path)
        try:
            sender, to, sent, subject = m.sender or "", m.to or "", m.date, m.subject or ""
            body = m.body or ""
            hb = getattr(m, "htmlBody", None)
            html_body = hb.decode("utf-8", "ignore") if isinstance(hb, bytes) else (hb or "")
        finally:
            m.close()
    else:
        from email import policy
        from email.parser import BytesParser
        with open(path, "rb") as fh:
            m = BytesParser(policy=policy.default).parse(fh)
        sender, to, sent, subject = str(m["from"] or ""), str(m["to"] or ""), m["date"], str(m["subject"] or "")
        p_plain, p_html = m.get_body(preferencelist=("plain",)), m.get_body(preferencelist=("html",))
        body = p_plain.get_content() if p_plain is not None else ""
        html_body = p_html.get_content() if p_html is not None else ""
    if isinstance(body, bytes):
        body = body.decode("utf-8", "ignore")
    if not str(body).strip() and html_body:
        body = html_to_text(html_body)
    sent = to_naive(sent) if isinstance(sent, datetime) else to_naive(parse_dt(sent))
    return {"sender": sender, "to": to, "sent": sent, "subject": subject, "body": body, "html": html_body}


HDR = re.compile(r"^[ \t>*]*From:[ \t]*(?P<frm>.+?)[ \t]*\n(?:.*\n){0,3}?"
                 r"^[ \t>*]*(?:Sent|Date):[ \t]*(?P<sent>.+?)[ \t]*$", re.M)
TO_LINE = re.compile(r"^[ \t>*]*To:[ \t]*(?P<to>.+?)[ \t]*$", re.M)


def display_name(s):
    """'Doe, Jane <Jane.Doe@corp.com>' or '"Doe; Jane" <...>' -> 'Jane Doe'."""
    s = str(s or "").strip()
    if not s:
        return ""
    q = re.match(r'\s*"([^"]+)"', s)
    if q:
        disp = q.group(1)
    else:
        first = re.split(r";\s*(?=[^;]*[@<])", s)[0] if s.count("@") > 1 else s
        disp = re.sub(r"<[^>]*>|mailto:\S+", " ", first).strip()
    if not disp or "@" in disp:
        m = re.search(r"([\w.\-']+)@", s)
        return " ".join(p.capitalize() for p in re.split(r"[._]", m.group(1)) if p) if m else disp
    for sep in (",", ";"):
        if sep in disp:
            last, first_ = disp.split(sep, 1)
            disp = f"{first_.strip()} {last.strip()}"
            break
    return re.sub(r"\s+", " ", disp).strip()


def clean_block(t):
    t = re.sub(r"^[ \t>*]*(To|Cc|Subject|Importance):.*$", "", t, flags=re.M)
    t = re.sub(r"(?i)classification:\s*\w+", "", t)
    t = re.sub(r"[ \t]*\n\s*\n+", "\n", t)
    return t.strip()


def classify(text):
    t = fold(text).lower()
    if any(m in t for m in CONFIG["request_markers"]):
        return "Request / reminder"
    rem = bool(REMOVE_RE.search(t))
    app = bool(APPROVE_RE.search(t))
    if rem and app:
        return "Approve + removal request"
    return "Removal request" if rem else "Approve" if app else "Unclear - read e-mail"


class _TableParser(HTMLParser):
    """Collects every <table> in an HTML document as a list of rows (lists of cell text), using only the
    standard library. Text inside a nested table is kept with that nested table, not the outer cell."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.tables, self.stack = [], []

    def handle_starttag(self, tag, attrs):
        t = tag.lower()
        if t == "table":
            tbl = {"rows": [], "row": None, "cell": None}
            self.stack.append(tbl)
            self.tables.append(tbl["rows"])
        elif not self.stack:
            return
        elif t == "tr":
            cur = self.stack[-1]
            self._close_row(cur)
            cur["row"] = []
        elif t in ("td", "th"):
            cur = self.stack[-1]
            self._close_cell(cur)
            if cur["row"] is None:
                cur["row"] = []
            cur["cell"] = []
        elif t in ("br", "p", "div") and self.stack[-1]["cell"] is not None:
            self.stack[-1]["cell"].append(" ")

    def handle_endtag(self, tag):
        t = tag.lower()
        if not self.stack:
            return
        cur = self.stack[-1]
        if t in ("td", "th"):
            self._close_cell(cur)
        elif t == "tr":
            self._close_row(cur)
        elif t == "table":
            self._close_row(cur)
            self.stack.pop()

    def handle_data(self, data):
        if self.stack and self.stack[-1]["cell"] is not None:
            self.stack[-1]["cell"].append(data)

    @staticmethod
    def _close_cell(cur):
        if cur["cell"] is not None:
            if cur["row"] is None:
                cur["row"] = []
            cur["row"].append(re.sub(r"\s+", " ", "".join(cur["cell"])).strip())
            cur["cell"] = None

    def _close_row(self, cur):
        self._close_cell(cur)
        if cur["row"] is not None:
            cur["rows"].append(cur["row"])
            cur["row"] = None


def html_tables(html_body):
    p = _TableParser()
    try:
        p.feed(html_body)
        p.close()
    except Exception as e:                      # malformed HTML: keep what was parsed
        warn(f"HTML table parsing stopped early: {e}")
    return p.tables


def request_tables(html_body):
    """Tables in the e-mail HTML whose header row has 'User code'. Outlook keeps the newest reply at the top,
    so the last such table in the document belongs to the oldest message: the original request."""
    if not html_body or "<table" not in html_body.lower():
        return []
    out = []
    for rows in html_tables(html_body):
        hdr_i = next((i for i, r in enumerate(rows[:6]) if any(norm_col(c).lower() == "user code" for c in r)), None)
        if hdr_i is None:
            continue
        header = [norm_col(c) for c in rows[hdr_i]]
        data = [r + [""] * (len(header) - len(r)) for r in rows[hdr_i + 1:] if any(str(c).strip() for c in r)]
        df = pd.DataFrame([r[:len(header)] for r in data], columns=header) if data else pd.DataFrame(columns=header)
        df = df.loc[:, [c for c in df.columns if c]]
        df = df.loc[:, ~pd.Index(df.columns).duplicated()]
        out.append(df)
    return out


def read_email_blocks(folder):
    """Every e-mail file is split into its own message + each quoted/forwarded message (newest first).
    Returns (blocks, subjects, n_files, request_rows). request_rows = the access rows pasted in the
    oldest message of each trail (the original review request), with its recipient and sent date."""
    files = sorted(f for f in glob.glob(os.path.join(folder, "**", "*"), recursive=True)
                   if f.lower().endswith((".msg", ".eml")))
    blocks, subjects, req_frames = [], [], []
    for f in files:
        try:
            em = read_one_email(f)
        except Exception as e:
            warn(f"could not read {os.path.basename(f)}: {e}")
            continue
        subjects.append(em["subject"])
        body = em["body"].replace("\r\n", "\n").replace("\r", "\n")
        hits = list(HDR.finditer(body))
        parts = [(em["sender"], em["to"], em["sent"], body[: hits[0].start()] if hits else body)]
        for i, h in enumerate(hits):
            end = hits[i + 1].start() if i + 1 < len(hits) else len(body)
            seg = body[h.end(): end]
            tm = TO_LINE.search(seg[:600])
            to = tm["to"] if (tm and seg[:tm.start()].count("\n") <= 3) else ""
            parts.append((h["frm"], to, to_naive(parse_dt(h["sent"])), seg))
        fblocks = []
        for lvl, (author, to, dt, text) in enumerate(parts):
            text = clean_block(text)
            fblocks.append({"File": os.path.basename(f), "Block": lvl, "Author": str(author).strip(),
                            "To": str(to).strip(), "Sent": dt, "Subject": em["subject"], "Text": text,
                            "Class": classify(text), "_toks": person_tokens(author),
                            "Oldest": lvl == len(parts) - 1, "File Manager": ""})
        blocks += fblocks

        tabs = request_tables(em["html"])
        if not tabs and "user code" in fold(em["body"] + em["html"]).lower():
            warn(f"{os.path.basename(f)}: mentions 'User code' but no pasted list could be read")
        if tabs:
            oldest = fblocks[-1]
            try:
                t = standardise_report(tabs[-1], f"request table in {os.path.basename(f)}")
            except ValueError as e:
                warn(str(e))
                continue
            t["Request File"] = os.path.basename(f)
            t["Request To"] = display_name(oldest["To"]) if oldest["To"] else ""
            t["Request From"] = display_name(oldest["Author"])
            t["Request Sent"] = oldest["Sent"]
            t["Table Manager"] = t["Manager"] if "Manager" in t.columns else ""
            mgr_vals = [m for m in t["Table Manager"].tolist() if str(m).strip()]
            file_mgr = max(set(mgr_vals), key=mgr_vals.count) if mgr_vals else display_name(oldest["To"])
            for b in fblocks:
                b["File Manager"] = file_mgr
            req_frames.append(t)

    # the same reply is quoted in many threads -> de-duplicate
    seen, out = set(), []
    for b in blocks:
        k = (tuple(sorted(b["_toks"])), b["Sent"].floor("min") if b["Sent"] is not None else None,
             b["Text"][:80], b["To"])
        if k not in seen:
            seen.add(k)
            out.append(b)
    req = pd.concat(req_frames, ignore_index=True) if req_frames else pd.DataFrame()
    return out, subjects, len(files), req


def due_from_subjects(subjects, year):
    for s in subjects:
        m = re.search(r"Due\s+(?:[A-Za-z]+day[,_]?\s+)?([A-Z][a-z]+)\s+(\d{1,2})", s or "")
        if m:
            try:
                return pd.Timestamp(datetime.strptime(f"{m[1]} {m[2]} {year}", "%B %d %Y"))
            except ValueError:
                continue
    return None


def pdf_check(path):
    res = {"pages": None, "created": "", "rows": None, "users": None, "note": ""}
    if not path:
        res["note"] = "PDF not found"
        return res
    try:
        import pdfplumber
    except ImportError:
        res["note"] = "pip install pdfplumber to read PDF"
        return res
    try:
        code_re = re.compile(r"\b[A-Za-z]\d{6}\b")
        text = []
        with pdfplumber.open(path) as pdf:
            res["pages"] = len(pdf.pages)
            for p in pdf.pages:
                best = p.extract_text() or ""
                # the TRAX report table is printed sideways: read rotated characters in each direction
                if p.chars and sum(not ch.get("upright", True) for ch in p.chars) > len(p.chars) / 2:
                    for ld, cd in (("ltr", "btt"), ("rtl", "btt"), ("ltr", "ttb"), ("rtl", "ttb")):
                        try:
                            t = p.extract_text(line_dir_rotated=ld, char_dir_rotated=cd) or ""
                        except TypeError:          # older pdfplumber without rotation options
                            break
                        if len(code_re.findall(t)) > len(code_re.findall(best)):
                            best = t
                text.append(best)
        text = "\n".join(text)
        m = re.search(r"Created\s*on:?\s*([A-Z][a-z]{2,8}\s+\d{1,2},\s*\d{4}\s+\d{1,2}:\d{2}\s*[AP]M)", text)
        res["created"] = re.sub(r"\s+", " ", m[1]) if m else "not found"
        codes = re.findall(r"\b[A-Za-z]\d{6}\b", text)
        if len(text.strip()) < 20 * max(res["pages"] or 1, 1):
            res["rows"] = res["users"] = None
            res["note"] = ("PDF has no text layer (scanned or printed as an image): row count and timestamp "
                           "must be checked manually or with OCR")
        elif not codes:
            res["rows"] = res["users"] = None
            res["note"] = "no user codes found in the PDF text - verify manually"
        else:
            res["rows"], res["users"] = len(codes), len({c.upper() for c in codes})
    except Exception as e:
        res["note"] = f"error reading PDF: {e}"
    return res


# =============================================================================
# Optional Azure OpenAI - reads manager e-mail replies (free text)
# =============================================================================
def load_env(paths):
    """Minimal .env reader: KEY=VALUE lines; existing environment variables win."""
    for p in paths:
        if p and os.path.isfile(p):
            with open(p, encoding="utf-8-sig") as fh:
                for line in fh:
                    line = line.strip()
                    if line and not line.startswith("#") and "=" in line:
                        k, v = line.split("=", 1)
                        os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


class AzureAI:
    def __init__(self, cache_path):
        self.endpoint = os.environ.get("AZURE_OPENAI_ENDPOINT", "").strip().rstrip("/")
        self.key = os.environ.get("AZURE_OPENAI_API_KEY", "").strip()
        self.deployment = os.environ.get("AZURE_OPENAI_DEPLOYMENT", "").strip()
        self.version = os.environ.get("AZURE_OPENAI_API_VERSION", "2024-12-01-preview").strip()
        self.enabled = bool(self.endpoint and self.key and self.deployment)
        self.cache_path = cache_path
        self.cache = {}
        if cache_path and os.path.isfile(cache_path):
            try:
                with open(cache_path, encoding="utf-8") as fh:
                    self.cache = json.load(fh)
            except Exception:
                warn("AI cache unreadable - starting a new one")
        self.json_mode = True
        self.calls = self.hits = self.errors = 0

    def _save(self):
        if self.cache_path:
            with open(self.cache_path, "w", encoding="utf-8") as fh:
                json.dump(self.cache, fh, indent=1, ensure_ascii=False)

    @staticmethod
    def _parse(txt):
        txt = re.sub(r"^```(?:json)?|```$", "", (txt or "").strip(), flags=re.M).strip()
        try:
            return json.loads(txt)
        except ValueError:
            m = re.search(r"\{.*\}", txt, re.S)
            return json.loads(m.group(0)) if m else None

    def chat(self, system, user, use_cache=True, messages=None):
        """messages (optional) = full message list for multi-turn chat; otherwise system + user."""
        msgs = messages or [{"role": "system", "content": system}, {"role": "user", "content": user}]
        h = hashlib.sha256((self.deployment + json.dumps(msgs, ensure_ascii=False)).encode("utf-8")).hexdigest()
        if use_cache and h in self.cache:
            self.hits += 1
            return self.cache[h]
        url = (f"{self.endpoint}/openai/deployments/{self.deployment}/chat/completions"
               f"?api-version={self.version}")
        body = {"messages": msgs}
        for attempt in range(5):
            if self.json_mode:
                body["response_format"] = {"type": "json_object"}
            else:
                body.pop("response_format", None)
            req = urllib.request.Request(url, data=json.dumps(body).encode("utf-8"), method="POST",
                                         headers={"Content-Type": "application/json", "api-key": self.key})
            try:
                with urllib.request.urlopen(req, timeout=180) as r:
                    data = json.loads(r.read().decode("utf-8"))
                content = data["choices"][0]["message"].get("content")
                out = self._parse(content)
                if not isinstance(out, dict):
                    raise ValueError("model did not return a JSON object")
                self.calls += 1
                if use_cache:
                    self.cache[h] = out
                    self._save()
                return out
            except urllib.error.HTTPError as e:
                msg = e.read().decode("utf-8", "ignore")[:300]
                if e.code == 400 and "response_format" in msg and self.json_mode:
                    self.json_mode = False                     # deployment without JSON mode
                    continue
                if e.code in (408, 429, 500, 502, 503, 504):
                    wait = int(e.headers.get("Retry-After", 0) or 0) or 3 * 2 ** attempt
                    time.sleep(min(wait, 60))
                    continue
                warn(f"Azure OpenAI HTTP {e.code}: {msg}")
                break
            except (urllib.error.URLError, TimeoutError, ValueError, KeyError) as e:
                warn(f"Azure OpenAI call failed ({type(e).__name__}: {str(e)[:150]}) - retrying")
                time.sleep(3 * 2 ** attempt)
        self.errors += 1
        return None


AI_SYSTEM = ("You assist an IT auditor testing a SOX ITGC periodic user access review. You read one e-mail "
             "a manager wrote in reply to an access review request and work out exactly what the manager "
             "decided for each user. Understand the meaning, not just keywords, but never infer a decision "
             "the e-mail does not make. Reply with a single JSON object and nothing else.")

AI_PROMPT = """Application: {app}. The control owner sent this manager the list of their direct reports' access
(elevated access codes: {elevated}) and asked them to confirm whether each user's access is still appropriate,
or to say what must be removed.

Manager: {manager}
Users in the list sent to this manager (user_code | user_name):
{users}

Reply written by {author} on {sent}:
\"\"\"
{text}
\"\"\"

Read the reply the way an experienced auditor would:
- Approval can be worded many ways: "approve", "approved", "I confirm", "confirmed", typos such as "comfirm",
  "should retain / keep the same access", "access is appropriate", "no changes needed", "all fine".
- Removal can be worded many ways: "disable", "remove", "revoke", "no longer needs access", "has left",
  "not required", "can be deactivated".
- Ignore greetings, apologies ("Sorry, I missed responding"), explanations of why access is needed, and the
  e-mail signature. A name in the signature, e.g. "Jane Doe / Service Delivery Manager", is the person
  writing, not a user being decided on.
- A bare approval that names nobody ("Approve", "I approve", "All good") applies to every listed user.
- If the reply only names particular users, decide only for those users. Do not extend it to unnamed users.
- A general approval with named exceptions ("Approve. X has left, please disable") approves everyone except X.
- Names may be partial, unaccented or misspelled ("Helene" for "Hélène Martin", "Rob" for
  "Robert Kwan"). Match to the list when the match is clear.
- A question, an out-of-office, a forward with no decision, or "I will review" is not a decision.

Return JSON with exactly these keys:
{{
  "is_decision": true or false,
  "decision": "approve_all" | "approve_some" | "remove_some" | "approve_with_removals" | "remove_all" | "unclear" | "no_decision",
  "applies_to_all_listed_users": true when the reply gives a general decision for the whole list, even if it also
      names exceptions (e.g. "Approve, but disable Kima" = general approve + Kima remove); false when it only
      decides for named users,
  "users": [{{"user_code": "code from the list", "name_in_email": "as written", "decision": "approve" or "remove",
             "scope": "all access, or the specific roles/codes named", "quote": "words copied exactly from the reply"}}],
  "not_matched": ["names in the reply that are not in the list"],
  "conditions": "any condition, exception or follow-up the manager states, else empty string",
  "decision_quote": "the sentence(s) copied exactly from the reply that state the decision",
  "summary": "one short plain-English sentence"
}}"""


def norm_txt(s):
    return re.sub(r"[^a-z0-9]+", " ", fold(s).lower()).strip()


def quote_ok(quote, text):
    q = norm_txt(quote)
    return bool(q) and q in norm_txt(text)


def _name_tokens(name, exclude=()):
    return {t for t in tokens(name, 3) if t not in COMMON_TOKENS and t not in exclude}


def _empty_read(method):
    return {"method": method, "is_decision": False, "decision": "no_decision", "general": None,
            "users": {}, "unmatched": [], "conditions": "", "decision_quote": "", "decision_quote_ok": False,
            "summary": ""}


def ai_read_reply(ai, manager, users, block):
    """Ask the model for per-user decisions and verify everything it returns against the reply text.
    users = [(user_code, user_name)]. Returns the common read structure, or None if the call failed."""
    text = block["Text"][:6000]
    if not text.strip():
        return _empty_read("AI")
    out = ai.chat(AI_SYSTEM, AI_PROMPT.format(
        app=CONFIG["app_name"], elevated=", ".join(CONFIG["elevated_codes"]), manager=manager,
        users="\n".join(f"{c} | {n}" for c, n in users) or "(none)", author=block["Author"],
        sent=f"{block['Sent']:%d-%b-%Y %H:%M}" if block["Sent"] is not None else "unknown", text=text))
    if not isinstance(out, dict):
        return None
    valid = {str(c).upper(): n for c, n in users}
    reply_toks = set(tokens(block["Text"], 3))
    res = _empty_read("AI")
    res.update({"is_decision": bool(out.get("is_decision")), "decision": str(out.get("decision") or ""),
                "summary": str(out.get("summary") or ""), "conditions": str(out.get("conditions") or ""),
                "decision_quote": str(out.get("decision_quote") or "")})
    res["decision_quote_ok"] = quote_ok(res["decision_quote"], block["Text"])
    if out.get("applies_to_all_listed_users") and res["decision"] in ("approve_all", "approve_with_removals", "remove_all"):
        if res["decision_quote_ok"]:
            res["general"] = "remove" if res["decision"] == "remove_all" else "approve"
        else:
            res["unmatched"].append("general decision (quote not found in the e-mail)")
    for it in out.get("users") or []:
        if not isinstance(it, dict):
            continue
        code = str(it.get("user_code") or "").strip().upper()
        dec = str(it.get("decision") or "").lower()
        q = str(it.get("quote") or "")
        if code not in valid or dec not in ("approve", "remove"):
            res["unmatched"].append(f'{it.get("name_in_email") or code or "?"} (no matching user in the list)')
            continue
        q_ok = quote_ok(q, block["Text"])
        # guard against invented answers: the quote must exist, or the user's name must appear in the reply
        if q_ok or (_name_tokens(valid[code]) & reply_toks):
            res["users"][code] = {"decision": dec, "quote": q, "quote_ok": q_ok, "scope": str(it.get("scope") or "")}
        else:
            res["unmatched"].append(f'{it.get("name_in_email") or valid[code]} (quote/name not found in the e-mail)')
    res["unmatched"] += [f"{n} (not in the list)" for n in out.get("not_matched") or [] if str(n).strip()]
    return res


def kw_read_reply(manager, users, block):
    """Keyword fallback with the same output as ai_read_reply: sentence by sentence, a sentence that contains an
    approval or removal word applies to the listed users it names; an approval naming nobody covers everyone."""
    res = _empty_read("Keyword")
    umap = {str(c).upper(): _name_tokens(n) for c, n in users}
    for raw in re.split(r"(?<=[.!?])\s+|\n", block["Text"]):
        s = raw.strip()
        if not s:
            continue
        low = fold(s).lower()
        is_rem = bool(REMOVE_RE.search(low))
        is_app = bool(APPROVE_RE.search(low))
        if not (is_rem or is_app):
            continue
        st = set(tokens(s, 3))
        named = [c for c, nt in umap.items() if (nt & st) or c.lower() in low]
        if named:
            for c in named:
                res["users"][c] = {"decision": "remove" if is_rem else "approve", "quote": s, "quote_ok": True,
                                   "scope": "all access"}
        elif is_app and not is_rem:
            res["general"] = res["general"] or "approve"
        res["is_decision"] = True
        res["decision_quote"] = res["decision_quote"] or s
    res["decision_quote_ok"] = bool(res["decision_quote"])
    n_rm = sum(1 for u in res["users"].values() if u["decision"] == "remove")
    n_ap = sum(1 for u in res["users"].values() if u["decision"] == "approve")
    res["decision"] = ("approve_all" if res["general"] == "approve" and not n_rm else
                       "approve_with_removals" if (res["general"] or n_ap) and n_rm else
                       "remove_some" if n_rm else "approve_some" if n_ap else "no_decision")
    res["summary"] = {"approve_all": "Approves all listed users", "approve_with_removals": "Approves, with removals",
                      "remove_some": "Asks for removals", "approve_some": "Approves the named users",
                      "no_decision": "No decision found"}[res["decision"]]
    return res


# =============================================================================
# Workbook writing helpers
# =============================================================================
FONT = "Arial"
HDR_FILL = PatternFill("solid", fgColor="1F3864")
SEC_FILL = PatternFill("solid", fgColor="D9E1F2")
RED_FILL = PatternFill("solid", fgColor="F8CBAD")
AMB_FILL = PatternFill("solid", fgColor="FFE699")
THIN = Border(*(Side(style="thin", color="BFBFBF"),) * 4)


def write_table(ws, headers, rows, widths=None, date_fmt=None):
    ws.append(headers)
    for r in rows:
        ws.append([None if (not isinstance(v, str) and pd.isna(v)) else
                   (v.to_pydatetime() if isinstance(v, pd.Timestamp) else v) for v in r])
    for c in ws[1]:
        c.font = Font(name=FONT, bold=True, color="FFFFFF")
        c.fill = HDR_FILL
        c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.freeze_panes = "A2"
    if ws.max_row >= 1:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(ws.max_row, 2)}"
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font = Font(name=FONT, size=10)
    for i, h in enumerate(headers, 1):
        w = (widths or {}).get(h)
        if w is None:
            sample = [len(str(h))] + [len(sval(ws.cell(r, i).value)) for r in range(2, min(ws.max_row, 300) + 1)]
            w = min(max(sample) + 2, 45)
        ws.column_dimensions[get_column_letter(i)].width = max(w, 9)
    for col, fmt in (date_fmt or {}).items():
        if col in headers:
            j = headers.index(col) + 1
            for r in range(2, ws.max_row + 1):
                ws.cell(r, j).number_format = fmt


def col_of(headers):
    return {h: get_column_letter(i + 1) for i, h in enumerate(headers)}


# =============================================================================
# Main
# =============================================================================
REQUIRED_FILES = ("user_report_xlsx", "review_xlsx", "after_xlsx")


def run_files(files, output, options=None):
    """Run the testing on explicitly supplied files (any file names).

    files   : user_report_xlsx, review_xlsx, after_xlsx (required);
              user_report_pdf, after_pdf, removed_xlsx, steps_doc, emails_dir (optional)
    options : due_date, start_date, after_date (YYYY-MM-DD), use_ai, cache_path, app_name,
              control_owner, status_approve, status_remove, elevated_codes, dormant_days, response_days
    Writes the Excel workbook to `output` and returns computed results for the UI / chat."""
    options = dict(options or {})
    missing = [k for k in REQUIRED_FILES if not files.get(k)]
    if missing:
        raise ValueError(f"missing required input(s): {', '.join(missing)}")
    saved = {k: (list(v) if isinstance(v, list) else v) for k, v in CONFIG.items()}
    try:
        for k in ("app_name", "control_owner", "status_approve", "status_remove", "elevated_codes", "dormant_days",
                  "response_days"):
            if options.get(k) not in (None, "", []):
                CONFIG[k] = options[k]
        CONFIG["dormant_days"] = int(CONFIG["dormant_days"])
        CONFIG["response_days"] = int(CONFIG["response_days"])
        CONFIG["elevated_codes"] = [str(c).strip().upper() for c in CONFIG["elevated_codes"] if str(c).strip()]
        return _run(files, output, options)
    finally:
        CONFIG.clear()
        CONFIG.update(saved)


def run(input_dir, due_override=None, output=None, use_ai=True, window=None):
    """Command-line entry: find the evidence files in a folder by name pattern."""
    log(f"Input folder: {input_dir}")
    files = {k: find_file(input_dir, k, required=k in REQUIRED_FILES) for k in CONFIG["patterns"]}
    ed = os.path.join(input_dir, CONFIG["emails_dir"])
    files["emails_dir"] = ed if os.path.isdir(ed) else None
    res = run_files(files, output or os.path.join(input_dir, CONFIG["output"]),
                    {"due_date": due_override, "use_ai": use_ai, "response_days": window,
                     "cache_path": os.path.join(input_dir, ".uar_ai_cache.json")})
    return res["output"]


def _run(files, output, options):
    A, R = CONFIG["status_approve"], CONFIG["status_remove"]
    use_ai = options.get("use_ai", True)
    f_ur, f_ur_pdf = files["user_report_xlsx"], files.get("user_report_pdf")
    f_al, f_al_pdf = files["after_xlsx"], files.get("after_pdf")
    f_rv, f_rm, f_steps = files["review_xlsx"], files.get("removed_xlsx"), files.get("steps_doc")
    emails_dir = files.get("emails_dir")
    log(f"Application: {CONFIG['app_name']}")

    # ---------------- read inputs ----------------
    log("Reading the start listing, review file and post-review listing ...")
    ur_info, al_info, rv_info = {}, {}, {}
    ur = standardise_report(read_table(f_ur, "User code")[0], "Start listing", ur_info)
    al = standardise_report(read_table(f_al, "User code")[0], "Post-review listing", al_info)
    rv = standardise_report(read_table(f_rv, "User code")[0], "Review file", rv_info)
    for c in ["Manager", "Access Confirmed", "Access Removed", "Active user"]:
        if c not in rv.columns:
            warn(f"review file has no '{c}' column")
            rv[c] = ""

    rm = pd.DataFrame(columns=["UserID"])
    if f_rm:
        rm, _ = read_table(f_rm, "UserID", sheet=CONFIG["removed_sheet"])
        rm = rm[[c for c in rm.columns if not c.lower().startswith("unnamed")]]
        rm = rm.apply(lambda s: s.map(sval))
        rm = rm[rm["UserID"] != ""].reset_index(drop=True)

    log("Checking the PDF reports (row count and 'Created on' timestamp) ...")
    p1, p2 = pdf_check(f_ur_pdf), pdf_check(f_al_pdf)

    def pick_date(opt, path, pdfres, label, xl_info):
        if opt:
            return pd.Timestamp(opt).normalize(), "entered by user"
        d = date_from_name(path)
        if d is not None:
            return d, f"file name: {os.path.basename(path)}"
        if xl_info.get("created") is not None:
            return xl_info["created"].normalize(), "Excel report 'Created on' footer"
        d = parse_dt(pdfres["created"]) if pdfres.get("created") not in ("", "not found", None) else None
        if d is not None:
            return d.normalize(), "PDF 'Created on' timestamp"
        warn(f"{label}: no date in the file name or PDF - using today's date")
        return pd.Timestamp.today().normalize(), "not found - today's date used"

    extract_dt, extract_src = pick_date(options.get("start_date"), f_ur, p1, "Start listing date", ur_info)
    after_dt, after_src = pick_date(options.get("after_date"), f_al, p2, "Post-review listing date", al_info)

    log("Reading manager e-mails ...")
    blocks, subjects, n_files, req = (read_email_blocks(emails_dir) if emails_dir and os.path.isdir(emails_dir)
                                      else ([], [], 0, pd.DataFrame()))
    if not n_files:
        warn("no manager e-mails supplied - every manager will show as 'No response'")
    else:
        log(f"  {n_files} e-mail files, {len(blocks)} distinct messages")

    due = (pd.Timestamp(options["due_date"]) if options.get("due_date") else
           pd.Timestamp(CONFIG["due_date"]) if CONFIG["due_date"] else
           due_from_subjects(subjects, extract_dt.year))
    due_src = "entered by user" if (options.get("due_date") or CONFIG["due_date"]) else "request e-mail subject"
    if due is None:
        due_src = "not found"
        log("  Due date not stated in e-mail subjects (timeliness uses the response window)")

    # ---------------- request e-mails: the population sent to each manager ----------------
    managers = sorted(m for m in rv["Manager"].unique() if m)

    def canon_mgr(name):
        """Map a name taken from an e-mail onto the review file's spelling of that manager."""
        name = display_name(name) if ("<" in str(name) or "@" in str(name) or "," in str(name)) else str(name or "").strip()
        if not name:
            return ""
        toks = person_tokens(name)
        return next((m for m in managers if same_person(m, toks)), name)

    have_req = len(req) > 0
    if have_req:
        req["Manager (e-mail)"] = [canon_mgr(t) if str(t).strip() else canon_mgr(to)
                                   for t, to in zip(req["Table Manager"], req["Request To"])]
        req["Manager (To:)"] = req["Request To"].map(canon_mgr)
        req = req.sort_values("Request Sent", na_position="last") \
                 .drop_duplicates(["Manager (e-mail)", "Key"], keep="first").reset_index(drop=True)
        log(f"  Request e-mails: {req['Request File'].nunique()} with a pasted list, {len(req)} access rows, "
            f"{req['Manager (e-mail)'].nunique()} managers")
    else:
        warn("no pasted access list found in the e-mails - the e-mail row-count check is skipped")
        req = pd.DataFrame(columns=["Key", "User code", "User name", "Type of membership", "Name", "Code",
                                    "Request File", "Request To", "Request From", "Request Sent", "Table Manager",
                                    "Manager (e-mail)", "Manager (To:)"])
    req_mgrs_of_key = req.groupby("Key")["Manager (e-mail)"].agg(lambda s: sorted(set(s))).to_dict() if have_req else {}

    # request date per manager: oldest request e-mail with a list, else oldest request block addressed to them
    req_date, req_src = {}, {}
    for mgr in managers:
        d = req.loc[req["Manager (e-mail)"] == mgr, "Request Sent"].dropna()
        if len(d):
            req_date[mgr], req_src[mgr] = d.min(), "request e-mail"
            continue
        cands = [b["Sent"] for b in blocks if b["Class"] == "Request / reminder" and b["Sent"] is not None
                 and b["To"] and same_person(mgr, person_tokens(b["To"]))]
        if cands:
            req_date[mgr], req_src[mgr] = min(cands), "request e-mail (no list found)"
        else:
            req_date[mgr], req_src[mgr] = extract_dt, "start listing date (no request e-mail found)"

    # ---------------- manager replies ----------------
    owner_toks = person_tokens(CONFIG["control_owner"])
    owner_core = set(tokens(CONFIG["control_owner"]))
    for b in blocks:
        fm = canon_mgr(b["File Manager"]) if b.get("File Manager") else ""
        if fm in managers:
            b["Matched manager"], b["Match basis"] = fm, "manager in the e-mailed list"
        else:
            b["Matched manager"] = next((m for m in managers if same_person(m, b["_toks"])), "")
            b["Match basis"] = "sender name" if b["Matched manager"] else ""
    mgr_blocks = {}
    for mgr in managers:
        mgr_blocks[mgr] = sorted(
            [b for b in blocks if b["Matched manager"] == mgr and b["Class"] != "Request / reminder"
             and not (owner_core and owner_core <= b["_toks"] and not same_person(mgr, owner_toks))],
            key=lambda b: b["Sent"] if b["Sent"] is not None else pd.Timestamp.max)

    users_of = {}
    for mgr in managers:
        u = rv.loc[rv["Manager"] == mgr, ["User code", "User name"]]
        if have_req:
            u = pd.concat([u, req.loc[req["Manager (e-mail)"] == mgr, ["User code", "User name"]]])
        users_of[mgr] = u.assign(_k=u["User code"].str.upper()).drop_duplicates("_k")[["User code", "User name"]].values.tolist()

    ai = AzureAI(options.get("cache_path") or os.path.join(os.path.dirname(f_rv), ".uar_ai_cache.json")) if use_ai else None
    use_ai_now = bool(ai and ai.enabled)
    if use_ai and not use_ai_now:
        warn("Azure OpenAI not configured (.env / environment variables) - using keyword rules")
    if use_ai_now:
        log(f"  Reading {sum(len(v) for v in mgr_blocks.values())} manager replies with Azure OpenAI "
            f"(deployment {ai.deployment}) ...")
    n_ai = n_kw_fallback = 0
    for mgr in managers:
        for b in mgr_blocks[mgr]:
            r_ = ai_read_reply(ai, mgr, users_of[mgr], b) if use_ai_now else None
            if r_ is None:
                r_ = kw_read_reply(mgr, users_of[mgr], b)
                n_kw_fallback += int(use_ai_now)
            else:
                n_ai += 1
            b["Read"] = r_
    if use_ai_now:
        log(f"  AI: {ai.calls} new calls, {ai.hits} from cache, {ai.errors} failed")
        method = (f"Azure OpenAI ({ai.deployment}); every answer checked against the e-mail text"
                  if not n_kw_fallback else
                  f"Azure OpenAI ({ai.deployment}) for {n_ai} replies; keyword rules for {n_kw_fallback} (AI call failed)")
    else:
        method = "Keyword rules"

    # response used for timeliness = first reply that contains a decision; decisions = all replies in date order
    resp, final_user, final_general, general_quote, mgr_unmatched, mgr_method = {}, {}, {}, {}, {}, {}
    for mgr in managers:
        own = mgr_blocks[mgr]
        decisive = [b for b in own if b["Read"]["is_decision"]]
        if decisive or own:
            resp[mgr] = (decisive or own)[0]
        fu, fg, gq, un = {}, None, "", []
        for b in own:
            rd = b["Read"]
            if rd["general"]:
                fg, gq = rd["general"], rd["decision_quote"]
                fu = {k: v for k, v in fu.items() if v["decision"] != fg}   # a later general reply supersedes
            for code, it in rd["users"].items():
                fu[code] = {**it, "file": b["File"], "sent": b["Sent"]}
            un += rd["unmatched"]
        final_user[mgr], final_general[mgr], general_quote[mgr], mgr_unmatched[mgr] = fu, fg, gq, un
        mgr_method[mgr] = "/".join(sorted({b["Read"]["method"] for b in own})) if own else ""

    # ---------------- per-row testing (all Python) ----------------
    window = int(options.get("response_days") or CONFIG["response_days"])
    yn = {True: "Yes", False: "No"}
    ukeys, rkeys, akeys = set(ur["Key"]), set(rv["Key"]), set(al["Key"])
    rm_ids = {str(x).upper() for x in rm["UserID"]} if "UserID" in rm.columns else set()
    have_rm = bool(f_rm)
    elev = {c.upper() for c in CONFIG["elevated_codes"]}

    def status(r):
        rem, conf = r["Access Removed"].lower(), r["Access Confirmed"].lower()
        if rem and rem not in ("n", "no", "0", "-"):
            return R
        if conf and conf not in ("n", "no", "0", "-"):
            return A
        return "No decision"

    def timeliness(mgr):
        b = resp.get(mgr)
        if b is None or b["Sent"] is None:
            return None, None, "No response"
        days = (b["Sent"].normalize() - req_date[mgr].normalize()).days
        return b["Sent"], days, ("Yes" if days <= window else f"No - {days} days")

    calc = rv.copy()
    calc["Status"] = calc.apply(status, axis=1)
    tl = {m: timeliness(m) for m in managers}
    calc["Request Sent"] = calc["Manager"].map(lambda m: req_date.get(m))
    calc["Response Date"] = calc["Manager"].map(lambda m: tl[m][0] if m in tl else None)
    calc["Days to Respond"] = whole_days(calc["Manager"].map(lambda m: tl[m][1] if m in tl else None))
    calc[f"Responded within {window} days"] = calc["Manager"].map(lambda m: tl[m][2] if m in tl else "No response")
    calc["Self-approval check"] = ["Self-approval" if (m and str(u).strip().lower() == str(m).strip().lower())
                                   else "No" for u, m in zip(calc["User name"], calc["Manager"])]

    def xcheck(r):
        mgr, code = r["Manager"], r["User code"].upper()
        if mgr not in resp:
            return "No manager e-mail", ""
        it = final_user[mgr].get(code)
        eff = it["decision"] if it else final_general[mgr]
        how = f" ({mgr_method[mgr]})"
        note = (f'{mgr_method[mgr]}: {it["decision"]} ({it["scope"] or "all access"}) - "{it["quote"]}"'
                + ("" if it["quote_ok"] else " [exact quote not found - name found in e-mail, verify]")) if it else \
               (f'{mgr_method[mgr]}: general {final_general[mgr]} - "{general_quote[mgr]}"'
                if final_general[mgr] else "")
        if r["Status"] == R:
            txt = "Supported - manager asked to remove" if eff == "remove" else \
                  "NOT supported - manager did not ask to remove this user"
        elif r["Status"] == A:
            txt = ("Conflict - manager asked to remove this user" if eff == "remove" else
                   "Consistent - manager approved" if eff == "approve" else
                   "NOT addressed - manager's reply does not cover this user")
        else:
            txt = ("E-mail asks removal; review file has no decision" if eff == "remove" else
                   "E-mail approves; review file has no decision" if eff == "approve" else "No decision in review file")
        return txt + how, note

    xc = calc.apply(xcheck, axis=1)
    calc["Email cross-check"] = [x[0] for x in xc]
    calc["Comments"] = [x[1] for x in xc]
    calc["Feedback Mail"] = calc["Manager"].map(lambda m: resp[m]["File"] if m in resp else "")
    calc["In Request E-mail"] = (calc["Key"].map(lambda k: yn[k in req_mgrs_of_key]) if have_req
                                 else "No request list found")
    calc["Manager per Request E-mail"] = calc["Key"].map(lambda k: ", ".join(req_mgrs_of_key.get(k, [])))
    calc["Elevated Access"] = calc["Code"].str.upper().isin(elev).map(yn)
    lls = calc["Last login"] if "Last login" in calc.columns else pd.Series([None] * len(calc), index=calc.index)
    calc["Days Since Last Login"] = whole_days([None if (x is None or pd.isna(x))
                                                else (extract_dt.normalize() - pd.Timestamp(x).normalize()).days
                                                for x in lls], calc.index)
    calc["Dormant"] = ["Never logged in" if (d is None or pd.isna(d)) else yn[d > CONFIG["dormant_days"]]
                       for d in calc["Days Since Last Login"]]
    calc["In After Listing"] = calc["Key"].isin(akeys).map(yn)
    calc["Post Review Match"] = [
        ("Removed - OK" if ia == "No" else "EXCEPTION - not removed") if st_ == R else
        ("Retained - OK" if ia == "Yes" else "Not in after listing - check") if st_ == A else "No decision"
        for st_, ia in zip(calc["Status"], calc["In After Listing"])]
    calc["In Client Removed-Accounts Analysis"] = (calc["User code"].str.upper().isin(rm_ids).map(yn) if have_rm
                                                   else "Not supplied")
    calc["In Start User Report"] = calc["Key"].isin(ukeys).map(yn)

    def exc_detail(r):
        d = []
        if r["In Start User Report"] == "No":
            d.append("Not in start User Report")
        if r["Manager"] == "":
            d.append("No manager assigned")
        if have_req and r["In Request E-mail"] == "No":
            d.append("Not in the list e-mailed to any manager")
        if have_req and r["Manager per Request E-mail"] and r["Manager"] and \
                r["Manager"] not in r["Manager per Request E-mail"].split(", "):
            d.append("Review-file manager differs from request e-mail")
        if r["Status"] == "No decision":
            d.append("No reviewer decision")
        tl_ = r[f"Responded within {window} days"]
        if tl_ == "No response":
            d.append("No manager response")
        elif tl_.startswith("No - "):
            d.append(f"Response after {window} days")
        if r["Post Review Match"] == "EXCEPTION - not removed":
            d.append("Rejected access not removed")
        if r["Self-approval check"] == "Self-approval":
            d.append("Self-approval")
        xc_ = str(r["Email cross-check"])
        if xc_.startswith("NOT supported"):
            d.append("Rejection not supported by e-mail")
        if xc_.startswith("NOT addressed"):
            d.append("User not covered by manager's reply")
        if xc_.startswith("Conflict"):
            d.append("E-mail asks removal but access approved")
        if have_rm and r["Status"] == R and r["In Client Removed-Accounts Analysis"] == "No":
            d.append("Not in client removed-accounts analysis")
        return "; ".join(d)

    calc["Exception Detail"] = calc.apply(exc_detail, axis=1)
    calc["Exception"] = calc["Exception Detail"].map(lambda x: "Yes" if x else "No")

    # ---------------- e-mail row counts per manager ----------------
    start_by_user = ur.assign(_u=ur["User code"].str.upper()).groupby("_u")["Key"].apply(set).to_dict()
    cnt_rows = []
    for mgr in sorted(set(managers) | (set(req["Manager (e-mail)"]) - {""} if have_req else set(managers))):
        sub = req[req["Manager (e-mail)"] == mgr] if have_req else req.iloc[0:0]
        e_keys = set(sub["Key"])
        e_users = set(sub["User code"].str.upper())
        s_keys = set().union(*[start_by_user.get(u, set()) for u in e_users]) if e_users else set()
        rv_n = int((rv["Manager"] == mgr).sum())
        miss, extra_ = sorted(s_keys - e_keys), sorted(e_keys - ukeys)
        to_names = sorted(set(sub["Manager (To:)"]) - {""})
        tbl_names = sorted({canon_mgr(x) for x in sub["Table Manager"] if str(x).strip()})
        cnt_rows.append({
            "Manager": mgr, "Request E-mail": ", ".join(sorted(set(sub["Request File"]))) or "Not found",
            "Request Sent": sub["Request Sent"].min() if len(sub) else None,
            "Manager per To: line": ", ".join(to_names), "Manager per pasted list": ", ".join(tbl_names),
            "Users in E-mail": len(e_users), "Rows in E-mail": len(e_keys),
            "Rows in Start User Report (same users)": len(s_keys), "Rows in Review File": rv_n,
            "Difference (E-mail - Start Report)": len(e_keys) - len(s_keys) if len(sub) else None,
            "Rows Match Start Report": ("No request list" if not len(sub) else
                                        yn[e_keys == s_keys and not extra_]),
            "Start-Report Rows Missing from E-mail": len(miss), "E-mail Rows Not in Start Report": len(extra_),
            "Missing Rows (first 10)": "; ".join(miss[:10]),
            "To: and List Agree": ("" if not len(sub) else
                                   yn[not to_names or not tbl_names or set(to_names) == set(tbl_names)]),
        })
    cnt_df = pd.DataFrame(cnt_rows)
    not_sent = ur[~ur["Key"].isin(set(req["Key"]))] if have_req else ur.iloc[0:0]

    # ---------------- manager level ----------------
    mgr_rows = []
    for mgr in managers:
        sub = calc[calc["Manager"] == mgr]
        b = resp.get(mgr)
        rd = b["Read"] if b is not None else None
        rej_codes = set(sub.loc[sub["Status"] == R, "User code"].str.upper())
        rm_codes = {c for c, it in final_user[mgr].items() if it["decision"] == "remove"}
        by_mgr = "" if b is None else ("Yes" if same_surname(mgr, b["_toks"]) else f"No - sent by {display_name(b['Author'])}")
        if b is None:
            cons = "No e-mail"
        else:
            parts = ([f"file rejects {', '.join(sorted(rej_codes - rm_codes))} not named for removal"]
                     if rej_codes - rm_codes else []) + \
                    ([f"e-mail removes {', '.join(sorted(rm_codes - rej_codes))} not rejected in file"]
                     if rm_codes - rej_codes else []) + \
                    ([f"{int(sub['Email cross-check'].str.startswith('NOT addressed').sum())} rows not covered by the reply"]
                     if sub["Email cross-check"].str.startswith("NOT addressed").any() else []) + \
                    ([f"unmatched: {'; '.join(mgr_unmatched[mgr])}"] if mgr_unmatched[mgr] else []) + \
                    ([f"reply {by_mgr[5:]}, not by the manager"] if by_mgr.startswith("No") else [])
            cons = ("Check - " + " | ".join(parts)) if parts else "Consistent"
        sent, days, ontime = tl[mgr]
        mgr_rows.append({
            "Manager": mgr, "Users": sub["User code"].nunique(), "Access Rows": len(sub),
            "Approved Rows": int((sub["Status"] == A).sum()), "Rejected Rows": int((sub["Status"] == R).sum()),
            "No-decision Rows": int((sub["Status"] == "No decision").sum()),
            "Request Sent": req_date[mgr], "Request Date Source": req_src[mgr],
            "Response Found": yn[b is not None], "Response Date": sent, "Days to Respond": days,
            f"Responded within {window} days": ontime,
            "Responded By": display_name(b["Author"]) if b is not None else "", "Reply Sent by Manager": by_mgr,
            "Matched Using": b.get("Match basis", "") if b is not None else "",
            "Feedback Mail": b["File"] if b is not None else "",
            "Reading Method": mgr_method[mgr], "Decision": rd["decision"] if rd else "",
            "Summary": rd["summary"] if rd else "", "Decision Quote": rd["decision_quote"] if rd else "",
            "Decision Quote Verified": ("" if not rd else yn[rd["decision_quote_ok"]]),
            "Users To Remove (per e-mail)": ", ".join(sorted(rm_codes)),
            "Users Approved by Name": ", ".join(sorted(c for c, it in final_user[mgr].items() if it["decision"] == "approve")),
            "General Decision": final_general[mgr] or "",
            "Names Not Matched / Not Verified": "; ".join(mgr_unmatched[mgr]),
            "Conditions": rd["conditions"] if rd else "",
            "E-mail vs Review File": cons, "Exception Rows": int((sub["Exception"] == "Yes").sum()),
            "Replies Found": len(mgr_blocks[mgr]), "Manager Reply Text": (b["Text"][:1500] if b is not None else ""),
        })
    mgr_df = pd.DataFrame(mgr_rows)
    if len(mgr_df):
        mgr_df["Days to Respond"] = whole_days(mgr_df["Days to Respond"], mgr_df.index)

    # ---------------- removals (user level) ----------------
    rej_up = set(calc.loc[calc["Status"] == R, "User code"].str.upper())
    gone_ = set(ur["User code"].str.upper()) - set(al["User code"].str.upper())
    after_cnt = al["User code"].str.upper().value_counts()
    start_cnt = ur["User code"].str.upper().value_counts()
    rm_by = {str(r["UserID"]).upper(): r for _, r in rm.iterrows()} if have_rm and len(rm) else {}
    chk3 = next((c for c in rm.columns if c.lower().startswith("check #3")), None)
    rem_rows = []
    for uc in sorted(rej_up | rm_ids | gone_):
        cu = calc[calc["User code"].str.upper() == uc]
        F, L = int((cu["Status"] == R).sum()), int(after_cnt.get(uc, 0))
        if F == 0:
            ver = "Removed outside review - check leaver/ticket" if L == 0 else "Not rejected - access retained"
        else:
            ver = ("Yes - rejected access removed" if not cu["Post Review Match"].str.startswith("EXCEPTION").any()
                   else "EXCEPTION - rejected access still present")
        src = (["Rejected in review"] if uc in rej_up else []) + \
              (["Client removed-accounts analysis"] if uc in rm_ids else []) + \
              (["Absent from After Listing"] if uc in gone_ else [])
        nm = (cu["User name"].iloc[0] if len(cu) else
              (ur.loc[ur["User code"].str.upper() == uc, "User name"].head(1).tolist() or
               [rm_by[uc].get("User Name", "") if uc in rm_by else ""])[0])
        rem_rows.append({"User code": uc, "User name": nm, "Manager": cu["Manager"].iloc[0] if len(cu) else "",
                         "Source": "; ".join(src), "Rows in Start Report": int(start_cnt.get(uc, 0)),
                         "Rows Rejected": F, "Email cross-check": cu["Email cross-check"].iloc[0] if F else "",
                         "In Client Removed-Accounts Analysis": (yn[uc in rm_ids] if have_rm else "Not supplied"),
                         "Client: Change Details": rm_by[uc].get("Change Details", "") if uc in rm_by else "",
                         "Client: Elevated Access?": rm_by[uc].get(chk3, "") if (uc in rm_by and chk3) else "",
                         "Rows in After Listing": L, "Removal Verified": ver})
    rem_df = pd.DataFrame(rem_rows, columns=["User code", "User name", "Manager", "Source", "Rows in Start Report",
                                             "Rows Rejected", "Email cross-check", "In Client Removed-Accounts Analysis",
                                             "Client: Change Details", "Client: Elevated Access?",
                                             "Rows in After Listing", "Removal Verified"])

    # ---------------- completeness + movement ----------------
    cc = ["User code", "User name", "Type of membership", "Code"]
    comp_df = pd.DataFrame(
        [{"Issue": "Missing from review file", **{k: r[k] for k in cc}} for _, r in ur[~ur["Key"].isin(rkeys)].iterrows()] +
        [{"Issue": "Extra in review file", **{k: r[k] for k in cc}} for _, r in rv[~rv["Key"].isin(ukeys)].iterrows()] +
        [{"Issue": "No manager in review file", **{k: r[k] for k in cc}} for _, r in rv[rv["Manager"] == ""].iterrows()] +
        [{"Issue": "Not in the list e-mailed to any manager", **{k: r[k] for k in cc}} for _, r in not_sent.iterrows()],
        columns=["Issue"] + cc)
    st_map = dict(zip(calc["Key"], calc["Status"]))
    mov_df = pd.DataFrame(
        [{"Movement": "Removed after review", **{k: r[k] for k in cc}, "Review Status": st_map.get(r["Key"], "Not in review"),
          "Explained by Review": "Yes" if st_map.get(r["Key"]) == R else "No - check leaver/ticket"}
         for _, r in ur[~ur["Key"].isin(akeys)].iterrows()] +
        [{"Movement": "Added after review", **{k: r[k] for k in cc}, "Review Status": "n/a",
          "Explained by Review": "New access - outside this review (provisioning control)"}
         for _, r in al[~al["Key"].isin(ukeys)].iterrows()],
        columns=["Movement"] + cc + ["Review Status", "Explained by Review"])

    # ---------------- IPE ----------------
    def ts_ok(created, expected, src=""):
        if src.startswith("PDF"):
            return "n/a - extraction date taken from the PDF"
        d = parse_dt(created) if created and created != "not found" else None
        if d is None or expected is None:
            return "Verify manually"
        return "Yes" if d.normalize() == expected.normalize() else f"No - PDF dated {d:%m-%d-%Y}"

    def xl_ok(info, expected, src=""):
        if src.startswith("Excel"):
            return "n/a - extraction date taken from this footer"
        d = info.get("created")
        if d is None:
            return "Verify manually - no 'Created on' footer"
        return "Yes" if d.normalize() == expected.normalize() else f"No - report dated {d:%m-%d-%Y}"

    def rc_txt(pdf_rows, xl_rows):
        return "Verify manually" if not isinstance(pdf_rows, int) else \
            ("Yes" if pdf_rows == xl_rows else f"No - difference {pdf_rows - xl_rows}")

    rowcount = rc_txt(p1["rows"], len(ur))
    ipe_df = pd.DataFrame([
        {"Check": "Excel file", "Start listing": os.path.basename(f_ur), "Post-review listing": os.path.basename(f_al)},
        {"Check": "Excel data rows", "Start listing": len(ur), "Post-review listing": len(al)},
        {"Check": "Excel unique users", "Start listing": int(ur["User code"].nunique()),
         "Post-review listing": int(al["User code"].nunique())},
        {"Check": "Extraction date used", "Start listing": f"{extract_dt:%m-%d-%Y} ({extract_src})",
         "Post-review listing": f"{after_dt:%m-%d-%Y} ({after_src})"},
        {"Check": "Excel footer lines ignored", "Start listing": ur_info.get("footer_rows", 0),
         "Post-review listing": al_info.get("footer_rows", 0)},
        {"Check": "Excel report 'Created on' (footer)", "Start listing": ur_info.get("created_text", "not found"),
         "Post-review listing": al_info.get("created_text", "not found")},
        {"Check": "Excel 'Created on' date = extraction date",
         "Start listing": xl_ok(ur_info, extract_dt, extract_src), "Post-review listing": xl_ok(al_info, after_dt, after_src)},
        {"Check": "PDF file", "Start listing": os.path.basename(f_ur_pdf) if f_ur_pdf else "not supplied",
         "Post-review listing": os.path.basename(f_al_pdf) if f_al_pdf else "not supplied"},
        {"Check": "PDF pages", "Start listing": p1["pages"], "Post-review listing": p2["pages"]},
        {"Check": "PDF 'Created on' timestamp", "Start listing": p1["created"], "Post-review listing": p2["created"]},
        {"Check": "PDF date = extraction date", "Start listing": ts_ok(p1["created"], extract_dt, extract_src),
         "Post-review listing": ts_ok(p2["created"], after_dt, after_src)},
        {"Check": "PDF rows (user-code occurrences)", "Start listing": p1["rows"], "Post-review listing": p2["rows"]},
        {"Check": "Row count PDF = Excel", "Start listing": rowcount, "Post-review listing": rc_txt(p2["rows"], len(al))},
        {"Check": "Extraction steps document", "Start listing": os.path.basename(f_steps) if f_steps else "not supplied",
         "Post-review listing": ""},
        {"Check": "PDF read note", "Start listing": p1["note"], "Post-review listing": p2["note"]},
    ])

    # ---------------- summary (values) ----------------
    n_unver = (sum(1 for m in managers if resp.get(m) is not None and resp[m]["Read"]["method"] == "AI"
                   and resp[m]["Read"]["is_decision"] and not resp[m]["Read"]["decision_quote_ok"])
               + int(calc["Comments"].str.contains("not found").sum())
               + sum(1 for m in managers if mgr_unmatched[m]))
    n_exc = int((calc["Exception"] == "Yes").sum())
    has_req_mgr = cnt_df["Rows Match Start Report"] != "No request list" if len(cnt_df) else pd.Series(dtype=bool)
    W = f"Responded within {window} days"
    sections = [
        ("1. Report extraction (IPE)", [
            ("Access rows - start User Report (Excel)", len(ur)),
            ("Unique users - start User Report", int(ur["User code"].nunique())),
            ("Row count agrees to PDF", rowcount),
            ("PDF 'Created on' date agrees to extraction date", ts_ok(p1["created"], extract_dt, extract_src)),
            ("Excel report 'Created on' agrees to extraction date", xl_ok(ur_info, extract_dt, extract_src)),
            ("Footer lines ignored in the listings", ur_info.get("footer_rows", 0) + al_info.get("footer_rows", 0)
             + rv_info.get("footer_rows", 0)),
        ]),
        ("2. Completeness of review population", [
            ("Rows in review file", len(rv)),
            ("Rows missing from review file", int((comp_df["Issue"] == "Missing from review file").sum())),
            ("Rows in review file not in User Report", int((comp_df["Issue"] == "Extra in review file").sum())),
            ("Rows with no manager assigned", int((comp_df["Issue"] == "No manager in review file").sum())),
            ("Managers with a request e-mail list", int(has_req_mgr.sum()) if len(cnt_df) else 0),
            ("Access rows in request e-mails", int(len(req))),
            ("Managers whose e-mailed rows match the start User Report",
             int((cnt_df["Rows Match Start Report"] == "Yes").sum()) if len(cnt_df) else 0),
            ("Start-report rows not e-mailed to any manager", int(len(not_sent)) if have_req else "No request lists found"),
            ("Rows whose review-file manager differs from the e-mail",
             int(calc["Exception Detail"].str.contains("differs from request e-mail").sum())),
        ]),
        ("3. Manager responses", [
            ("Managers in scope", len(managers)),
            ("Managers who responded", int((mgr_df["Response Found"] == "Yes").sum()) if len(mgr_df) else 0),
            (f"Responses within {window} days of the request", int((mgr_df[W] == "Yes").sum()) if len(mgr_df) else 0),
            (f"Responses after {window} days", int(mgr_df[W].str.startswith("No - ").sum()) if len(mgr_df) else 0),
            ("No response", int((mgr_df[W] == "No response").sum()) if len(mgr_df) else 0),
            ("Replies not sent by the manager themselves (check)",
             int(mgr_df["Reply Sent by Manager"].str.startswith("No").sum()) if len(mgr_df) else 0),
            ("E-mail vs review file needs checking",
             int(mgr_df["E-mail vs Review File"].str.startswith("Check").sum()) if len(mgr_df) else 0),
            ("How e-mails were read", method),
            ("AI answers not verified against the e-mail (check)", n_unver),
        ]),
        ("4. Review decisions", [
            ("Rows approved", int((calc["Status"] == A).sum())),
            ("Rows rejected", int((calc["Status"] == R).sum())),
            ("Rows with no decision", int((calc["Status"] == "No decision").sum())),
            ("Self-approvals (user = manager)", int((calc["Self-approval check"] == "Self-approval").sum())),
            ("Rejections not supported by manager e-mail", int(calc["Email cross-check"].str.startswith("NOT supported").sum())),
            ("Approved rows the e-mail asked to remove", int(calc["Email cross-check"].str.startswith("Conflict").sum())),
            ("Approved rows not covered by the manager's reply",
             int(calc["Email cross-check"].str.startswith("NOT addressed").sum())),
        ]),
        ("5. Post-review (removals)", [
            ("Users with rejected access", int((rem_df["Rows Rejected"] > 0).sum())),
            ("Rejected rows removed in After Listing", int((calc["Post Review Match"] == "Removed - OK").sum())),
            ("Rejected rows still present (exception)", int(calc["Post Review Match"].str.startswith("EXCEPTION").sum())),
            ("Users removed outside the review (check)", int(rem_df["Removal Verified"].str.startswith("Removed outside").sum())),
            ("Rows added after the review (new access)", int((mov_df["Movement"] == "Added after review").sum())),
        ]),
        ("6. Other observations", [
            ("Elevated access rows (" + "/".join(CONFIG["elevated_codes"]) + ")", int((calc["Elevated Access"] == "Yes").sum())),
            ("Dormant / never-logged-in rows approved",
             int(((calc["Status"] == A) & calc["Dormant"].isin(["Yes", "Never logged in"])).sum())),
        ]),
        ("Overall", [
            ("Rows with exceptions", n_exc),
            ("Conclusion", "No exceptions noted" if (n_exc == 0 and rowcount == "Yes" and
                                                     (not have_req or len(not_sent) == 0))
             else "Exceptions / items to verify - see UAR_Analysis, Exception Detail"),
        ]),
    ]
    sheet_of = {"1": "IPE_Check", "2": "Completeness", "3": "Manager_Responses", "4": "UAR_Analysis",
                "5": "Removals_Testing", "6": "UAR_Analysis", "O": "UAR_Analysis"}
    summary = [{"section": t, "metric": lab, "value": v, "sheet": sheet_of[t[0]]} for t, items in sections for lab, v in items]

    uar_cols = ["User code", "User name", "Active user", "Manager", "Last login", "Type of membership", "Name", "Code",
                "Status", "Request Sent", "Response Date", "Days to Respond", W, "Self-approval check", "Comments",
                "Feedback Mail", "In Request E-mail", "Manager per Request E-mail", "Elevated Access",
                "Days Since Last Login", "Dormant", "In After Listing", "Post Review Match",
                "In Client Removed-Accounts Analysis", "Email cross-check", "In Start User Report",
                "Exception Detail", "Exception"]
    uar_df = calc[[c for c in uar_cols if c in calc.columns]].reset_index(drop=True)
    req_out = req[[c for c in ["Manager (e-mail)", "Manager (To:)", "Request File", "Request Sent", "User code",
                               "User name", "Type of membership", "Name", "Code", "Key"] if c in req.columns]].copy()
    if len(req_out):
        req_out["In Start User Report"] = req_out["Key"].isin(ukeys).map(yn)
    email_df = pd.DataFrame([{"File": b["File"], "Message": "oldest" if b["Oldest"] else ("newest" if b["Block"] == 0 else "quoted"),
                              "Author": b["Author"], "To": b["To"], "Sent": b["Sent"], "Subject": b["Subject"],
                              "Classification": b["Class"], "Matched Manager": b["Matched manager"],
                              "Matched Using": b.get("Match basis", ""),
                              "Read By": (b.get("Read") or {}).get("method", ""),
                              "Decision": (b.get("Read") or {}).get("decision", ""),
                              "Summary": (b.get("Read") or {}).get("summary", ""), "Text": b["Text"][:1500]}
                             for b in blocks],
                            columns=["File", "Message", "Author", "To", "Sent", "Subject", "Classification",
                                     "Matched Manager", "Matched Using", "Read By", "Decision", "Summary", "Text"])

    tables = {"uar": uar_df, "email_counts": cnt_df, "managers": mgr_df, "removals": rem_df, "completeness": comp_df,
              "movement": mov_df, "ipe": ipe_df, "request_rows": req_out, "emails": email_df}
    meta = {"app_name": CONFIG["app_name"], "start_date": extract_dt, "start_date_source": extract_src,
            "after_date": after_dt, "after_date_source": after_src, "due_date": due, "due_date_source": due_src,
            "response_days": window, "method": method, "status_approve": A, "status_remove": R,
            "generated": datetime.now(), "email_files": n_files}

    # ---------------- workbook: values only, no formulas ----------------
    write_workbook(output, tables, summary, meta, {"ur": ur, "rv": rv, "al": al, "rm": rm})
    rej_keys = set(calc.loc[calc["Status"] == R, "Key"])
    log(f"  Start report rows: {len(ur)} | review rows: {len(rv)} | after listing rows: {len(al)}")
    log(f"  Managers: {len(managers)} | responses found: {len(resp)} | response window: {window} days")
    log(f"  Approved rows: {(calc['Status'] == A).sum()} | Rejected rows: {len(rej_keys)} | "
        f"Rejected but still in After Listing: {len(rej_keys & akeys)}")
    log(f"Saved workbook: {os.path.basename(output)}")
    return {"output": output, "meta": meta, "summary": summary, "tables": tables}


SHEETS = [("UAR_Analysis", "uar"), ("Email_Row_Count", "email_counts"), ("Manager_Responses", "managers"),
          ("Removals_Testing", "removals"), ("Completeness", "completeness"), ("Movement", "movement"),
          ("IPE_Check", "ipe"), ("Request_Email_Rows", "request_rows"), ("Email_Log", "emails")]
WIDE = {"Exception Detail": 48, "Comments": 50, "Email cross-check": 42, "Feedback Mail": 34, "Manager Reply Text": 60,
        "Text": 70, "Decision Quote": 45, "Summary": 40, "E-mail vs Review File": 45, "Missing Rows (first 10)": 50,
        "Names Not Matched / Not Verified": 40, "Conditions": 35, "Subject": 40, "Source": 38, "Removal Verified": 40}
DATE_COLS = {"Request Sent": "mm-dd-yyyy", "Response Date": "mm-dd-yyyy", "Sent": "mm-dd-yyyy h:mm",
             "Last login": "mmm d, yyyy h:mm AM/PM"}


def whole_days(values, index=None):
    """Integer day counts with blanks kept blank (pandas would otherwise turn them into 15.0 / NaN)."""
    vals = list(values)
    out = [None if (v is None or (not isinstance(v, str) and pd.isna(v))) else int(v) for v in vals]
    return pd.Series(out, index=index if index is not None else getattr(values, "index", None), dtype=object)


def _rows(df):
    return df.values.tolist() if len(df) else []


def write_workbook(path, tables, summary, meta, src):
    """All values are computed in Python; the workbook contains no formulas."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws["A1"] = f"{meta['app_name']} - Periodic User Access Review - Testing Summary"
    ws["A1"].font = Font(name=FONT, bold=True, size=14, color="1F3864")
    ws["A2"] = (f"Generated {meta['generated']:%d-%b-%Y %H:%M}. Start listing {meta['start_date']:%m-%d-%Y}; "
                f"post-review listing {meta['after_date']:%m-%d-%Y}; response window {meta['response_days']} days "
                f"from each manager's request e-mail.")
    ws["A2"].font = Font(name=FONT, italic=True, size=9, color="808080")
    r = 4
    for j, h in enumerate(("Metric", "Result", "Detail sheet"), 1):
        c = ws.cell(r, j, h)
        c.font, c.fill = Font(name=FONT, bold=True, color="FFFFFF"), HDR_FILL
    sec = None
    for s in summary:
        if s["section"] != sec:
            sec = s["section"]
            r += 1
            ws.cell(r, 1, sec).font = Font(name=FONT, bold=True)
            for j in (1, 2, 3):
                ws.cell(r, j).fill = SEC_FILL
        r += 1
        ws.cell(r, 1, s["metric"]).font = Font(name=FONT, size=10)
        v = ws.cell(r, 2, s["value"])
        v.font = Font(name=FONT, size=10, bold=s["metric"] == "Conclusion",
                      color="B42318" if (s["metric"] == "Conclusion" and str(s["value"]).startswith("Exceptions")) else None)
        v.alignment = Alignment(horizontal="left")
        ws.cell(r, 3, s["sheet"]).font = Font(name=FONT, size=10, color="808080")
        for j in (1, 2, 3):
            ws.cell(r, j).border = THIN
    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width, ws.column_dimensions["C"].width = 58, 52, 20
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0

    wp = wb.create_sheet("Parameters")
    params = [("Parameter", "Value"), ("Application", meta["app_name"]),
              ("Start listing date", f"{meta['start_date']:%m-%d-%Y} ({meta['start_date_source']})"),
              ("Post-review listing date", f"{meta['after_date']:%m-%d-%Y} ({meta['after_date_source']})"),
              ("Response window", f"{meta['response_days']} days from each manager's request e-mail"),
              ("Due date stated in request e-mail",
               f"{meta['due_date']:%m-%d-%Y} ({meta['due_date_source']})" if meta["due_date"] is not None else "not found"),
              ("Dormant threshold (days)", CONFIG["dormant_days"]),
              ("Elevated access codes", ", ".join(CONFIG["elevated_codes"])),
              ("Status labels", f"{meta['status_approve']} / {meta['status_remove']}"),
              ("Control owner", CONFIG["control_owner"]), ("E-mails read by", meta["method"]),
              ("E-mail files read", meta["email_files"])]
    write_table(wp, list(params[0]), [list(p) for p in params[1:]], widths={"Parameter": 36, "Value": 80})

    for name, key in SHEETS:
        df = tables[key]
        w = wb.create_sheet(name)
        write_table(w, list(df.columns), _rows(df), widths=WIDE,
                    date_fmt={c: f for c, f in DATE_COLS.items() if c in df.columns})
        cols = list(df.columns)
        flag = (("Exception", lambda v: v == "Yes") if "Exception" in cols else
                ("Removal Verified", lambda v: str(v).startswith("EXCEPTION")) if "Removal Verified" in cols else
                ("Rows Match Start Report", lambda v: v == "No") if "Rows Match Start Report" in cols else
                (None, None))
        if flag[0]:
            j = cols.index(flag[0]) + 1
            for rr in range(2, w.max_row + 1):
                if flag[1](w.cell(rr, j).value):
                    for c in w[rr]:
                        c.fill = RED_FILL
        if "Status" in cols:
            j = cols.index("Status") + 1
            for rr in range(2, w.max_row + 1):
                if w.cell(rr, j).value == meta["status_remove"] and w.cell(rr, j).fill != RED_FILL:
                    w.cell(rr, j).fill = AMB_FILL
        for c_ in ("Manager Reply Text", "Text", "Comments", "Exception Detail"):
            if c_ in cols:
                j = cols.index(c_) + 1
                for rr in range(2, w.max_row + 1):
                    w.cell(rr, j).alignment = Alignment(wrap_text=True, vertical="top")

    for name, key in (("Src_User_Report", "ur"), ("Src_Review_File", "rv"), ("Src_After_Listing", "al"),
                      ("Src_Removed_Accounts", "rm")):
        df = src[key]
        write_table(wb.create_sheet(name), list(df.columns), _rows(df),
                    date_fmt={"Last login": DATE_COLS["Last login"]} if "Last login" in df.columns else None)
    wb.save(path)


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="TRAX LA-CP4 user access review testing")
    ap.add_argument("--input", default=".", help="folder with the extracted LA-CP4 evidence")
    ap.add_argument("--due", default=None, help="due date stated in the request (informational)")
    ap.add_argument("--window", type=int, default=None, help="response window in days (default 30)")
    ap.add_argument("--output", default=None, help="output .xlsx path")
    ap.add_argument("--no-ai", action="store_true", help="do not call Azure OpenAI; keyword rules only")
    a = ap.parse_args()
    inp = os.path.abspath(a.input)
    load_env([os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"), os.path.join(inp, ".env"),
              os.path.join(os.getcwd(), ".env")])
    run(inp, a.due, a.output, use_ai=not a.no_ai, window=a.window)
