"""Chat with a run's testing results.

Azure OpenAI is given the table schema and writes SQLite SELECT queries; the backend executes them
read-only and feeds the rows back, so every figure in an answer comes from the data. Each answer
returns the SQL and rows used, and every exchange is appended to the run's chat_log.jsonl.

Read-only is enforced three ways: SELECT/WITH-only text check, a single statement, and an SQLite
authorizer that denies every operation except reading.
"""
import json
import re
import sqlite3
import threading
from datetime import datetime

import pandas as pd

TABLE_DOC = {
    "uar": "One row per access row tested (a user's role/code), the main testing sheet: status, request sent date, "
           "response date, days to respond and whether that is within the response window, self-approval, whether the "
           "row was in the manager's request e-mail, elevated access, dormancy, post-review match, e-mail cross-check "
           "(what the manager actually wrote) and exception detail.",
    "email_counts": "One row per manager: the access list pasted in their request e-mail compared with the start "
                    "User Report (rows in e-mail, rows in start report for the same users, rows in the review file, "
                    "difference, whether they match, and whether the To: line agrees with the list's Manager column).",
    "managers": "One row per reviewing manager: rows reviewed, approvals/rejections, request sent date, response date, "
                "days to respond, within-window flag, how the reply was read (AI or keyword), the decision and its quote, "
                "users the e-mail asked to remove, and consistency with the review file.",
    "removals": "One row per user whose access was rejected, who is in the client's removed-accounts analysis, or who "
                "disappeared from the post-review listing; shows whether removal is verified.",
    "completeness": "Access rows missing from or extra in the review file, rows with no manager, and start-report rows "
                    "that were not in any manager's request e-mail.",
    "movement": "Access rows removed or added between the start listing and the post-review listing.",
    "ipe": "Report extraction (IPE) checks: row counts, PDF timestamps, dates used.",
    "request_rows": "Every access row pasted in the request e-mails, with the manager it was sent to and the sent date.",
    "emails": "Every distinct e-mail message read (newest, quoted and oldest in each trail), its classification, the "
              "manager it was matched to and how it was read.",
    "summary": "Headline test metrics: section, metric, value.",
}

SYSTEM = """You are the data assistant inside Protiviti's user access review (UAR) testing tool.
Application tested: {app}. Start listing date: {start}. Post-review listing date: {after}.
Response due date: {due}. Status labels: approved = '{approve}', rejected = '{remove}'.

You answer an auditor's questions using ONLY the testing results stored in these SQLite tables:
{schema}

Headline metrics already computed (use these directly when they answer the question):
{summary}

How to work:
- To look something up, reply with JSON: {{"action": "query", "sql": "<one SQLite SELECT statement>"}}
  You will receive the result rows, then you may query again or answer.
- When you have what you need, reply with JSON: {{"action": "answer", "answer": "<your answer>"}}
- Reply with JSON only, no other text.

Rules:
- Never guess a number, name or date. Every fact in the answer must come from a query result or the
  headline metrics above. If the data cannot answer the question, say so plainly.
- Use only the tables and columns listed. Text comparisons are case-sensitive: prefer LIKE '%...%' or lower().
- Dates are stored as text 'YYYY-MM-DD HH:MM'. Show dates to the user as MM-DD-YYYY.
- When listing users, include user code and user name; list at most 25 rows and say how many there are in total.
- Keep answers short and in plain English, written for an auditor. Use simple bullet lists when listing items.
"""

ALLOWED_ACTIONS = {sqlite3.SQLITE_SELECT, sqlite3.SQLITE_READ, sqlite3.SQLITE_FUNCTION}
if hasattr(sqlite3, "SQLITE_RECURSIVE"):
    ALLOWED_ACTIONS.add(sqlite3.SQLITE_RECURSIVE)
FORBIDDEN = re.compile(r"\b(insert|update|delete|drop|alter|create|replace|attach|detach|pragma|vacuum|reindex|"
                       r"analyze|begin|commit|rollback|savepoint|release|load_extension)\b", re.I)


def snake(c):
    return re.sub(r"[^a-z0-9]+", "_", str(c).lower()).strip("_") or "col"


def _sql_value(v):
    if v is None:
        return None
    if isinstance(v, float) and pd.isna(v):
        return None
    if v is pd.NaT or v is pd.NA:
        return None
    if isinstance(v, (pd.Timestamp, datetime)):
        return pd.Timestamp(v).strftime("%Y-%m-%d %H:%M")
    if isinstance(v, (int, float, str)):
        return v
    return str(v)


class RunDB:
    """In-memory SQLite copy of one run's result tables, read-only once loaded."""

    def __init__(self, results):
        self.lock = threading.Lock()
        self.conn = sqlite3.connect(":memory:", check_same_thread=False)
        self.schema = {}
        tables = dict(results["tables"])
        tables["summary"] = pd.DataFrame(results["summary"])
        for name, df in tables.items():
            df = df.copy()
            cols, seen = [], set()
            for c in df.columns:
                sc = snake(c)
                while sc in seen:
                    sc += "_2"
                seen.add(sc)
                cols.append(sc)
            df.columns = cols
            for c in cols:   # object dtype keeps whole numbers whole (a list with blanks would become floats)
                df[c] = pd.Series([_sql_value(v) for v in df[c].tolist()], index=df.index, dtype=object)
            ints = {c: "INTEGER" for c in cols
                    if df[c].notna().any() and all(isinstance(v, int) and not isinstance(v, bool)
                                                   for v in df[c].dropna().tolist())}
            df.to_sql(name, self.conn, index=False, dtype=ints or None)
            examples = {}
            for c in cols:
                vals = [v for v in pd.Series(df[c]).dropna().unique().tolist() if v != ""]
                if 0 < len(vals) <= 12 and all(isinstance(v, str) and len(v) <= 60 for v in vals):
                    examples[c] = vals
            self.schema[name] = {"columns": cols, "rows": len(df), "examples": examples}
        self.conn.set_authorizer(self._authorize)
        self.conn.set_progress_handler(self._progress, 10000)
        self._budget = 0

    @staticmethod
    def _authorize(action, *_):
        return sqlite3.SQLITE_OK if action in ALLOWED_ACTIONS else sqlite3.SQLITE_DENY

    def _progress(self):
        self._budget -= 1
        return 1 if self._budget < 0 else 0          # non-zero aborts a runaway query

    def schema_text(self):
        out = []
        for name, info in self.schema.items():
            out.append(f"- {name} ({info['rows']} rows): {TABLE_DOC.get(name, '')}")
            out.append(f"  columns: {', '.join(info['columns'])}")
            for c, vals in info["examples"].items():
                out.append(f"  {c} values: " + ", ".join(f"'{v}'" for v in vals))
        return "\n".join(out)

    def query(self, sql, limit=500):
        sql = (sql or "").strip().rstrip(";").strip()
        if not re.match(r"(?is)^(select|with)\b", sql):
            raise ValueError("only SELECT queries are allowed")
        if ";" in sql:
            raise ValueError("only one statement is allowed")
        if FORBIDDEN.search(re.sub(r"'[^']*'", "''", sql)):
            raise ValueError("only read-only queries are allowed")
        with self.lock:
            self._budget = 2000                      # ~20M VM steps
            cur = self.conn.execute(sql)
            cols = [d[0] for d in cur.description or []]
            rows = cur.fetchmany(limit + 1)
        truncated = len(rows) > limit
        return {"columns": cols, "rows": [list(r) for r in rows[:limit]], "row_count": min(len(rows), limit),
                "truncated": truncated}


def _fmt_date(v):
    try:
        return pd.Timestamp(v).strftime("%m-%d-%Y") if v is not None and not pd.isna(v) else "not set"
    except Exception:
        return str(v)


def answer_question(ai, db, results, question, history=None, max_steps=5):
    """Returns {"answer": str, "steps": [{"sql", "columns", "rows", "row_count", "error"}]}."""
    meta = results["meta"]
    summary_txt = "\n".join(f"- {s['metric']}: {s['value']}" for s in results["summary"])
    system = SYSTEM.format(app=meta.get("app_name", ""), start=_fmt_date(meta.get("start_date")),
                           after=_fmt_date(meta.get("after_date")), due=_fmt_date(meta.get("due_date")),
                           approve=meta.get("status_approve", ""), remove=meta.get("status_remove", ""),
                           schema=db.schema_text(), summary=summary_txt)
    msgs = [{"role": "system", "content": system}]
    for h in (history or [])[-6:]:
        if h.get("role") in ("user", "assistant") and h.get("content"):
            msgs.append({"role": h["role"], "content": str(h["content"])[:2000]})
    msgs.append({"role": "user", "content": question})
    steps = []
    for step in range(max_steps):
        if step == max_steps - 1:
            msgs.append({"role": "user", "content": "Answer now using only the results you already have."})
        out = ai.chat(None, None, use_cache=False, messages=msgs)
        if not isinstance(out, dict):
            return {"answer": "The AI service did not respond. Try again in a moment.", "steps": steps}
        action = out.get("action")
        if action == "answer" or (not out.get("sql") and out.get("answer")):
            return {"answer": str(out.get("answer") or "").strip() or "No answer returned.", "steps": steps}
        sql = str(out.get("sql") or "")
        msgs.append({"role": "assistant", "content": json.dumps({"action": "query", "sql": sql})})
        try:
            res = db.query(sql)
            steps.append({"sql": sql, **res, "rows": res["rows"][:50], "error": None})
            preview = json.dumps({"columns": res["columns"], "rows": res["rows"][:100],
                                  "row_count": res["row_count"], "truncated": res["truncated"]},
                                 default=str, ensure_ascii=False)
            msgs.append({"role": "user", "content": f"QUERY RESULT: {preview[:12000]}"})
        except Exception as e:
            steps.append({"sql": sql, "columns": [], "rows": [], "row_count": 0, "error": str(e)})
            msgs.append({"role": "user", "content": f"QUERY ERROR: {e}. Fix the query or answer without it."})
    return {"answer": "I could not reach an answer within the query limit. Try a narrower question.", "steps": steps}
