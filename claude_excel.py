"""
claude_excel.py — build Excel the way claude.ai builds it.

WHY THIS FILE EXISTS
--------------------
The user asked a fair question: claude.ai produced a beautiful, working
dashboard from uploaded sheets. Why can this platform not do the same?

The answer is not the model. It is the METHOD.

  What this platform does today:
      ask the model for a JSON description of a workbook,
      then a fixed Python renderer draws whatever that JSON says.
      The model never writes spreadsheet code, never runs it, never opens the
      result, and never discovers that something did not work.

  What claude.ai does:
      the model WRITES Python code, RUNS it in a sandbox, OPENS the file it
      just made, checks it, fixes its own mistakes, and repeats until the
      workbook is right.

That second loop is why one output feels alive and the other feels like a
template. It is a difference in architecture, not intelligence.

Anthropic exposes the same sandbox on the API (the code execution tool, public
beta). This module uses it. Claude writes openpyxl code, runs it, verifies the
file opens, and we download the finished workbook.

SECURITY NOTE, AND IT IS THE REASON I CHOSE THIS ROUTE
-----------------------------------------------------
Running model-written code is the single most dangerous thing a platform can
add. Done on our own server it would be arbitrary code execution inside the
same container that holds customer data - a serious and hard-to-contain risk.

Here the code runs in ANTHROPIC'S sandbox, not ours. Nothing executes on
Railway. Our server only sends a prompt and downloads a finished file. The
dangerous capability is used without the dangerous exposure.
"""

import json
import os
import time
import urllib.request
import urllib.error

API = "https://api.anthropic.com/v1"
BETAS = "code-execution-2025-08-25,files-api-2025-04-14"
MODEL = os.environ.get("CLAUDE_EXCEL_MODEL", "claude-sonnet-4-5")

# A workbook that has to be built, run and checked needs room to think and to
# write real code. This is deliberately generous; it is one call per document.
MAX_TOKENS = 16000
HTTP_TIMEOUT = 180


def _post(path, payload, key, extra_headers=None):
    body = json.dumps(payload).encode()
    headers = {
        "x-api-key": key,
        "anthropic-version": "2023-06-01",
        "anthropic-beta": BETAS,
        "content-type": "application/json",
    }
    if extra_headers:
        headers.update(extra_headers)
    req = urllib.request.Request(API + path, data=body, headers=headers, method="POST")
    with urllib.request.urlopen(req, timeout=HTTP_TIMEOUT) as r:
        return json.loads(r.read().decode())


def _download_file(file_id, key):
    req = urllib.request.Request(
        API + "/files/%s/content" % file_id,
        headers={"x-api-key": key, "anthropic-version": "2023-06-01",
                 "anthropic-beta": BETAS},
    )
    with urllib.request.urlopen(req, timeout=HTTP_TIMEOUT) as r:
        return r.read()


def _brief(objective, context, data, sym):
    """The instruction Claude receives. This is the whole product, really -
    everything below is plumbing."""
    return f"""Build a production-grade Excel workbook and save it to /tmp/workbook.xlsx

OBJECTIVE
{objective}

COMPANY CONTEXT
{context or "(none given)"}

SOURCE MATERIAL — every figure in the workbook must come from here or be
derived from it by arithmetic. Do not invent numbers. Where something is an
assumption, put it in the Assumptions sheet and label it.
{data[:60000] if data else "(none supplied — build the structure with clearly marked input cells)"}

HOW TO BUILD IT
Write Python using openpyxl. Run it. Then OPEN the file you created, check it,
and fix anything wrong before you finish.

STRUCTURE — use the sheets that genuinely fit this request, in this order:
  README        what this workbook is, how to use it, what to fill in
  Control Panel named input cells the user changes; nothing else is editable
  RAW DATA      where the user pastes new data each period
  Calculations  the engine — formulas only, no typed-in results
  Analysis      variance, trend, ratios, whatever the objective needs
  Dashboard     the summary a director reads first, with charts
  Assumptions   every assumption, each in its own labelled cell
  Audit         control totals and reconciliation checks

NON-NEGOTIABLE RULES
1. Every derived cell must be a LIVE FORMULA (=B4*C4), never a typed result.
   This is what makes the workbook keep working when new data is pasted in.
2. Every assumption is a single named cell referenced everywhere else. Never
   hardcode the same number twice.
3. Use openpyxl Table objects with structured references so ranges grow.
4. Use openpyxl.chart for real Excel charts — not images of charts.
5. Conditional formatting where it aids reading: variances, RAG status, aging.
6. Data validation on input cells so bad entries are refused at the point of
   typing.
7. Freeze panes on every data sheet. Sensible column widths. Number formats
   with thousands separators, currency {sym}, and percentages as percentages.
8. Control totals on the Audit sheet that reconcile to zero, so a user can see
   at a glance whether the workbook still adds up.
9. Colour convention: input cells blue, formulas black, cross-sheet links green.
10. Sheet names must avoid : \\ / ? * [ ] and stay under 31 characters.

DO NOT
- Do not write a .xlsm or claim macros. A real macro project cannot be created
  this way and a renamed file is a broken file.
- Do not leave placeholder text such as TBD or "insert data here".
- Do not put a number in a formula where a named input cell belongs.

FINISH BY
Reopening /tmp/workbook.xlsx with openpyxl, printing the sheet names and the
number of formula cells, and confirming it loads without error."""


def build_workbook(objective, context, data, claude_key, sym="\u20b9"):
    """Returns (bytes, mode, reason). Never raises - a failure returns
    (None, 'unavailable', why) so the caller can fall back to the existing
    renderer rather than breaking."""
    key = (claude_key or "").strip()
    if not key:
        return None, "unavailable", "no Claude key supplied"

    payload = {
        "model": MODEL,
        "max_tokens": MAX_TOKENS,
        "tools": [{"type": "code_execution_20250825", "name": "code_execution"}],
        "messages": [{"role": "user", "content": _brief(objective, context, data, sym)}],
    }

    try:
        resp = _post("/messages", payload, key)
    except urllib.error.HTTPError as e:
        detail = ""
        try:
            detail = e.read().decode()[:200]
        except Exception:
            pass
        return None, "unavailable", "Claude code execution refused (HTTP %s) %s" % (e.code, detail)
    except Exception as e:
        return None, "unavailable", "Claude code execution unreachable: %s" % str(e)[:160]

    # Find the workbook among the files the sandbox produced. We look for an
    # .xlsx by name rather than taking the first file, because the run may also
    # produce charts, logs or intermediate CSVs.
    candidates = []
    for block in resp.get("content", []) or []:
        if block.get("type") != "code_execution_tool_result":
            continue
        content = block.get("content") or {}
        for f in (content.get("content") or []):
            fid = f.get("file_id")
            name = str(f.get("name") or "")
            if fid and name.lower().endswith(".xlsx"):
                candidates.append((fid, name))
            elif fid and not name:
                candidates.append((fid, ""))

    if not candidates:
        return None, "unavailable", "Claude ran but produced no .xlsx file"

    for fid, _name in candidates:
        try:
            blob = _download_file(fid, key)
        except Exception:
            continue
        # The same rule the browser now uses: an xlsx is a ZIP, and every ZIP
        # starts PK\x03\x04. Anything else is not a spreadsheet, whatever it is
        # called.
        if blob[:4] == b"PK\x03\x04" and len(blob) > 2000:
            return blob, "claude-code", "built and verified by Claude in Anthropic's sandbox"

    return None, "unavailable", "Claude returned a file that was not a valid workbook"
