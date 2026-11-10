#requirements:
#wmill
#requests
#openpyxl

import base64
import io
import json
import re
import time
from datetime import date, datetime
from urllib.parse import quote

import requests
import wmill
from openpyxl import Workbook
from openpyxl.styles import Border, Font, PatternFill, Side

# =========================================================
# NNAS REVENUE REPORT GENERATOR v7 (payment-date primary, EST-aligned) -> SharePoint Applications/zDATA
# =========================================================
# Shapes confirmed by probe (app 35542766):
#   - task list entries carry: id, name ("T - Expedited Service Payment..."),
#     status (INT: 1 not started, 2 in progress, 3 completed), completed_at
#     ("2026-07-27T13:32:57"). All five payment tasks exist on every app;
#     only the applicable one reaches status 3.
#   - completed payment task detail data: amount ("750"), charge_id,
#     service_response (full Stripe charge).
#   - custom_fields inline on application detail; 153648 populated
#     ("2026-07-27").
#   - RB selections: task 2048019 "Select Nursing Discipline", fields
#     faO2qkS4PX (RN) and IWgvFrNZHX (LPN); responses are option INDEX
#     arrays. Mapped via RN_OPTIONS/LPN_OPTIONS (from calibration script).
#   - Issued date: field 153677, falling back to the completion timestamp
#     of the "Applicant Credential Report Copies" task (confirmed present
#     with the report PDF on issued applications).
#
# Report semantics for month M / year Y unchanged (see v2 header).
# =========================================================

SMAPPLY_TOKEN_VAR = "u/epham2/SMAPPLY_ACCESS_TOKEN"
SMAPPLY_BASE = "https://nnas.smapply.io/api"

FIELD_FIRST_SUBMISSION = "153648"
FIELD_REPORT_ISSUED = "153677"

PAYMENT_TAX = [  # longest-first; matching is substring so the "T - " prefix is fine
    ("Expedited Service Payment HST ON", 97.5),
    ("Expedited Service Payment HST NS", 105.0),
    ("Expedited Service Payment HST MT", 112.5),
    ("Expedited Service Payment GST", 37.5),
    ("Expedited Service Payment", 0.0),
]
COST_CAD = 750.0
UNIT_PRICE_USD = 350
STATUS_COMPLETED = 3

# RB selections live in task 2048019 (Select Nursing Discipline) as option
# INDEX arrays on these two fields. RN_OPTIONS / LPN_OPTIONS map index ->
# label; run smapply_rb_calibration.py once and paste its output here.
# While either list is None, Q-F rows carry "RN option <i>" placeholders
# and the summary warns.
RB_TASK_ID = 2048019
RN_FIELD = "faO2qkS4PX"
LPN_FIELD = "IWgvFrNZHX"
RN_OPTIONS = [
    "NL College of Nurses",
    "PEI College of Nursing & Midwifery",
    "College of Nursing of New Brunswick",
    "Nova Scotia College of Nursing",
    "College of Nurses of Ontario",
    "College of Registered Nurses of Manitoba",
    "College of Registered Nurses of Saskatchewan",
    "College of Registered Nurses of Alberta",
    "BC College of Nurses and Midwives",
]
LPN_OPTIONS = [
    "NL College of Nurses",
    "PEI College of Nursing & Midwifery",
    "Association of New Brunswick Licensed Practical Nurses",
    "Nova Scotia College of Nursing",
    "College of Nurses of Ontario",
    "College of Licensed Practical Nurses of Manitoba",
    "College of Licensed Practical Nurses of Saskatchewan",
    "College of Licensed Practical Nurses and Health Care Aides of Alberta",
    "BC College of Nurses and Midwives",
]

# Issued-date fallback: completion of the applicant report copies task.
ISSUED_TASK_NAME = "Applicant Credential Report Copies"

_NOT_PRESENT = object()

# ---- SharePoint / Microsoft Graph (same variables as the intake script) ----
TENANT_ID = wmill.get_variable("u/epham2/AZURE_TENANT_ID")
CLIENT_ID = wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_ID")
CLIENT_SECRET = wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_SECRET")

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SHAREPOINT_HOSTNAME = "nnasca.sharepoint.com"
SITE_PATH = "/sites/NNAS-Portal-Sharepoint"
DOCUMENT_LIBRARY_NAME = "Applications"
UPLOAD_FOLDER = "zDATA"


# ---------------------------------------------------------
# SMAPPLY HTTP
# ---------------------------------------------------------

def _sm_headers():
    token = wmill.get_variable(SMAPPLY_TOKEN_VAR)
    return {"Authorization": "Bearer %s" % token,
            "Content-Type": "application/json"}


def _get(url, params=None, retries=3):
    for _ in range(retries):
        r = requests.get(url, headers=_sm_headers(), params=params, timeout=60)
        if r.status_code == 429:
            wait = int(r.headers.get("Retry-After", "5"))
            print("  429 rate limited, sleeping %ss" % wait)
            time.sleep(wait)
            continue
        r.raise_for_status()
        return r.json()
    raise RuntimeError("giving up after retries: %s" % url)


def iter_applications(program_id):
    page = 1
    while True:
        body = _get("%s/applications/" % SMAPPLY_BASE,
                    params={"program": program_id, "page": page})
        results = body.get("results", [])
        if not results:
            break
        for item in results:
            yield item
        num_pages = body.get("num_pages")
        if num_pages is not None:
            if page >= num_pages:
                break
        elif body.get("next") in (None, ""):
            break
        page += 1


def get_tasks(appid):
    tasks, page = [], 1
    while True:
        body = _get("%s/applications/%s/tasks/" % (SMAPPLY_BASE, appid),
                    params={"page": page})
        if isinstance(body, list):
            return body
        tasks.extend(body.get("results", []))
        num_pages = body.get("num_pages")
        if num_pages is not None:
            if page >= num_pages:
                break
        elif body.get("next") in (None, ""):
            break
        page += 1
    return tasks


def get_task_detail(appid, taskid):
    return _get("%s/applications/%s/tasks/%s/" % (SMAPPLY_BASE, appid, taskid))


# ---------------------------------------------------------
# EXTRACTION
# ---------------------------------------------------------

def find_cf(custom_fields, field_id):
    target = str(field_id)
    if isinstance(custom_fields, list):
        for cf in custom_fields:
            if str(cf.get("id")) == target:
                return cf.get("value")
    elif isinstance(custom_fields, dict):
        if target in custom_fields:
            v = custom_fields[target]
            return v.get("value") if isinstance(v, dict) else v
        for v in custom_fields.values():
            if isinstance(v, dict) and str(v.get("id")) == target:
                return v.get("value")
    return _NOT_PRESENT


def parse_date(value):
    if value in (None, "", _NOT_PRESENT):
        return None
    s = str(value).strip()
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", s)
    if m:
        return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    for fmt in ("%m/%d/%Y", "%d-%m-%Y", "%m/%d/%y"):
        try:
            return datetime.strptime(s[:10], fmt).date()
        except ValueError:
            pass
    return None


TZ_OFFSET_HOURS = 5  # set by main(); API completed_at is UTC, TM/TruMerit is UTC-5


def parse_ts_local(value):
    """Task timestamp (UTC) -> local date after subtracting TZ_OFFSET_HOURS."""
    if value in (None, ""):
        return None
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2}):(\d{2})", str(value))
    if not m:
        return parse_date(value)
    from datetime import timedelta
    dt = datetime(*(int(g) for g in m.groups())) - timedelta(hours=TZ_OFFSET_HOURS)
    return dt.date()


def title_of(t):
    return str(t.get("name") or t.get("title") or "")


def task_completed_date(t):
    """Completed iff status == 3 (int, confirmed) or completed_at is set."""
    if t.get("status") == STATUS_COMPLETED or t.get("completed_at"):
        return parse_ts_local(t.get("completed_at")) or parse_ts_local(t.get("updated_at"))
    return None


def extract_payment(tasks):
    """(tax, paid_date, task) for the completed payment task. All five
    payment tasks exist per app; only the applicable one is status 3."""
    for t in tasks:
        tt = title_of(t).lower()
        for name, tax in PAYMENT_TAX:
            if name.lower() in tt:
                d = task_completed_date(t)
                if d:
                    return tax, d, t
                break  # a payment task, but not the completed one
    return None, None, None


def _map_options(indexes, options, prefix):
    out = []
    for i in indexes:
        if options is not None and isinstance(i, int) and 0 <= i < len(options) \
                and options[i]:
            out.append(options[i])
        else:
            out.append("%s option %s" % (prefix, i))
    return out


def extract_rbs(appid, misses):
    """Regulator labels (RN + LPN, uncollapsed) from task 2048019 option
    indexes, mapped through RN_OPTIONS / LPN_OPTIONS."""
    try:
        detail = get_task_detail(appid, RB_TASK_ID)
    except Exception as e:
        misses.append("app %s RB task error: %s" % (appid, e))
        return []
    data = detail.get("data") or {}

    def idxs(field):
        v = data.get(field) or {}
        resp = v.get("response") if isinstance(v, dict) else None
        return list(resp) if isinstance(resp, list) else []

    return (_map_options(sorted(idxs(RN_FIELD)), RN_OPTIONS, "RN") +
            _map_options(sorted(idxs(LPN_FIELD)), LPN_OPTIONS, "LPN"))


def extract_applicant(item):
    a = item.get("applicant")
    if isinstance(a, dict):
        return {"sysid": a.get("id"), "first": a.get("first_name") or "",
                "last": a.get("last_name") or "", "email": a.get("email") or ""}
    return {"sysid": a, "first": item.get("applicant_first_name", ""),
            "last": item.get("applicant_last_name", ""),
            "email": item.get("applicant_email", "")}


# ---------------------------------------------------------
# SHAREPOINT
# ---------------------------------------------------------

_GRAPH_TOKEN = None
_SP_DRIVE_ID = None


def get_graph_token():
    global _GRAPH_TOKEN
    if _GRAPH_TOKEN:
        return _GRAPH_TOKEN
    r = requests.post(
        "https://login.microsoftonline.com/%s/oauth2/v2.0/token" % TENANT_ID.strip(),
        data={"client_id": CLIENT_ID.strip(),
              "client_secret": CLIENT_SECRET.strip(),
              "scope": "https://graph.microsoft.com/.default",
              "grant_type": "client_credentials"},
        timeout=60)
    if r.status_code != 200:
        raise RuntimeError("Graph token request failed: %s %s"
                           % (r.status_code, r.text[:1000]))
    _GRAPH_TOKEN = r.json()["access_token"]
    return _GRAPH_TOKEN


def _graph(method, endpoint, expected=(200,), **kwargs):
    headers = kwargs.pop("headers", {})
    headers["Authorization"] = "Bearer %s" % get_graph_token()
    r = requests.request(method, GRAPH_BASE + endpoint, headers=headers,
                         timeout=120, **kwargs)
    if r.status_code not in expected:
        raise RuntimeError("Graph %s %s failed: %s %s"
                           % (method, endpoint, r.status_code, r.text[:2000]))
    return r.json() if r.content else None


def upload_to_zdata(filename, content_bytes):
    global _SP_DRIVE_ID
    if not _SP_DRIVE_ID:
        site = _graph("GET", "/sites/%s:%s" % (SHAREPOINT_HOSTNAME, SITE_PATH))
        drives = _graph("GET", "/sites/%s/drives" % site["id"])["value"]
        drive = next((d for d in drives
                      if d.get("name") == DOCUMENT_LIBRARY_NAME), None)
        if not drive:
            raise RuntimeError("Drive '%s' not found; available: %s"
                               % (DOCUMENT_LIBRARY_NAME,
                                  [d.get("name") for d in drives]))
        _SP_DRIVE_ID = drive["id"]
        print("Connected to SharePoint drive: %s" % DOCUMENT_LIBRARY_NAME)
    path = quote("%s/%s" % (UPLOAD_FOLDER, filename), safe="/")
    item = _graph("PUT", "/drives/%s/root:/%s:/content" % (_SP_DRIVE_ID, path),
                  headers={"Content-Type": "application/octet-stream"},
                  data=content_bytes, expected=(200, 201))
    print("Uploaded to SharePoint: %s/%s" % (UPLOAD_FOLDER, filename))
    print("URL: %s" % item.get("webUrl"))
    return item.get("webUrl")


# ---------------------------------------------------------
# XLSX
# ---------------------------------------------------------

HDR_FILL = PatternFill("solid", start_color="E7F3FD")
THIN = Side(style="thin", color="B0B0B0")


def _sheet(wb, name, title, headers):
    ws = wb.create_sheet(name)
    ws["A1"] = title
    ws["A1"].font = Font(bold=True, size=12)
    for j, h in enumerate(headers, start=1):
        c = ws.cell(row=4, column=j, value=h)
        c.font = Font(bold=True, size=9)
        c.fill = HDR_FILL
        c.border = Border(bottom=THIN)
    ws.freeze_panes = "A5"
    return ws


def _rows(ws, rows, widths):
    for i, row in enumerate(rows, start=5):
        for j, v in enumerate(row, start=1):
            c = ws.cell(row=i, column=j, value=v)
            if isinstance(v, date):
                c.number_format = "mm/dd/yyyy"
    for j, w in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=4, column=j).column_letter].width = w


# ---------------------------------------------------------
# MAIN
# ---------------------------------------------------------

def main(report_month: int = 7,
         report_year: int = 2026,
         program_id: int = 295185,
         probe_app_id: int = 0,
         task_fetch_sleep: float = 0.0,
         upload_to_sharepoint: bool = True,
         tz_offset_hours: int = 5):
    # API task timestamps are UTC; 5 aligns dates with TruMerit's UTC-5
    # convention (their invoice basis). Set 0 to count by raw UTC.
    global TZ_OFFSET_HOURS
    TZ_OFFSET_HOURS = tz_offset_hours

    if probe_app_id:
        tasks = get_tasks(probe_app_id)
        print("=== tasks (%d) ===" % len(tasks))
        for t in tasks:
            print(json.dumps({k: t.get(k) for k in
                              ("id", "name", "status", "completed_at")},
                             default=str))
        print("\n=== data containers per task (keys + truncated values) ===")
        for t in tasks:
            try:
                detail = get_task_detail(probe_app_id, t.get("id"))
            except Exception as e:
                print("task %s detail error: %s" % (t.get("id"), e))
                continue
            data = detail.get("data") or {}
            if not isinstance(data, dict) or not data:
                continue
            print("\n-- task %s: %s --" % (t.get("id"), title_of(t)))
            for k, v in data.items():
                s = json.dumps(v, default=str)
                print("  field %-14r = %s" % (k, s[:300]))
        return "probe complete"

    month_start = date(report_year, report_month, 1)
    month_end = (date(report_year + 1, 1, 1) if report_month == 12
                 else date(report_year, report_month + 1, 1))

    apps, misses = [], []
    stats = {"seen": 0, "no_payment": 0, "task_paid_used": 0,
             "cf_paid_fallback": 0, "no_rbs": 0,
             "issued_task_fallback": 0, "issued_known": 0}

    for item in iter_applications(program_id):
        stats["seen"] += 1
        appid = item.get("id")
        title = item.get("title") or item.get("name") or str(appid)
        who = extract_applicant(item)
        cfs = item.get("custom_fields")

        tasks = get_tasks(appid)
        if task_fetch_sleep:
            time.sleep(task_fetch_sleep)
        tax, task_paid, _ = extract_payment(tasks)

        # Payment task completion is the PAYMENT date (matches TruMerit's
        # invoice basis); 153648 is the SUBMISSION date and can lag payment
        # by weeks, so it is only a fallback when no payment task completed.
        cf_paid = parse_date(find_cf(cfs, FIELD_FIRST_SUBMISSION))
        paid = task_paid or cf_paid
        if task_paid:
            stats["task_paid_used"] += 1
        elif cf_paid:
            stats["cf_paid_fallback"] += 1
        if not paid:
            stats["no_payment"] += 1
            continue

        issued = parse_date(find_cf(cfs, FIELD_REPORT_ISSUED))
        if not issued:
            copies = next((t for t in tasks
                           if ISSUED_TASK_NAME.lower() in title_of(t).lower()), None)
            if copies is not None:
                issued = task_completed_date(copies)
                if issued:
                    stats["issued_task_fallback"] += 1
        if issued:
            stats["issued_known"] += 1

        regs = []
        if month_start <= paid < month_end:
            regs = extract_rbs(appid, misses)
            if not regs:
                stats["no_rbs"] += 1
        apps.append({"appid": appid, "title": title, **who,
                     "paid": paid, "tax": tax if tax is not None else 0.0,
                     "issued": issued, "regs": regs})
        if stats["seen"] % 100 == 0:
            print("  ...%d applications processed" % stats["seen"])

    q = [a for a in apps if month_start <= a["paid"] < month_end]
    issued_pg = [a for a in apps if a["issued"] and
                 month_start <= a["issued"] < month_end]
    deferred = [a for a in apps if a["paid"] < month_end and
                (a["issued"] is None or a["issued"] >= month_end)]

    wb = Workbook()
    wb.remove(wb.active)
    period = "%02d/%d" % (report_month, report_year)

    ws = _sheet(wb, "DEFERRED REVENUE",
                "NNAS Deferred Revenue - month-end %s (paid on/before, not issued by month-end)" % period,
                ["Order Name", "Cgfns Id", "Order Headers Id", "Trans Date",
                 "Last Name", "First Name", "Billed CAD", "Issued Date"])
    _rows(ws, [["Expedited Application (T)", a["sysid"], a["title"], a["paid"],
                a["last"], a["first"], COST_CAD + a["tax"], a["issued"]]
               for a in sorted(deferred, key=lambda x: x["paid"])],
          [24, 12, 20, 12, 18, 18, 12, 12])

    ws = _sheet(wb, "REVENUE - REPORTS ISSUED",
                "NNAS Revenue from Reports Issued - issued in %s" % period,
                ["#", "Cgfns Id", "Order Name", "Order #", "Trans Date",
                 "Last Name", "First Name", "Billed CAD", "Revenue Date"])
    _rows(ws, [[i + 1, a["sysid"], "Expedited Application (T)", a["title"],
                a["paid"], a["last"], a["first"], COST_CAD + a["tax"], a["issued"]]
               for i, a in enumerate(sorted(issued_pg, key=lambda x: x["issued"]))],
          [5, 12, 24, 20, 12, 18, 18, 12, 12])

    ws = _sheet(wb, "ATTACHMENT Q",
                "ATTACHMENT Q - Expedited Service payments %s" % period,
                ["NNAS / MONERIS CUSTOMER ID", "ORDER NUMBER", "TRANSACTION DATE",
                 "LAST NAME", "FIRST NAME", "PAYMENT TRANSFERRED",
                 "COST ($CAD)", "TAX ($CAD)", "TOTAL ($CAD)"])
    _rows(ws, [[a["sysid"], a["title"], a["paid"], a["last"], a["first"], 0,
                COST_CAD, a["tax"], COST_CAD + a["tax"]]
               for a in sorted(q, key=lambda x: x["paid"], reverse=True)],
          [16, 20, 14, 18, 18, 12, 11, 11, 11])

    qf_rows = []
    for a in sorted(q, key=lambda x: x["paid"], reverse=True):
        for reg in a["regs"]:
            qf_rows.append([a["sysid"], a["title"], a["paid"], a["last"],
                            a["first"], reg, 0, 0, 0, 0])
    ws = _sheet(wb, "ATTACHMENT Q-F",
                "ATTACHMENT Q-F - one row per selected regulatory body, %s" % period,
                ["NNAS / MONERIS CUSTOMER ID", "ORDER NUMBER", "TRANSACTION DATE",
                 "LAST NAME", "FIRST NAME", "DESTINATION",
                 "PAYMENT TRANSFERRED", "COST ($USD)", "TAX ($USD)", "TOTAL ($USD)"])
    _rows(ws, qf_rows, [16, 20, 14, 18, 18, 58, 12, 11, 11, 11])

    ws = _sheet(wb, "INVOICE SUMMARY",
                "INVOICE SUMMARY - %s" % period,
                ["Sheet", "Qty", "Description", "Unit Price (US$)", "Total (US$)"])
    tax_total = sum(a["tax"] for a in q)
    _rows(ws, [
        ["Q", len(q), "Expedited Service payments %s" % period,
         UNIT_PRICE_USD, len(q) * UNIT_PRICE_USD],
        ["Q-F", len(qf_rows), "Free reports (one per regulator)", 0, 0],
        ["", "", "CAD collected (incl. tax): %.2f" % (len(q) * COST_CAD + tax_total), "", ""],
        ["", "", "of which tax (CAD): %.2f" % tax_total, "", ""],
    ], [8, 8, 60, 16, 14])

    buf = io.BytesIO()
    wb.save(buf)
    filename = "NNAS_Revenue_Report_%d-%02d_%s.xlsx" % (
        report_year, report_month, datetime.now().strftime("%Y%m%d_%H%M%S"))

    print("\n--- Summary ---")
    print("  Applications seen        : %d" % stats["seen"])
    print("  Unpaid (skipped)         : %d" % stats["no_payment"])
    print("  Paid via payment task    : %d" % stats["task_paid_used"])
    print("  Paid via 153648 fallback : %d" % stats["cf_paid_fallback"])
    print("  Q rows (%s)         : %d" % (period, len(q)))
    print("  Q-F regulator rows       : %d  (apps with no RBs found: %d)"
          % (len(qf_rows), stats["no_rbs"]))
    print("  Reports issued (%s) : %d" % (period, len(issued_pg)))
    print("  Deferred at month-end    : %d" % len(deferred))
    print("  Issued dates known       : %d (of which via report-copies task: %d)"
          % (stats["issued_known"], stats["issued_task_fallback"]))
    if RN_OPTIONS is None or LPN_OPTIONS is None:
        print("  WARNING: RN_OPTIONS/LPN_OPTIONS not set; Q-F destinations are "
              "index placeholders. Run smapply_rb_calibration.py and paste "
              "its output into this script.")
    for msg in misses[:20]:
        print("  MISS: %s" % msg)

    result = {"filename": filename,
              "counts": {"q": len(q), "qf": len(qf_rows),
                         "issued": len(issued_pg), "deferred": len(deferred)}}

    if upload_to_sharepoint:
        try:
            result["sharepoint_url"] = upload_to_zdata(filename, buf.getvalue())
            return result
        except Exception as e:
            print("SharePoint upload FAILED, returning base64 instead: %s" % e)

    result["xlsx_base64"] = base64.b64encode(buf.getvalue()).decode()
    return result