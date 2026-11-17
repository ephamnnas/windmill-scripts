# requirements:
# requests
# stripe
# wmill

"""
Windmill path: f/refunds/common

Shared configuration and helpers for the NNAS refund workflow. Every other
script in f/refunds imports from here, so all settings live in CONFIG below.

Deploy this script first. It has a main() only so Windmill accepts it; running
it just prints the active configuration.
"""

import csv
import io
import json
import os
import re
from datetime import date, datetime, timezone
from decimal import Decimal, InvalidOperation
from urllib.parse import quote

import requests
import stripe
import wmill

# ======================================================================
# CONFIG
# ======================================================================

# --- Roles (lowercase emails) ------------------------------------------
REQUESTERS = {"thobbs@nnas.ca", "aonikan@nnas.ca", "cbond@nnas.ca"}
APPROVERS = {"tparsons@nnas.ca"}
PROCESSORS = {"tparsons@nnas.ca"}  # issues the refund in Stripe, presses Lock
ADMINS = {"knguyen@nnas.ca"}  # sees everything, can request and process
ADMINS_CAN_APPROVE = False
REQUESTERS_SEE_ALL = True  # False = requesters only see their own requests

# Whoever saves and deploys the two Windmill apps. See SETUP step 7.
APP_PUBLISHER_EMAIL = "knguyen@nnas.ca"

# --- Testing ------------------------------------------------------------
# Lock only these SMApply application ids. Set to None for production.
LOCK_ALLOWLIST = {35367543}

# Lets an ADMIN run every step alone (request, approve, process) so you can
# test without other people. Only applies when the application is in
# LOCK_ALLOWLIST AND the Stripe key is a test-mode key. Set False for
# production.
TEST_SOLO = True

# --- Windmill data table -----------------------------------------------
# Name of the data table in Workspace settings > Data Tables. Tables live in
# its "refund" schema (created by 01_refund_register.up.sql).
DATATABLE = "main"

# --- Windmill paths -----------------------------------------------------
STRIPE_KEY_VAR = "f/refunds/STRIPE_RESTRICTED_KEY"
SMAPPLY_TOKEN_VAR = "u/epham2/SMAPPLY_ACCESS_TOKEN"
AZURE_TENANT_VAR = "u/epham2/AZURE_TENANT_ID"
SP_CLIENT_ID_VAR = "u/epham2/SHAREPOINT_CLIENT_ID"
SP_CLIENT_SECRET_VAR = "u/epham2/SHAREPOINT_CLIENT_SECRET"

# --- SMApply (verified 2026-11-17 on application 35367543) --------------
SMAPPLY_BASE = "https://nnas.smapply.io/api"
APPLICATION_STAGE_ID = 1361209  # Application Stage
REFUNDED_LOCKED_STATUS_ID = 366216  # Refunded - Locked
LOCKED_STATE = 1

# --- Requests -----------------------------------------------------------
CURRENCIES = {"cad", "usd"}
REASONS = [
    "Duplicate payment",
    "Overpayment",
    "Processing error",
    "Service not provided",
    "Withdrawal",
    "Other",
]
MIN_DETAILS_CHARS = 15
OPEN_STATUSES = ["pending_approval", "approved", "refunded"]
CLOSED_STATUSES = ["rejected", "closed"]

# --- CSV backup to SharePoint (same Graph setup as the volume report) ---
BACKUP_ENABLED = True
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SP_HOST = "nnasca.sharepoint.com"
SP_SITE = "/sites/NNAS-Portal-Sharepoint"
SP_LIBRARY = "Applications"
SP_FOLDER = "zDATA/refund_register"

TIMEOUT = 60


# ======================================================================
# Identity and roles
# ======================================================================


def resolve_caller(ctx_email: str = "") -> str:
    """
    WM_EMAIL is set by Windmill for every job. When an app runs a script it
    may hold the app publisher rather than the person who clicked. If it does,
    fall back to ctx.email, which each app passes in. SETUP step 7 checks
    which one your Windmill version delivers.
    """
    wm = (os.environ.get("WM_EMAIL") or "").strip().lower()
    ctx = (ctx_email or "").strip().lower()
    if wm and wm != APP_PUBLISHER_EMAIL.lower():
        return wm
    return ctx or wm


def roles_of(email: str) -> dict:
    e = (email or "").strip().lower()
    admin = e in ADMINS
    return {
        "email": e,
        "admin": admin,
        "can_request": e in REQUESTERS or admin,
        "can_approve": e in APPROVERS or (admin and ADMINS_CAN_APPROVE),
        "can_process": e in PROCESSORS or admin,
        "can_view_all": admin
        or e in APPROVERS
        or e in PROCESSORS
        or (REQUESTERS_SEE_ALL and e in REQUESTERS),
    }


def require(email: str, capability: str):
    if not roles_of(email).get(capability):
        action = capability.replace("can_", "").replace("_", " ")
        raise PermissionError(f"{email or 'Unknown user'} is not permitted to {action}.")


def solo_test_allowed(email: str, application_id, stripe_mode: str) -> bool:
    """True only for an admin, on an allow-listed app, with a test-mode key."""
    return bool(
        TEST_SOLO
        and roles_of(email)["admin"]
        and LOCK_ALLOWLIST is not None
        and int(application_id) in LOCK_ALLOWLIST
        and stripe_mode == "test"
    )


# ======================================================================
# Database (Windmill data table)
# ======================================================================
#
# Each query is its own round trip, so there are no multi-call transactions.
# Every state change is therefore ONE guarded statement, for example
#   UPDATE ... WHERE id = $1 AND status = 'pending_approval' RETURNING ...
# If the guard fails, no row comes back and nothing changed. The audit row
# is written by a trigger inside that same statement.


def _dt():
    return wmill.datatable(DATATABLE)


def q_all(sql: str, *args) -> list:
    return _dt().query(sql, *args).fetch() or []


def q_one(sql: str, *args):
    return _dt().query(sql, *args).fetch_one()


def is_unique_violation(exc: Exception, constraint: str = "") -> bool:
    msg = str(exc).lower()
    return "duplicate key" in msg and (not constraint or constraint.lower() in msg)


def jsonable(row: dict) -> dict:
    out = {}
    for k, v in dict(row).items():
        if isinstance(v, (datetime, date)):
            out[k] = v.isoformat()
        elif isinstance(v, Decimal):
            out[k] = str(v)
        elif isinstance(v, (dict, list)):
            out[k] = json.dumps(v, default=str)
        else:
            out[k] = v
    return out


# ======================================================================
# Money
# ======================================================================


def to_cents(value) -> int:
    raw = str(value if value is not None else "").replace("$", "").replace(",", "").strip()
    try:
        d = Decimal(raw)
    except InvalidOperation:
        raise ValueError(f"Amount {value!r} is not a number.")
    if d <= 0:
        raise ValueError("Amount must be greater than zero.")
    if d != d.quantize(Decimal("0.01")):
        raise ValueError("Amount has more than two decimal places.")
    return int(d * 100)


def money(cents, currency: str = "") -> str:
    return f"{(cents or 0) / 100:,.2f} {currency.upper()}".strip()


def _ts(epoch):
    return datetime.fromtimestamp(epoch, tz=timezone.utc).isoformat() if epoch else None


# ======================================================================
# SMApply
# ======================================================================

_sm_token = None


def _sm_headers() -> dict:
    global _sm_token
    if not _sm_token:
        _sm_token = wmill.get_variable(SMAPPLY_TOKEN_VAR).strip()
    return {
        "Authorization": f"Bearer {_sm_token}",
        "Accept": "application/json",
        "Content-Type": "application/json",
    }


def _sm_body(resp):
    try:
        return resp.json()
    except Exception:
        return {"text": (resp.text or "")[:400]}


def smapply_snapshot(application_id) -> dict:
    """Application info only. Never returns applicant name or email."""
    application_id = int(application_id)
    r = requests.get(
        f"{SMAPPLY_BASE}/applications/{application_id}/",
        headers=_sm_headers(),
        timeout=TIMEOUT,
    )
    if r.status_code == 404:
        raise ValueError(f"SMApply application {application_id} was not found.")
    if not (200 <= r.status_code < 300):
        raise RuntimeError(
            f"SMApply GET {application_id} failed: HTTP {r.status_code} "
            f"{json.dumps(_sm_body(r), default=str)[:300]}"
        )
    app = r.json()
    status = app.get("status") if isinstance(app.get("status"), dict) else {}
    stage = app.get("current_stage") if isinstance(app.get("current_stage"), dict) else {}
    program = app.get("program") if isinstance(app.get("program"), dict) else {}
    return {
        "application_id": application_id,
        "reference_id": app.get("reference_id"),
        "program": program.get("name"),
        "state": app.get("state"),
        "status_id": status.get("id"),
        "status_name": status.get("name"),
        "stage_id": stage.get("id"),
        "stage_title": stage.get("title"),
    }


def _patch_verify(application_id, step, payload, check, done):
    r = requests.patch(
        f"{SMAPPLY_BASE}/applications/{application_id}/",
        headers=_sm_headers(),
        json=payload,
        timeout=TIMEOUT,
    )
    after = smapply_snapshot(application_id)
    if not (200 <= r.status_code < 300) or not check(after):
        raise RuntimeError(
            f"SMApply step '{step}' failed: HTTP {r.status_code} "
            f"{json.dumps(_sm_body(r), default=str)[:300]}. "
            f"Completed before failure: {done or 'none'}. "
            f"Now: stage={after['stage_title']}, status={after['status_name']}, "
            f"state={after['state']}"
        )
    done.append(step)


def lock_application(application_id) -> dict:
    """
    Stage, then status, then lock. Statuses belong to a stage, so 366216 is
    only accepted inside Application Stage. A stage move can restore an old
    status for that stage, so status is always re-read and set explicitly.
    Each step is verified by a fresh GET. Safe to rerun.
    """
    application_id = int(application_id)
    if LOCK_ALLOWLIST is not None and application_id not in LOCK_ALLOWLIST:
        raise PermissionError(
            f"Application {application_id} is not in LOCK_ALLOWLIST. "
            f"Testing mode is on; set LOCK_ALLOWLIST = None in f/refunds/common "
            f"for production."
        )

    before = smapply_snapshot(application_id)
    done, skipped = [], []

    if before["stage_id"] == APPLICATION_STAGE_ID:
        skipped.append("stage")
    else:
        _patch_verify(
            application_id, "stage", {"current_stage": APPLICATION_STAGE_ID},
            lambda s: s["stage_id"] == APPLICATION_STAGE_ID, done,
        )

    if smapply_snapshot(application_id)["status_id"] == REFUNDED_LOCKED_STATUS_ID:
        skipped.append("status")
    else:
        _patch_verify(
            application_id, "status", {"status": REFUNDED_LOCKED_STATUS_ID},
            lambda s: s["status_id"] == REFUNDED_LOCKED_STATUS_ID, done,
        )

    if smapply_snapshot(application_id)["state"] == LOCKED_STATE:
        skipped.append("lock")
    else:
        _patch_verify(
            application_id, "lock", {"state": LOCKED_STATE},
            lambda s: s["state"] == LOCKED_STATE, done,
        )

    return {
        "completed": done,
        "already_in_place": skipped,
        "before": before,
        "after": smapply_snapshot(application_id),
    }


# ======================================================================
# Stripe (read only)
# ======================================================================


def stripe_init() -> str:
    key = wmill.get_variable(STRIPE_KEY_VAR).strip()
    stripe.api_key = key
    return "test" if "_test_" in key else "live"


def _id(x):
    if x is None or isinstance(x, str):
        return x
    return getattr(x, "id", None)


def _meta(obj) -> dict:
    m = getattr(obj, "metadata", None)
    if not m:
        return {}
    to_dict = getattr(m, "to_dict", None)
    if callable(to_dict):
        return to_dict()
    try:
        return dict(m)
    except Exception:
        return {}


def _invoice_payment_intent(invoice_id: str):
    """
    Find the payment intent behind an invoice. Older API versions expose
    invoice.payment_intent; newer ones list it under invoice.payments.
    """
    try:
        inv = stripe.Invoice.retrieve(invoice_id, expand=["payments"])
    except Exception:
        inv = stripe.Invoice.retrieve(invoice_id)

    pi = _id(getattr(inv, "payment_intent", None))
    if not pi:
        payments = getattr(inv, "payments", None)
        for p in (getattr(payments, "data", None) or []):
            pay = getattr(p, "payment", None)
            cand = _id(getattr(pay, "payment_intent", None)) if pay else None
            if cand:
                pi = cand
                break
    if not pi:
        raise ValueError(
            f"Could not find a payment on invoice {invoice_id}. Open the invoice "
            f"in Stripe and use its payment's pi_ id instead."
        )
    return pi, _meta(inv).get("nnas_application_id")


def resolve_charge(reference: str) -> dict:
    """
    Accepts a payment intent (pi_), charge (ch_ or py_) or invoice (in_) id and
    returns the paid charge. No customer details are returned.
    """
    ref = (reference or "").strip()
    invoice_app_id = None

    if ref.startswith(("ch_", "py_")):
        charge = stripe.Charge.retrieve(ref)
    elif ref.startswith("pi_"):
        pi = stripe.PaymentIntent.retrieve(ref)
        cid = _id(getattr(pi, "latest_charge", None))
        if not cid:
            raise ValueError(f"{ref} has no charge. It was never paid.")
        charge = stripe.Charge.retrieve(cid)
    elif ref.startswith("in_"):
        pi_id, invoice_app_id = _invoice_payment_intent(ref)
        pi = stripe.PaymentIntent.retrieve(pi_id)
        cid = _id(getattr(pi, "latest_charge", None))
        if not cid:
            raise ValueError(f"Invoice {ref} has no completed charge.")
        charge = stripe.Charge.retrieve(cid)
    else:
        raise ValueError(
            "Enter the Stripe pi_, ch_ or in_ id of the original payment "
            "(from the Stripe payment page)."
        )

    if getattr(charge, "status", None) != "succeeded" or not getattr(charge, "paid", False):
        raise ValueError(f"Charge {charge.id} is not a successful payment.")

    amount = int(getattr(charge, "amount", 0) or 0)
    refunded = int(getattr(charge, "amount_refunded", 0) or 0)
    return {
        "charge_id": charge.id,
        "payment_intent_id": _id(getattr(charge, "payment_intent", None)),
        "amount_cents": amount,
        "refunded_cents": refunded,
        "refundable_cents": amount - refunded,
        "currency": (getattr(charge, "currency", "") or "").lower(),
        "created_at": _ts(getattr(charge, "created", None)),
        "livemode": bool(getattr(charge, "livemode", False)),
        "invoice_application_id": invoice_app_id,
    }


def verify_refund(refund_id: str, charge_id: str, expected_cents: int, currency: str) -> dict:
    rid = (refund_id or "").strip()
    if not rid.startswith(("re_", "pyr_")):
        raise ValueError("Enter the Stripe refund id (starts with re_) from the refund in Stripe.")
    rf = stripe.Refund.retrieve(rid)

    rf_charge = _id(getattr(rf, "charge", None))
    if rf_charge != charge_id:
        raise ValueError(
            f"Refund {rid} belongs to charge {rf_charge}, not to this request's "
            f"charge {charge_id}."
        )
    status = getattr(rf, "status", None)
    if status != "succeeded":
        raise ValueError(
            f"Refund {rid} is '{status}', not 'succeeded'. If it is pending, try "
            f"again once Stripe shows it as succeeded."
        )
    amount = int(getattr(rf, "amount", 0) or 0)
    if amount != int(expected_cents):
        raise ValueError(
            f"Refund {rid} is {money(amount, currency)} but the approved amount is "
            f"{money(expected_cents, currency)}. Investigate before closing."
        )
    rf_currency = (getattr(rf, "currency", "") or "").lower()
    if rf_currency != currency.lower():
        raise ValueError(f"Refund currency {rf_currency} does not match {currency}.")
    return {
        "refund_id": rid,
        "amount_cents": amount,
        "status": status,
        "created_at": _ts(getattr(rf, "created", None)),
    }


# ======================================================================
# CSV backup to SharePoint
# ======================================================================

_graph_cache = {}


def _graph_token():
    if "token" in _graph_cache:
        return _graph_cache["token"]
    tenant = wmill.get_variable(AZURE_TENANT_VAR).strip()
    r = requests.post(
        f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token",
        data={
            "client_id": wmill.get_variable(SP_CLIENT_ID_VAR).strip(),
            "client_secret": wmill.get_variable(SP_CLIENT_SECRET_VAR).strip(),
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials",
        },
        timeout=TIMEOUT,
    )
    if r.status_code != 200:
        raise RuntimeError(f"Graph token failed [{r.status_code}]: {r.text[:400]}")
    _graph_cache["token"] = r.json()["access_token"]
    return _graph_cache["token"]


def _graph(method, endpoint, ok=(200, 201), **kw):
    headers = kw.pop("headers", {})
    headers["Authorization"] = f"Bearer {_graph_token()}"
    r = requests.request(method, f"{GRAPH_BASE}{endpoint}", headers=headers, timeout=120, **kw)
    if r.status_code not in ok:
        raise RuntimeError(f"Graph {method} {endpoint} [{r.status_code}]: {r.text[:400]}")
    return r.json() if r.content else None


def _drive_id():
    if "drive" in _graph_cache:
        return _graph_cache["drive"]
    site = _graph("GET", f"/sites/{SP_HOST}:{SP_SITE}")
    drives = _graph("GET", f"/sites/{site['id']}/drives")["value"]
    d = next((x for x in drives if x.get("name") == SP_LIBRARY), None)
    if not d:
        raise RuntimeError(f"Library {SP_LIBRARY} not found.")
    _graph_cache["drive"] = d["id"]
    return d["id"]


def _ensure_folder(path: str):
    cur = ""
    for part in path.strip("/").split("/"):
        cur = f"{cur}/{part}" if cur else part
        r = requests.get(
            f"{GRAPH_BASE}/drives/{_drive_id()}/root:/{quote(cur, safe='/')}",
            headers={"Authorization": f"Bearer {_graph_token()}"},
            timeout=TIMEOUT,
        )
        if r.status_code == 404:
            parent = "/".join(cur.split("/")[:-1])
            ep = (
                f"/drives/{_drive_id()}/root:/{quote(parent, safe='/')}:/children"
                if parent
                else f"/drives/{_drive_id()}/root/children"
            )
            _graph(
                "POST", ep,
                headers={"Content-Type": "application/json"},
                json={"name": part, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"},
                ok=(201,),
            )


def _upload(path: str, data: bytes):
    item = _graph(
        "PUT",
        f"/drives/{_drive_id()}/root:/{quote(path, safe='/')}:/content",
        headers={"Content-Type": "application/octet-stream"},
        data=data,
    )
    return (item or {}).get("webUrl")


def _csv_bytes(rows: list) -> bytes:
    buf = io.StringIO()
    if rows:
        cols = list(rows[0].keys())
        w = csv.DictWriter(buf, fieldnames=cols, extrasaction="ignore")
        w.writeheader()
        for row in rows:
            w.writerow(jsonable(row))
    return buf.getvalue().encode("utf-8-sig")


def backup_register(event: str) -> str:
    """
    Writes the whole register and the audit log as CSV to SharePoint:
        <SP_FOLDER>/refund_register_latest.csv      overwritten each time
        <SP_FOLDER>/refund_audit_latest.csv         overwritten each time
        <SP_FOLDER>/history/refund_register_<ts>_<event>.csv   one per change
    Never raises. A failed backup must not undo a recorded refund, so the
    result is returned as text and shown to the user instead.
    """
    if not BACKUP_ENABLED:
        return "backup disabled"
    try:
        register = _csv_bytes(q_all("SELECT * FROM refund.requests ORDER BY id"))
        audit_log = _csv_bytes(
            q_all("SELECT id, request_id, action, actor, note, at FROM refund.audit ORDER BY id")
        )

        ts = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
        tag = re.sub(r"[^a-z0-9_]+", "_", (event or "event").lower())[:30]
        _ensure_folder(f"{SP_FOLDER}/history")
        url = _upload(f"{SP_FOLDER}/refund_register_latest.csv", register)
        _upload(f"{SP_FOLDER}/refund_audit_latest.csv", audit_log)
        _upload(f"{SP_FOLDER}/history/refund_register_{ts}_{tag}.csv", register)
        return url or "uploaded"
    except Exception as exc:
        return f"BACKUP FAILED: {exc}"


# ======================================================================
# Entrypoint (configuration check only)
# ======================================================================


def main():
    return {
        "requesters": sorted(REQUESTERS),
        "approvers": sorted(APPROVERS),
        "processors": sorted(PROCESSORS),
        "admins": sorted(ADMINS),
        "admins_can_approve": ADMINS_CAN_APPROVE,
        "requesters_see_all": REQUESTERS_SEE_ALL,
        "lock_allowlist": sorted(LOCK_ALLOWLIST) if LOCK_ALLOWLIST is not None else "OFF (production)",
        "test_solo": TEST_SOLO,
        "backup": f"{SP_LIBRARY}/{SP_FOLDER}" if BACKUP_ENABLED else "disabled",
        "datatable": DATATABLE,
        "register_rows": q_one("SELECT count(*) AS n FROM refund.requests"),
    }