#requirements:
#stripe
#requests
#wmill

"""
Stripe webhook target — translation payment received.

On invoice.paid, set the application status to Paid_Translation (366865).
That is the whole job. A SMApply automation watching for that status sends the
email; nothing here touches mail.

Register the Windmill webhook URL in Stripe -> Developers -> Webhooks for:
    invoice.paid
    invoice.payment_failed
    invoice.voided
    invoice.marked_uncollectible

Only invoice.paid changes anything. The others are subscribed so the run history
records what happened to a charge — an unpaid invoice should be visible, not
silent.

WHY NO prior_status LOOKUP
--------------------------
An earlier design restored whatever status the application carried before
billing, from custom field 153838. Moving to a fixed Paid_Translation status
removes that dependency entirely: nothing has to be populated ahead of time, so
there is no way for a paid applicant to end up stranded because a field was
blank. 153838 is now unused by this flow.

SECURITY
--------
Stripe signature verification needs the exact raw request bytes. Windmill parses
incoming JSON into script arguments, so those bytes are gone by the time this
runs. Rather than verify a signature we cannot reconstruct, this takes only the
event id from the body and re-fetches the event from Stripe with our own key. A
forged payload is inert: either the id does not exist in our account, or it does
and the authoritative data comes back from Stripe.

requirements:
stripe
requests
wmill
"""

import json

import requests
import stripe
import wmill

# ======================================================================
# CONFIG
# ======================================================================

BASE = "https://nnas.smapply.io/api"
TIMEOUT = 60

TOKEN_VAR = "u/epham2/SMAPPLY_ACCESS_TOKEN"

# Restricted Stripe key (rk_...), not the one SMApply uses for application fees.
STRIPE_VAR = "u/epham2/STRIPE_SECRET_KEY"

# On payment, clear the billing flag: if the application is still on
# Bill_for_translation, move it to Waiting for Required Documents.
#
# Conditional on purpose. If someone has already moved the application on, this
# leaves it alone rather than dragging it backwards.
#
# ORDERING CAVEAT — worth reading before trusting this.
# Statuses are scoped to the application's CURRENT stage. 351474 lives on stage
# 1362840 (Document Intake & Review). This write happens BEFORE the task
# completion, because completing the task hands control to your automation,
# which may move the stage — and a write issued afterwards could land
# mid-transition.
#
# That is only safe if 366184 and 351474 are on the SAME stage. Check with
# `GET /api/stages/1362840/statuses/`: if both appear, this works as written.
# If Bill_for_translation lives on a different stage, this PATCH will be
# rejected as "status is not valid" and the status change belongs in your
# SMApply automation instead — after its stage move, where 351474 is valid.
# The failure is non-fatal and says exactly that, so one test run will tell you.
CLEAR_STATUS_FROM = 366184    # "Bill_for_translation"
CLEAR_STATUS_TO = 351474      # "Waiting for Required Documents"

BILL_STATUS_ID = 366184       # "Bill_for_translation" — the expected prior
                              # state. A mismatch is logged, never blocked.

# SMApply suppresses workflow automations on API writes unless you ask for
# them. Without this header the status changes and nothing downstream fires —
# which is exactly what "status set correctly but no email" looks like.
#
# Confirmed previously on task PATCH; the same header governs application PATCH.
APPLY_AUTOMATIONS = "full"

# --- Triggering the applicant email ------------------------------------
#
# A status change made through the API does NOT fire SMApply automations. Eight
# header and query variants were tested against /applications/{id}/ and none
# worked, while the same change made by hand in the UI does. Route mapping also
# confirmed SMApply exposes no applicant-messaging endpoint — every candidate
# 404s except /notes/, which is internal notes, not email.
#
# What DOES work through the API is task completion with `apply-automations`.
# That is confirmed behaviour, not a guess. So: give the application an admin
# task, mark it complete here, and trigger the SMApply automation off the task
# completion rather than the status change.
#
# 0 = skip. Set once the task exists — see 27_PAYMENT_SETUP.md.
PAID_TASK_ID = 2095855       # "Translation Paid" — completing this is what
                             # fires the SMApply automation

# Optional: leave an internal note on the application recording the payment.
# Staff-visible in SMApply, and a useful audit trail independent of Stripe.
WRITE_NOTE = True

HANDLED = {
    "invoice.paid": "paid",
    "invoice.payment_succeeded": "paid",
    "invoice.payment_failed": "failed",
    "invoice.voided": "voided",
    "invoice.marked_uncollectible": "uncollectible",
}


# ======================================================================
# Helpers
# ======================================================================

def _meta(obj) -> dict:
    """
    Metadata off a Stripe object, as a real dict.

    stripe-python 15 returns a StripeObject here. It looks dict-like but is not
    — `.get()` raises AttributeError rather than returning None, and attribute
    access on a missing key raises too. Convert once, use dict semantics after.
    """
    m = getattr(obj, "metadata", None)
    if not m:
        return {}
    if isinstance(m, dict):
        return dict(m)
    to_dict = getattr(m, "to_dict", None)
    if callable(to_dict):
        return to_dict()
    try:
        return dict(m)
    except Exception:
        return {}


def _sm_headers(token: str, automations: bool = False) -> dict:
    """
    `automations=True` adds the header that lets SMApply's own workflow
    automations run in response to this write. Reads never need it; the status
    write does, or the Paid_Translation email never goes out.
    """
    h = {"Authorization": f"Bearer {token}", "Accept": "application/json",
         "Content-Type": "application/json"}
    if automations and APPLY_AUTOMATIONS:
        h["apply-automations"] = APPLY_AUTOMATIONS
    return h


def _body(resp: requests.Response):
    """SMApply 500s return HTML. Never assume JSON."""
    try:
        return resp.json()
    except Exception:
        return {"text": (resp.text or "")[:800]}


def get_task(token: str, application_id: int, task_id: int) -> dict:
    r = requests.get(f"{BASE}/applications/{application_id}/tasks/{task_id}/",
                     headers=_sm_headers(token), timeout=TIMEOUT)
    if not (200 <= r.status_code < 300):
        raise RuntimeError(
            f"GET task {task_id} on application {application_id} failed: "
            f"HTTP {r.status_code}. If this is a 404 the task is a REVIEW task "
            f"and cannot be completed through this endpoint.")
    return _body(r)


def get_status(token: str, application_id: int):
    r = requests.get(f"{BASE}/applications/{application_id}/",
                     headers=_sm_headers(token), timeout=TIMEOUT)
    if not (200 <= r.status_code < 300):
        raise RuntimeError(f"GET application {application_id} failed: "
                           f"HTTP {r.status_code} "
                           f"{json.dumps(_body(r), default=str)[:400]}")
    st = _body(r).get("status")
    return st.get("id") if isinstance(st, dict) else st


def set_status(token: str, application_id: int, status_id: int):
    """
    Bare int only — {"status": {"id": N}} returns "status must be an integer".

    Sent with `apply-automations` so SMApply's own automations fire. Without
    that header the write succeeds and nothing downstream happens — the status
    reads correctly in the UI while the applicant never gets an email.

    Statuses are scoped to the application's CURRENT stage. If the application
    moved stages between billing and payment, Paid_Translation will not exist
    there and this returns 400 "status is not valid". That is raised rather than
    swallowed: the applicant has paid, and someone has to unblock them.
    """
    r = requests.patch(f"{BASE}/applications/{application_id}/",
                       headers=_sm_headers(token, automations=True),
                       json={"status": int(status_id)}, timeout=TIMEOUT)
    if not (200 <= r.status_code < 300):
        raise RuntimeError(f"status write failed: HTTP {r.status_code} "
                           f"{json.dumps(_body(r), default=str)[:400]}")
    st = _body(r).get("status")
    return st.get("id") if isinstance(st, dict) else st


def complete_task(token: str, application_id: int, task_id: int) -> dict:
    """
    Mark an application task complete, with automations enabled.

    This is the lever that actually fires SMApply automations from the API.
    Constraints, all confirmed the hard way:

      * the payload is exactly {"status": 3} — a bare int
      * `data` and `status` cannot go in the same PATCH
      * `completed_at` is read-only; including it returns an error
      * without `apply-automations`, completion automations do not run

    Non-fatal: the status change is the durable outcome, and an already-complete
    task is not an error.
    """
    url = f"{BASE}/applications/{application_id}/tasks/{task_id}/"
    r = requests.patch(url, headers=_sm_headers(token, automations=True),
                       json={"status": 3}, timeout=TIMEOUT)
    body = _body(r)
    ok = 200 <= r.status_code < 300
    if not ok:
        print(f"WARN: could not complete task {task_id}: HTTP {r.status_code} "
              f"{json.dumps(body, default=str)[:300]}")
    else:
        print(f"Task {task_id} marked complete with apply-automations — this "
              f"is what fires the applicant email.")
    return {"task_id": task_id, "http_status": r.status_code, "completed": ok}


def write_note(token: str, application_id: int, text: str) -> dict:
    """
    Leave an internal note recording the payment.

    /applications/{id}/notes/ was the only messaging-adjacent route that exists,
    and it allows POST. The field name is undocumented, so try the plausible
    ones. Purely an audit convenience — never fatal.
    """
    url = f"{BASE}/applications/{application_id}/notes/"
    for key in ("content", "text", "note", "body", "message"):
        r = requests.post(url, headers=_sm_headers(token),
                          json={key: text}, timeout=TIMEOUT)
        if 200 <= r.status_code < 300:
            print(f"Internal note written (field {key!r}).")
            return {"written": True, "field": key}
    print(f"WARN: could not write an internal note; last status "
          f"{r.status_code}. Not important — skipping.")
    return {"written": False}


# ======================================================================
# Entrypoint
# ======================================================================

def main(id: str = "", type: str = "", event_id: str = "", **kwargs):
    """
    Stripe posts the whole event; Windmill maps top-level keys to arguments, so
    `id` and `type` arrive populated. `event_id` is a manual-replay override.
    Everything else is swallowed by **kwargs and ignored.
    """
    evt_id = event_id or id
    if not evt_id:
        return {"ok": False, "reason": "no event id in payload"}

    stripe.api_key = wmill.get_variable(STRIPE_VAR)
    try:
        event = stripe.Event.retrieve(evt_id)      # authoritative
    except Exception as exc:
        return {"ok": False, "reason": f"event {evt_id} not retrievable: {exc}"}

    etype = getattr(event, "type", None)
    if etype not in HANDLED:
        return {"ok": True, "ignored": etype, "event_id": evt_id}

    obj = event.data.object
    meta = _meta(obj)

    # Only ours. SMApply pays application fees via PaymentIntents and creates no
    # Invoice objects, so this should never fire for them — the filter is what
    # keeps that an invariant rather than an assumption.
    if meta.get("nnas_service") != "translation":
        return {"ok": True, "ignored": "not a translation invoice",
                "stripe_invoice_id": getattr(obj, "id", None)}

    application_id = meta.get("nnas_application_id")
    result = {
        "ok": True,
        "event_id": evt_id,
        "type": etype,
        "application_id": application_id,
        "reference_id": meta.get("nnas_reference_id"),
        "applicant_email": meta.get("nnas_applicant_email"),
        "pages": meta.get("nnas_pages"),
        "stripe_invoice_id": getattr(obj, "id", None),
        "invoice_number": getattr(obj, "number", None),
        "amount_paid_cad": (getattr(obj, "amount_paid", None) or 0) / 100,
        "tax_cad": (getattr(obj, "tax", None) or 0) / 100,
    }

    if HANDLED[etype] != "paid":
        result["note"] = (f"{etype}: charge not settled, status left on "
                          f"Bill_for_translation")
        print(json.dumps(result, indent=2, default=str))
        return result

    if not application_id:
        result.update({"ok": False,
                       "reason": "no nnas_application_id in invoice metadata"})
        return result

    token = wmill.get_variable(TOKEN_VAR)
    application_id = int(application_id)

    if not PAID_TASK_ID:
        result.update({"ok": False,
                       "reason": "PAID_TASK_ID is unset — nothing would fire"})
        print("PAID_TASK_ID is 0. Set it to the 'Translation Paid' task id.")
        return result

    # --- idempotency keys on the TASK, not the status --------------------
    # The task is the trigger, so its completion state is the honest record of
    # whether this event has already been processed. Stripe retries on any
    # non-2xx, and completing the task twice would re-fire the automation —
    # a second applicant email, and possibly a second stage move.
    task = get_task(token, application_id, PAID_TASK_ID)
    result["task_status_before"] = task.get("status")

    if task.get("status") == 3:
        result.update({"task_completed": False, "already_applied": True})
        print(f"Task {PAID_TASK_ID} is already complete on application "
              f"{application_id}. Nothing to do.")
        print(json.dumps(result, indent=2, default=str))
        return result

    status_now = get_status(token, application_id)
    result["status_before"] = status_now
    if BILL_STATUS_ID and status_now != BILL_STATUS_ID:
        # Not a reason to stop — the applicant has paid and that must be
        # recorded — but worth surfacing: someone moved the application while
        # the invoice was outstanding.
        print(f"NOTE: application {application_id} was on {status_now}, not "
              f"{BILL_STATUS_ID} (Bill_for_translation).")
        result["unexpected_prior_status"] = status_now

    # --- order matters ----------------------------------------------------
    # Everything that is ours happens BEFORE the task completion, because
    # completing the task hands control to SMApply's automation — which may
    # move the application to another stage. Writes issued after that could
    # land mid-transition, or against a stage where the value is no longer
    # valid.

    if CLEAR_STATUS_TO and status_now == CLEAR_STATUS_FROM:
        try:
            landed = set_status(token, application_id, CLEAR_STATUS_TO)
            result.update({"status_changed": True, "status_after": landed})
            print(f"Status cleared {CLEAR_STATUS_FROM} -> {landed} "
                  f"(Waiting for Required Documents).")
        except Exception as exc:
            # Non-fatal. The task completion is what matters, and this most
            # likely means 351474 is not configured on the stage the
            # application is currently in — in which case the status change
            # belongs in the automation, after its stage move.
            result.update({"status_changed": False,
                           "status_error": str(exc)[:300]})
            print(f"WARN: could not set status {CLEAR_STATUS_TO}: {exc}")
            if "not valid" in str(exc).lower():
                print("      That error means the status is not configured on "
                      "this application's current stage. Move this step into "
                      "the SMApply automation, after its stage move.")
    else:
        result["status_changed"] = False
        if status_now != CLEAR_STATUS_FROM:
            print(f"Status left as {status_now} — not on "
                  f"{CLEAR_STATUS_FROM}, so nothing to clear.")

    if WRITE_NOTE:
        result["note"] = write_note(
            token, application_id,
            f"Translation payment received. "
            f"${result.get('amount_paid_cad', 0):,.2f} CAD "
            f"({result.get('pages')} pages), Stripe invoice "
            f"{result.get('invoice_number') or result.get('stripe_invoice_id')}.")

    # --- the trigger, last -----------------------------------------------
    result["task"] = complete_task(token, application_id, PAID_TASK_ID)
    if not result["task"].get("completed"):
        result["ok"] = False
        result["reason"] = ("task completion failed — the applicant paid but "
                            "no automation fired")
        print(f"ERROR: {result['reason']}")

    print(json.dumps(result, indent=2, default=str))
    return result