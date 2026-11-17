# requirements:
# requests
# stripe
# wmill

"""
Windmill path: f/refunds/find_payments

Step 1 of a new request. The requester gives an application number (or SMApply
ID) and, optionally, the email the payment was made with. Returns the Stripe
payments found for that email so the requester can pick one. No Stripe ids
need to be known in advance.

Email used:
  - the one entered, if any (for payments made under a different email)
  - otherwise the applicant email on the SMApply application

The email is shown back masked and is never stored.
"""

from f.refunds.common import (
    LOCKED_STATE,
    clean_email,
    find_charges_by_email,
    mask_email,
    q_all,
    require,
    resolve_application_id,
    resolve_caller,
    smapply_application,
    stripe_init,
)


def main(application: str, email: str = "", ctx_email: str = ""):
    who = resolve_caller(ctx_email)
    require(who, "can_request")

    app, applicant_email, created_at = smapply_application(resolve_application_id(application))
    entered = clean_email(email)
    use = entered or applicant_email
    if not use:
        raise ValueError(
            "This application has no applicant email on file. Enter the email "
            "the payment was made with."
        )

    notes = []
    if entered and applicant_email and entered != applicant_email:
        notes.append(
            "The email entered is different from the applicant email on this "
            "application. The approver will see this."
        )
    if app["state"] == LOCKED_STATE:
        notes.append("This application is already locked in SMApply.")

    mode = stripe_init()
    payments, stats = find_charges_by_email(use, scan_since_iso=created_at)

    open_by_charge = {
        r["charge_id"]: r
        for r in q_all(
            "SELECT charge_id, ref, status FROM refund.requests "
            "WHERE status IN ('pending_approval', 'approved', 'refunded')"
        )
    }
    for p in payments:
        hit = open_by_charge.get(p["payment_id"])
        p["open_request"] = f"{hit['ref']} ({hit['status']})" if hit else None
        p["selectable"] = p["refundable_cents"] > 0 and not hit

    if not payments:
        notes.append(
            "No Stripe payments were found for this email. The payment may have "
            "been made with a different email: enter that email and search again."
        )
        if stats["scan_capped"]:
            notes.append("The search stopped at its limit before covering all payments. Contact IT.")

    return {
        "application": {
            "application_id": app["application_id"],
            "reference_id": app["reference_id"],
            "program": app["program"],
            "stage": app["stage_title"],
            "status": app["status_name"],
            "locked": app["state"] == LOCKED_STATE,
        },
        "email_used": mask_email(use),
        "email_source": "entered" if entered else "application",
        "stripe_mode": mode,
        "payments": payments,
        "notes": notes,
    }