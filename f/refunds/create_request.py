# requirements:
# requests
# stripe
# wmill

"""
Windmill path: f/refunds/create_request

Creates a refund request after checking the SMApply application and the
original Stripe payment. Run with dry_run=True first (the app's Check button)
to see what will be saved.

Full refund: leave refund_amount blank. The amount is taken from Stripe
(original amount minus anything already refunded).
"""

from f.refunds.common import (
    CURRENCIES,
    LOCKED_STATE,
    MIN_DETAILS_CHARS,
    REASONS,
    REFUNDED_LOCKED_STATUS_ID,
    backup_register,
    is_unique_violation,
    money,
    q_one,
    require,
    resolve_caller,
    resolve_charge,
    smapply_snapshot,
    stripe_init,
    to_cents,
)


def main(
    application_id: int,
    stripe_reference: str,
    refund_type: str = "full",
    refund_amount: str = "",
    reason: str = "Duplicate payment",
    reason_details: str = "",
    dry_run: bool = True,
    ctx_email: str = "",
):
    who = resolve_caller(ctx_email)
    require(who, "can_request")

    refund_type = (refund_type or "").strip().lower()
    if refund_type not in ("full", "partial"):
        raise ValueError("Refund type must be full or partial.")
    if reason not in REASONS:
        raise ValueError(f"Reason must be one of: {', '.join(REASONS)}.")
    details = (reason_details or "").strip()
    if len(details) < MIN_DETAILS_CHARS:
        raise ValueError(
            f"Explain the reason in at least {MIN_DETAILS_CHARS} characters. "
            f"Do not include card or banking details."
        )

    app = smapply_snapshot(application_id)
    warnings = []
    if app["state"] == LOCKED_STATE:
        warnings.append("This application is already locked in SMApply.")
    if app["status_id"] == REFUNDED_LOCKED_STATUS_ID:
        warnings.append("This application is already at Refunded - Locked.")

    mode = stripe_init()
    ch = resolve_charge(stripe_reference)

    if ch["invoice_application_id"] and str(ch["invoice_application_id"]) != str(app["application_id"]):
        raise ValueError(
            f"This Stripe invoice belongs to application "
            f"{ch['invoice_application_id']}, not {app['application_id']}."
        )
    if ch["currency"] not in CURRENCIES:
        raise ValueError(f"Currency {ch['currency']} is not supported.")
    if ch["refundable_cents"] <= 0:
        raise ValueError("This payment has already been fully refunded in Stripe.")

    if refund_type == "full":
        cents = ch["refundable_cents"]
        if str(refund_amount or "").strip() and to_cents(refund_amount) != cents:
            raise ValueError(
                f"A full refund is {money(cents, ch['currency'])}. Leave the amount "
                f"blank, or choose partial."
            )
        if ch["refunded_cents"] > 0:
            warnings.append(
                f"{money(ch['refunded_cents'], ch['currency'])} was already refunded "
                f"on this payment. Full refund covers the remaining amount."
            )
    else:
        if not str(refund_amount or "").strip():
            raise ValueError("Enter the amount for a partial refund.")
        cents = to_cents(refund_amount)
        if cents > ch["refundable_cents"]:
            raise ValueError(
                f"Requested {money(cents, ch['currency'])} is more than the "
                f"refundable {money(ch['refundable_cents'], ch['currency'])}."
            )
        if cents == ch["refundable_cents"]:
            refund_type = "full"
            warnings.append("Amount equals the full refundable amount, saved as a full refund.")

    existing = q_one(
        "SELECT ref, status FROM refund.requests "
        "WHERE charge_id = $1::text "
        "AND status IN ('pending_approval', 'approved', 'refunded')",
        ch["charge_id"],
    )
    if existing:
        raise ValueError(f"{existing['ref']} is already open for this payment ({existing['status']}).")

    preview = {
        "application": {
            "application_id": app["application_id"],
            "reference_id": app["reference_id"],
            "program": app["program"],
            "stage": app["stage_title"],
            "status": app["status_name"],
            "locked": app["state"] == LOCKED_STATE,
        },
        "payment": {
            "stripe_mode": mode,
            "charge_id": ch["charge_id"],
            "original": money(ch["amount_cents"], ch["currency"]),
            "already_refunded": money(ch["refunded_cents"], ch["currency"]),
            "refundable": money(ch["refundable_cents"], ch["currency"]),
            "paid_on": ch["created_at"],
        },
        "request": {
            "refund_type": refund_type,
            "amount": money(cents, ch["currency"]),
            "reason": reason,
            "requested_by": who,
        },
        "warnings": warnings,
    }

    if dry_run:
        preview["result"] = "Check passed. Nothing saved yet. Press Submit to create the request."
        return preview

    try:
        row = q_one(
            """
            INSERT INTO refund.requests (
                status, stripe_mode,
                application_id, reference_id, program,
                app_stage_at_request, app_status_at_request, app_state_at_request,
                stripe_input_ref, charge_id, payment_intent_id,
                charge_amount_cents, charge_refunded_cents, currency, charge_created_at,
                requested_amount_cents, refund_type, reason, reason_details, requested_by,
                last_action, last_actor, last_note
            ) VALUES (
                'pending_approval', $1::text,
                $2::bigint, $3::text, $4::text,
                $5::text, $6::text, $7::integer,
                $8::text, $9::text, $10::text,
                $11::integer, $12::integer, $13::text, $14::timestamptz,
                $15::integer, $16::text, $17::text, $18::text, $19::text,
                'requested', $19::text, $20::text
            )
            RETURNING id, ref
            """,
            mode,
            app["application_id"], app["reference_id"], app["program"],
            app["stage_title"], app["status_name"], app["state"],
            stripe_reference.strip(), ch["charge_id"], ch["payment_intent_id"],
            ch["amount_cents"], ch["refunded_cents"], ch["currency"], ch["created_at"],
            cents, refund_type, reason, details, who,
            "; ".join(warnings) or None,
        )
    except Exception as exc:
        if is_unique_violation(exc, "refund_one_open_per_charge"):
            raise ValueError("An open request already exists for this payment.")
        raise

    preview["result"] = f"{row['ref']} submitted for approval."
    preview["ref"] = row["ref"]
    preview["backup"] = backup_register("requested")
    return preview