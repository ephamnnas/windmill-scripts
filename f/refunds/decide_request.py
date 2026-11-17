# requirements:
# requests
# stripe
# wmill

"""
Windmill path: f/refunds/decide_request

Approve or reject a pending request. On approve, the payment is re-checked in
Stripe so a payment refunded elsewhere since the request cannot be approved.
Rejections need a comment. Nobody can approve their own request.
"""

from f.refunds.common import (
    backup_register,
    money,
    q_one,
    require,
    resolve_caller,
    resolve_charge,
    roles_of,
    solo_test_allowed,
    stripe_init,
)


def main(request_id: int, decision: str, comment: str = "", ctx_email: str = ""):
    who = resolve_caller(ctx_email)
    decision = (decision or "").strip().lower()
    if decision not in ("approve", "reject"):
        raise ValueError("Decision must be approve or reject.")
    comment = (comment or "").strip()
    if decision == "reject" and not comment:
        raise ValueError("Add a comment explaining the rejection.")

    row = q_one("SELECT * FROM refund.requests WHERE id = $1::bigint", int(request_id))
    if not row:
        raise ValueError(f"Request {request_id} not found.")

    solo = solo_test_allowed(who, row["application_id"], row["stripe_mode"])
    if not roles_of(who)["can_approve"] and not solo:
        require(who, "can_approve")
    if row["status"] != "pending_approval":
        raise ValueError(f"{row['ref']} is {row['status']}, not awaiting approval.")
    if row["requested_by"] == who and not solo:
        raise PermissionError("You cannot approve or reject your own request.")

    notes = []
    if decision == "approve":
        stripe_init()
        ch = resolve_charge(row["charge_id"])
        if ch["refundable_cents"] < row["requested_amount_cents"]:
            raise ValueError(
                f"Stripe now shows only {money(ch['refundable_cents'], row['currency'])} "
                f"refundable on this payment, less than the requested "
                f"{money(row['requested_amount_cents'], row['currency'])}. "
                f"Check whether it was refunded elsewhere."
            )
        notes.append(f"refundable at approval {money(ch['refundable_cents'], row['currency'])}")
    if solo:
        notes.append("solo test")

    new_status = "approved" if decision == "approve" else "rejected"
    updated = q_one(
        """
        UPDATE refund.requests
        SET status = $2::text, decided_by = $3::text, decided_at = now(),
            decision_comment = $4::text,
            last_action = $2::text, last_actor = $3::text, last_note = $5::text,
            updated_at = now()
        WHERE id = $1::bigint AND status = 'pending_approval'
        RETURNING id, ref
        """,
        int(request_id), new_status, who, comment or None, "; ".join(notes) or None,
    )
    if not updated:
        raise RuntimeError(f"{row['ref']} changed while deciding. Refresh and retry.")

    return {
        "result": f"{updated['ref']} {new_status}.",
        "next": "Issue the refund in Stripe, then press Lock application." if new_status == "approved" else None,
        "backup": backup_register(new_status),
    }