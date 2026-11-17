# requirements:
# requests
# stripe
# wmill

"""
Windmill path: f/refunds/complete_and_lock

The "Issue refund" button, used by the processor after the refund was made
manually in Stripe.

  1. Verifies the Stripe refund (re_ id): right charge, succeeded, exact
     approved amount, same currency. Records it, so the refund is on record
     even if step 2 fails.
  2. Moves the SMApply application to Application Stage, sets Refunded -
     Locked, and locks it. Marks the request closed.
  3. Emails the approver and the requester that the refund was issued.

If step 2 fails, the request stays 'refunded' with the error saved. Fix the
cause and press the button again with the refund id blank; the refund is not
checked twice.
"""

from f.refunds.common import (
    backup_register,
    is_unique_violation,
    lock_application,
    money,
    notify,
    q_one,
    require,
    resolve_caller,
    roles_of,
    solo_test_allowed,
    stripe_init,
    verify_refund,
)


def main(request_id: int, stripe_refund_id: str = "", ctx_email: str = ""):
    who = resolve_caller(ctx_email)
    rid = (stripe_refund_id or "").strip()
    request_id = int(request_id)

    row = q_one("SELECT * FROM refund.requests WHERE id = $1::bigint", request_id)
    if not row:
        raise ValueError(f"Request {request_id} not found.")

    solo = solo_test_allowed(who, row["application_id"], row["stripe_mode"])
    if not roles_of(who)["can_process"] and not solo:
        require(who, "can_process")
    if who in (row["requested_by"], row["decided_by"]) and not solo:
        raise PermissionError(
            "You requested or approved this refund, so someone else must issue it."
        )

    ref = row["ref"]
    steps = []

    # ---- 1. Record the Stripe refund ------------------------------------
    if row["status"] == "approved":
        stripe_init()
        rf = verify_refund(rid, row["charge_id"], row["requested_amount_cents"], row["currency"])
        try:
            updated = q_one(
                """
                UPDATE refund.requests
                SET status = 'refunded', stripe_refund_id = $2::text,
                    refunded_amount_cents = $3::integer, refund_status = $4::text,
                    refund_created_at = $5::timestamptz, refund_verified_by = $6::text,
                    refund_verified_at = now(),
                    last_action = 'refund_verified', last_actor = $6::text,
                    last_note = $7::text, updated_at = now()
                WHERE id = $1::bigint AND status = 'approved'
                RETURNING id
                """,
                request_id, rf["refund_id"], rf["amount_cents"], rf["status"],
                rf["created_at"], who, "solo test" if solo else None,
            )
        except Exception as exc:
            if is_unique_violation(exc):
                raise ValueError("That Stripe refund id is already recorded on another request.")
            raise
        if not updated:
            raise RuntimeError(f"{ref} changed while processing. Refresh and retry.")
        steps.append("refund_verified")
    elif row["status"] == "refunded":
        if rid and rid != row["stripe_refund_id"]:
            raise ValueError(
                f"{ref} already has refund {row['stripe_refund_id']} on record. "
                f"Leave the refund id blank to retry the lock."
            )
        steps.append("lock_retry")
    else:
        raise ValueError(f"{ref} is {row['status']}. Only approved requests can be completed.")

    # ---- 2. Lock in SMApply ---------------------------------------------
    try:
        result = lock_application(row["application_id"])
    except Exception as exc:
        q_one(
            """
            UPDATE refund.requests
            SET lock_error = $2::text, last_action = 'lock_failed',
                last_actor = $3::text, last_note = NULL, updated_at = now()
            WHERE id = $1::bigint AND status = 'refunded'
            RETURNING id
            """,
            request_id, str(exc)[:1000], who,
        )
        backup_register("lock_failed")
        raise RuntimeError(
            f"Refund is recorded for {ref}, but the SMApply lock failed: {exc} "
            f"Fix the cause and press Lock application again with the refund id blank."
        )

    closed = q_one(
        """
        UPDATE refund.requests
        SET status = 'closed', lock_completed_at = now(), lock_by = $2::text,
            lock_error = NULL, app_stage_at_close = $3::text,
            app_status_at_close = $4::text,
            last_action = 'locked', last_actor = $2::text, last_note = $5::text,
            updated_at = now()
        WHERE id = $1::bigint AND status = 'refunded'
        RETURNING id
        """,
        request_id, who,
        result["after"]["stage_title"], result["after"]["status_name"],
        f"completed {result['completed']}, already {result['already_in_place']}",
    )
    if not closed:
        raise RuntimeError(f"{ref} was locked in SMApply but its status changed meanwhile. Check the register.")

    amount = money(row["requested_amount_cents"], row["currency"])
    alert = notify(
        [row["decided_by"], row["requested_by"]],
        f"Refund {ref} issued: {row['reference_id']}",
        [
            f"{ref} has been refunded and closed by {who}.",
            f"Application: {row['reference_id']}",
            f"Amount refunded: {amount}",
            f"Stripe refund: {row['stripe_refund_id'] or rid}",
            "The application is now locked at Refunded - Locked in SMApply.",
        ],
    )

    return {
        "result": f"{ref} refund issued. Application {row['reference_id']} is locked at Refunded - Locked.",
        "steps": steps + ["locked"],
        "smapply": result["after"],
        "backup": backup_register("closed"),
        "alert": alert,
    }