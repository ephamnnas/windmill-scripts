# requirements:
# requests
# stripe
# wmill

"""
Windmill path: f/refunds/list_requests

Rows for the app tables. view = open | closed | all.
Users who cannot view all only see their own requests.
"""

from f.refunds.common import money, q_all, resolve_caller, roles_of

VIEWS = {
    "open": "status IN ('pending_approval', 'approved', 'refunded')",
    "closed": "status IN ('rejected', 'closed')",
    "all": "TRUE",
}


def main(view: str = "open", ctx_email: str = ""):
    who = resolve_caller(ctx_email)
    r = roles_of(who)
    if not (r["can_request"] or r["can_view_all"]):
        raise PermissionError(f"{who or 'Unknown user'} has no access to refunds.")

    where = VIEWS.get((view or "open").lower(), VIEWS["open"])
    sql = f"""
        SELECT id, ref, status, stripe_mode, application_id, reference_id, program,
               requested_amount_cents, currency, refund_type, reason, reason_details,
               requested_by, requested_at, decided_by, decided_at, decision_comment,
               charge_id, stripe_refund_id, refund_verified_at,
               lock_completed_at, lock_error
        FROM refund.requests
        WHERE {where}
    """
    if r["can_view_all"]:
        rows = q_all(sql + " ORDER BY id DESC LIMIT 500")
    else:
        rows = q_all(sql + " AND requested_by = $1::text ORDER BY id DESC LIMIT 500", who)

    for row in rows:
        row["amount"] = money(row.pop("requested_amount_cents", 0), row.get("currency") or "")
    return rows