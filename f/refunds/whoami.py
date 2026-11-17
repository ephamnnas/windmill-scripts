# requirements:
# requests
# stripe
# wmill

"""
Windmill path: f/refunds/whoami

Shows who the workflow thinks you are and what you can do. Used by the refund app
header and by SETUP step 6 to confirm identity works inside the app.
"""

import os

from f.refunds.common import APP_PUBLISHER_EMAIL, LOCK_ALLOWLIST, TEST_SOLO, resolve_caller, roles_of


def main(ctx_email: str = ""):
    who = resolve_caller(ctx_email)
    return {
        "you_are": who,
        "roles": roles_of(who),
        "test_mode": LOCK_ALLOWLIST is not None,
        "test_solo": TEST_SOLO,
        "test_application_ids": sorted(LOCK_ALLOWLIST) if LOCK_ALLOWLIST is not None else [],
        "debug": {
            "WM_EMAIL": os.environ.get("WM_EMAIL"),
            "ctx_email": ctx_email,
            "app_publisher": APP_PUBLISHER_EMAIL,
        },
    }