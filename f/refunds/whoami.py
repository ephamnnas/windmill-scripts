# requirements:
# requests
# stripe
# wmill

"""
Windmill path: f/refunds/whoami

Shows who the workflow thinks you are and what you can do. Used by both apps
and by SETUP step 7 to confirm identity works inside an app.
"""

import os

from f.refunds.common import APP_PUBLISHER_EMAIL, resolve_caller, roles_of


def main(ctx_email: str = ""):
    who = resolve_caller(ctx_email)
    return {
        "you_are": who,
        "roles": roles_of(who),
        "debug": {
            "WM_EMAIL": os.environ.get("WM_EMAIL"),
            "ctx_email": ctx_email,
            "app_publisher": APP_PUBLISHER_EMAIL,
        },
    }