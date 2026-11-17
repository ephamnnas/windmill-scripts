# requirements:
# requests
# stripe
# wmill

"""
Windmill path: f/refunds/backup_csv

Writes the full register and audit log to SharePoint as CSV. Every change
already does this; schedule this script daily as a safety net.
"""

from f.refunds.common import backup_register


def main():
    result = backup_register("scheduled")
    if result.startswith("BACKUP FAILED"):
        raise RuntimeError(result)
    return {"backup": result}