#requirements:
#requests
#wmill

"""
NNAS: stamp trigger date into an SMApply custom field
=====================================================

On webhook trigger, writes the current date (execution time) into the custom
field selected by additional_content's "update_field":

    "update_field": "submission"  ->  field 153648  (original)
    "update_field": "payment"     ->  field 152033

via PATCH /api/applications/{id}/. Accepts the same inputs as the intake
scripts so the same webhook payload binds cleanly; only APPLICATION_ID is
required. An optional UPDATE_FIELD input overrides the blob for manual runs.
An unknown update_field errors rather than stamping the wrong field.

The value written is the moment this script RUNS (the trigger time), captured
here -- not the webhook's date_created. Set USE_DATE_CREATED = True to use the
payload's creation time instead.
"""

import json
from datetime import datetime, timezone

import wmill
import requests


# =========================================================
# CONFIGURATION
# =========================================================

BASE = "https://nnas.smapply.io"
SMAPPLY_API_TOKEN = wmill.get_variable("u/epham2/SMAPPLY_ACCESS_TOKEN")

# Which custom field gets stamped is selected by additional_content's
# "update_field" value (or the UPDATE_FIELD input override):
#   "submission" -> 153648  (the original field)
#   "payment"    -> 152033
FIELD_MAP = {
    "submission": 153648,
    "payment":    152033,
    "issueddate": 153677,
}
# Used only when no update_field is supplied at all (keeps the original
# behavior). Set to None to instead make update_field mandatory.
DEFAULT_UPDATE_FIELD = "submission"

# Value format. SMApply date fields typically expect "%Y-%m-%d"; switch to a
# full ISO timestamp here if 153648 is a datetime field.
DATE_FORMAT = "%Y-%m-%d"          # e.g. "%Y-%m-%dT%H:%M:%S" for datetime

# Trigger time is captured in UTC on the Windmill worker. NNAS is Pacific; if
# the field should reflect local date, apply an offset or use zoneinfo here.
USE_UTC = True

# If True, use the webhook's date_created (application creation time) instead
# of the execution time. Default False = "time of trigger" as asked.
USE_DATE_CREATED = False

HEADERS = {
    "Authorization": f"Bearer {SMAPPLY_API_TOKEN}",
    "Content-Type": "application/json",
}


# =========================================================
# HELPERS
# =========================================================

def _parse_date_created(additional_content, kwargs) -> datetime | None:
    """Best-effort parse of the webhook's date_created ('2026 Jun 7 03:43 AM')."""
    raw = kwargs.get("date_created") or kwargs.get("DATE_CREATED")
    if not raw and additional_content:
        try:
            blob = additional_content if isinstance(additional_content, dict) \
                else json.loads(additional_content)
            if isinstance(blob, str):
                blob = json.loads(blob)
            raw = blob.get("date_created")
        except Exception:
            raw = None
    if not raw:
        return None
    for fmt in ("%Y %b %d %I:%M %p", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(str(raw).strip(), fmt)
        except ValueError:
            continue
    return None


def _load_blob(additional_content):
    """Tolerant parse of additional_content into a dict, or None."""
    if not additional_content:
        return None
    if isinstance(additional_content, dict):
        return additional_content
    s = str(additional_content).strip()
    try:
        blob = json.loads(s)
        if isinstance(blob, str):
            blob = json.loads(blob)
        return blob if isinstance(blob, dict) else None
    except Exception:
        return None


def resolve_field_id(update_field_override, additional_content, kwargs):
    """
    (field_id, resolved_key). Precedence: explicit UPDATE_FIELD input, then
    additional_content["update_field"], then DEFAULT_UPDATE_FIELD. An unknown
    key raises rather than silently stamping the wrong field.
    """
    key = (update_field_override or "").strip().lower()
    if not key:
        blob = _load_blob(additional_content)
        if blob:
            # update_field lives inside application_info in the NNAS webhook,
            # but tolerate a top-level placement too.
            ai = blob.get("application_info") or {}
            key = str(ai.get("update_field")
                      or blob.get("update_field") or "").strip().lower()
    if not key:
        key = kwargs.get("update_field", "")  # tolerate a top-level webhook key
        key = str(key or "").strip().lower()
    if not key:
        key = (DEFAULT_UPDATE_FIELD or "").strip().lower()

    if key in FIELD_MAP:
        return FIELD_MAP[key], key

    raise RuntimeError(
        f"Cannot resolve update_field. Got '{key or '(empty)'}'; expected one of: "
        f"{', '.join(sorted(FIELD_MAP))}. Set it in additional_content "
        "(\"update_field\": \"payment\"|\"submission\") or via the UPDATE_FIELD input."
    )


def trigger_value(additional_content, kwargs) -> str:
    """The string written to the custom field."""
    if USE_DATE_CREATED:
        dt = _parse_date_created(additional_content, kwargs) or \
            (datetime.now(timezone.utc) if USE_UTC else datetime.now())
    else:
        dt = datetime.now(timezone.utc) if USE_UTC else datetime.now()
    return dt.strftime(DATE_FORMAT)


# =========================================================
# MAIN
# =========================================================

def main(id: int = None,
         reference_id: str = "",
         user: str = "",
         additional_content: str = "",
         update_field: str = "",
         **kwargs):
    """
    Inputs bind to the webhook's own field names:
        id                 -> the application ID (PATCH target); the webhook
                              sample carries the numeric app id here, not in an
                              "APPLICATION_ID" key
        reference_id       -> NNAS-...      (logging only)
        user               -> applicant     (logging only)
        additional_content -> the JSON blob (carries application_info.update_field)
        update_field       -> optional direct override for manual runs

    Only `id` is required. Back-compat: an APPLICATION_ID passed via kwargs is
    still honored.
    """
    application_id = id if id not in (None, "") else kwargs.get("APPLICATION_ID")
    if not application_id:
        return {"status": "error",
                "message": "No application id. Bind the webhook 'id' field."}

    field_id, field_key = resolve_field_id(update_field, additional_content, kwargs)
    value = trigger_value(additional_content, kwargs)

    url = f"{BASE}/api/applications/{application_id}/"
    payload = {"custom_fields": [{"id": field_id, "value": value}]}

    response = requests.patch(url, json=payload, headers=HEADERS, timeout=120)

    if response.status_code in (200, 201, 204):
        print(f"Set custom field {field_id} ({field_key}) = {value} "
              f"on application {application_id}")
        return {
            "status": "ok",
            "application_id": application_id,
            "reference_id": reference_id,
            "update_field": field_key,
            "custom_field_id": field_id,
            "value": value,
        }

    print(f"Error updating custom field [{response.status_code}]: {response.text[:1000]}")
    return {
        "status": "error",
        "application_id": application_id,
        "update_field": field_key,
        "custom_field_id": field_id,
        "value": value,
        "http_status": response.status_code,
        "response": response.text[:1000],
    }