#requirements:
#wmill
#requests

import wmill
import requests


# =========================================================
# CONFIGURATION
# =========================================================

# SmApply custom fieldss
# 153438 = "TM_ReviewNum" — 1-based index into REVIEWERS list
# 153439 =               — reviewer's name (read-only in this script)
# 153497 =               — reviewer's email (read-only in this script)
SMAPPLY_CUSTOM_FIELD_ID   = "153438"
SMAPPLY_REVIEWER_FIELD_ID = "153439"
SMAPPLY_EMAIL_FIELD_ID    = "153497"

# ---------------------------------------------------------
# REVIEWER ROSTER — must stay in sync with the original script
# Index position (1-based) must match what's stored in field 153438
# ---------------------------------------------------------
REVIEWERS = [
    {"id": 131224447, "name": "Taylor Gasser - SYS",    "email": "tgasser@trumerit.org"},   # 1
    {"id": 131514568, "name": "Kimberly Huang - SYS",   "email": "khuang@trumerit.org"},    # 2
    {"id": 131514573, "name": "Alyson O'Shea - SYS",    "email": "ao'shea@trumerit.org"},   # 3
    {"id": 131514574, "name": "Tatiana Lunina - SYS",   "email": "tlunina@trumerit.org"},   # 4
    {"id": 131514576, "name": "Ibeth Luck - SYS",       "email": "iluck@trumerit.org"},     # 5
    {"id": 131514578, "name": "Katherine Winder - SYS", "email": "kwinder@trumerit.org"},   # 6
    {"id": 131514579, "name": "Nathan Leonard - SYS",   "email": "nleonard@trumerit.org"},  # 7
    {"id": 131224463, "name": "Yamina Toumi - SYS",     "email": "ytoumi@trumerit.org"},    # 8
]

SMAPPLY_STAGE_IDS = [
    1362842, # Application Review
    1362861, # Mailroom Document Intake & Upload
    1362863, # Identity Document Review
    1362864, # Education Document Review
    1362865, # Registration Document Review
    1363446, # Customer Support View
    1362867, # Assessment Notes
    1367706, # Credential Report Validation - Finalize Application
]


# =========================================================
# SMAPPLY — READ
# =========================================================

def get_smapply_application(appid):
    """Fetch the full application object from SmApply."""
    access_token = wmill.get_variable("u/epham2/SMAPPLY_ACCESS_TOKEN")
    url = "https://nnas.smapply.io/api/applications/%s/" % appid
    response = requests.get(
        url,
        headers={"Authorization": "Bearer %s" % access_token},
    )
    response.raise_for_status()
    return response.json()


def get_reviewer_index(application):
    """
    Extract the 1-based reviewer index from custom field 153438.
    Returns an int, or raises if the field is missing or invalid.
    """
    for field in application.get("custom_fields", []):
        if str(field.get("id")) == SMAPPLY_CUSTOM_FIELD_ID:
            value = field.get("value")
            if value is None or str(value).strip() == "":
                raise ValueError(
                    "Custom field %s is empty — application may not have been "
                    "processed by the original assignment script yet." % SMAPPLY_CUSTOM_FIELD_ID
                )
            index = int(value)
            if index < 1 or index > len(REVIEWERS):
                raise ValueError(
                    "Reviewer index %d is out of range (expected 1–%d)."
                    % (index, len(REVIEWERS))
                )
            return index
    raise ValueError(
        "Custom field %s not found on application." % SMAPPLY_CUSTOM_FIELD_ID
    )


def reviewer_for_index(index):
    """Map a 1-based index to its reviewer dict ({'id', 'name', 'email'})."""
    return REVIEWERS[index - 1]


# =========================================================
# SMAPPLY — WRITE
# =========================================================

def _existing_assignment_pairs(appid, headers):
    """
    Return a set of (stage_id, reviewer_id) tuples already assigned for this
    application, so we can skip re-POSTing them.
    """
    url = "https://nnas.smapply.io/api/assignments/?application=%s" % appid
    r = requests.get(url, headers=headers)
    r.raise_for_status()
    pairs = set()
    for item in r.json().get("results", []):
        s = (item.get("stage") or {}).get("id")
        rv = (item.get("reviewer") or {}).get("id")
        if s is not None and rv is not None:
            pairs.add((int(s), int(rv)))
    return pairs


def create_review_assignments(appid, reviewer_id):
    """
    Create a review assignment for this application on every stage in
    SMAPPLY_STAGE_IDS, all assigned to the same reviewer.
    Idempotent: stages already assigned to this reviewer are skipped.
    """
    access_token = wmill.get_variable("u/epham2/SMAPPLY_ACCESS_TOKEN")
    url = "https://nnas.smapply.io/api/assignments/"
    headers = {
        "Authorization": "Bearer %s" % access_token,
        "Content-Type": "application/json",
    }

    already = _existing_assignment_pairs(appid, headers)

    for stage_id in SMAPPLY_STAGE_IDS:
        if (int(stage_id), int(reviewer_id)) in already:
            print("  Skip stage %d — reviewer %d already assigned" % (stage_id, reviewer_id))
            continue
        payload = {
            "applications": [int(appid)],
            "stage": stage_id,
            "reviewer": reviewer_id,
        }
        response = requests.post(url, json=payload, headers=headers)
        if not response.ok:
            print("  Assignment FAILED (stage %d, reviewer %d): %s %s"
                  % (stage_id, reviewer_id, response.status_code, response.text[:1000]))
        response.raise_for_status()
        print("  Assigned reviewer %d -> stage %d" % (reviewer_id, stage_id))


# =========================================================
# MAIN
# =========================================================
# Accepts the same webhook payload shape as the original script.
# Does NOT advance or modify the round-robin counter.

def main(REFERENCE_ID, USER_ID, USER, appid, ADDITIONAL_CONTENT):
    application = get_smapply_application(appid)
    index = get_reviewer_index(application)
    reviewer = reviewer_for_index(index)

    print("Reassigning appid %s to reviewer #%d: %s <%s> (id %d)"
          % (appid, index, reviewer["name"], reviewer["email"], reviewer["id"]))

    create_review_assignments(appid, reviewer["id"])

    print("Reassignment complete.")

    return {
        "appid": appid,
        "reviewer_index": index,
        "reviewer": reviewer["name"],
        "reviewer_email": reviewer["email"],
        "reviewer_id": reviewer["id"],
    }
    