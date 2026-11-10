#requirements:
#wmill
#requests

import wmill
import requests


# =========================================================
# CONFIGURATION
# =========================================================

# SmApply custom fields
# 153438 = "TM_ReviewNum"  — receives the rotating counter (stored as string)
# 153439 =                 — receives the assigned reviewer's name
# 153497 =                 — receives the assigned reviewer's email
SMAPPLY_CUSTOM_FIELD_ID   = "153438"
SMAPPLY_REVIEWER_FIELD_ID = "153439"
SMAPPLY_EMAIL_FIELD_ID    = "153497"

# ---------------------------------------------------------
# REVIEWER ROSTER  — edit this list to add/remove reviewers
# COUNTER_MAX tracks len(REVIEWERS) automatically.
# ---------------------------------------------------------
REVIEWERS = [
    {"id": 131224447, "name": "Taylor Gasser - SYS",    "email": "tgasser@trumerit.org"},   # 1
    {"id": 131514568, "name": "Kimberly Huang - SYS",   "email": "khuang@trumerit.org"},    # 2
    {"id": 131514573, "name": "Alyson O'Shea - SYS",    "email": "ao'shea@trumerit.org"},   # 3
    {"id": 131514574, "name": "Tatiana Lunina - SYS",   "email": "tlunina@trumerit.org"},   # 4
    {"id": 131514576, "name": "Ibeth Luck - SYS",       "email": "iluck@trumerit.org"},     # 5
    {"id": 131514578, "name": "Katherine Winder - SYS", "email": "kwinder@trumerit.org"},   # 6
    {"id": 131514579, "name": "Nathan Leonard - SYS",   "email": "nleonard@trumerit.org"},  # 7
]

COUNTER_MAX = len(REVIEWERS)

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
# PERSISTENT ROUND-ROBIN COUNTER  (Windmill state)
# =========================================================
# NOTE: state is per-script. This isolated script gets its OWN state slot,
# independent of licensure_form_tm.py — see the deployment notes below.

def next_counter():
    """Return the next value in the 1..COUNTER_MAX cycle. Does NOT persist."""
    last = wmill.get_state()
    last = int(last) if last is not None else 0
    nxt = last + 1
    if nxt > COUNTER_MAX:
        nxt = 1
    return nxt


def commit_counter(value):
    """Persist the value so the next run continues the cycle."""
    wmill.set_state(int(value))


def reviewer_for_counter(counter):
    """Map a 1-based counter value to its reviewer dict ({'id', 'name', 'email'})."""
    return REVIEWERS[counter - 1]


# =========================================================
# SMAPPLY
# =========================================================

def set_smapply_custom_field(appid, field_id, value):
    """Write a value to a SmApply custom field on an application."""
    access_token = wmill.get_variable("u/epham2/SMAPPLY_ACCESS_TOKEN")
    url = "https://nnas.smapply.io/api/applications/%s/" % appid
    response = requests.patch(
        url,
        json={"custom_fields": [{"id": int(field_id), "value": str(value)}]},
        headers={
            "Authorization": "Bearer %s" % access_token,
            "Content-Type": "application/json",
        },
    )
    response.raise_for_status()
    print("  Set SmApply custom field %s = %s" % (field_id, value))


def _existing_assignment_pairs(appid, headers):
    """
    Return a set of (stage_id, reviewer_id) tuples already assigned for this
    application, so we can skip re-POSTing them on a retry.
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
# Signature matches licensure_form_tm.py so this can be triggered with the
# SAME webhook payload. Only `appid` is used here; the other four parameters
# are accepted and ignored so you don't have to reshape the trigger body.

def main(REFERENCE_ID, USER_ID, USER, appid, ADDITIONAL_CONTENT):
    counter = next_counter()
    reviewer = reviewer_for_counter(counter)

    # Write fields first, then assignments, then commit the counter LAST.
    # Any failure above raises before commit, so a failed run never burns a
    # number — the next run retries the same counter value.
    set_smapply_custom_field(appid, SMAPPLY_CUSTOM_FIELD_ID, counter)
    set_smapply_custom_field(appid, SMAPPLY_REVIEWER_FIELD_ID, reviewer["name"])
    set_smapply_custom_field(appid, SMAPPLY_EMAIL_FIELD_ID, reviewer["email"])
    create_review_assignments(appid, reviewer["id"])

    commit_counter(counter)
    print("Counter committed: %d -> %s <%s> (id %d)"
          % (counter, reviewer["name"], reviewer["email"], reviewer["id"]))

    return {
        "appid": appid,
        "counter": counter,
        "reviewer": reviewer["name"],
        "reviewer_email": reviewer["email"],
        "reviewer_id": reviewer["id"],
    }