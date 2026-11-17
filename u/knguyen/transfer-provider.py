# py: 3.12
# requirements:
# wmill
# requests
# openpyxl

"""
SMApply cross-stream application clone
======================================

Creates a NEW application in the opposite expedited program and copies the
source application's task responses into it.

  CES "Expedited Application (C)"  program 295865
  TM  "Expedited Application (T)"  program 295185

The new application is owned by the SAME applicant as the source; it is created
by the API token's admin account, so SMApply will show it as admin-created.

Direction is parameterized: pass target_program="tm" | "ces" | "auto"
("auto" = the program the source is NOT in).

Safety
------
dry_run=True (the default) performs NO writes. It resolves the eligibility
check, the full task and field crosswalk, builds every payload, and returns
exactly what it would do. Run it dry first and read `warnings`.

Eligibility workflow
--------------------
SMApply refuses to create an application for an applicant who does not meet
the target program's eligibility criteria, even through the admin API. Before
creating anything, the script reads the applicant's eligibility answers and
checks them against the target program's criteria (PROGRAM_CRITERIA).

  * Already eligible: the transfer proceeds.
  * Not eligible, no eligibility_changes passed: the run stops and reports
    which condition fails. Nothing is changed.
  * Not eligible, eligibility_changes passed: only with approved_by and
    applicant_confirmation filled in, and only if the changes would make the
    applicant eligible. The script then
      1. appends the planned change to the change log workbook in SharePoint
         (the run stops here if the log cannot be written),
      2. updates the eligibility answers and re-reads them to verify that only
         the intended answers changed (restoring the originals if not),
      3. creates and populates the new application,
      4. logs the outcome; if the create fails, the original answers are
         restored and logged (restore_eligibility_on_failure).

The change log is ONE workbook with ONE Excel table, appended to row by row
through the Microsoft Graph workbook API, so concurrent runs cannot overwrite
each other and SharePoint version history keeps every earlier state. Each
changed question is one row per event, so the log is append-only.

Eligibility updates use PATCH /api/users/{id}/eligibility/ as documented:
{"is_completed": <current value>, "responses": {question id: value}}. The
token's admin account needs the "Assume role" permission. The docs do not say
whether choice answers are written as option indices (as they are read) or
as labels, so live changes stay disabled until a test run on a TEST applicant
(mode="test_eligibility_update") passes and
ELIGIBILITY_UPDATE_FORMAT_CONFIRMED is set to True.

What is NOT copied
------------------
  * payment tasks, sign-off, mailroom/credential-report/staff tasks
    (see EXCLUDE_SUBSTRINGS)
  * fields with no counterpart in the target program (flagged, not silent)
  * ambiguous repeat-block fields (see below)

Ambiguity
---------
The task-definition `fields` object is ASCII-sorted, so position does not
identify which repeat block a field belongs to. Where a label maps to several
keys on BOTH sides and the keys differ between programs, the pairing cannot be
established from the API. Default policy is to skip those and report them for
manual entry. Set ambiguous_field_policy="positional" to pair them by sorted
order instead. This WILL scramble block order if the sort happens not to
agree with form order, so only use it when the fields are empty anyway.

In practice this affects one group: the seven "What language was the
examination given in?" fields in Registration/Licensing History.

Exam blocks
-----------
CES "Other Nursing Licensing Examinations" carries 4 repeat blocks; TM carries
1. Going CES -> TM, only the block whose keys exist on both sides is copied and
the remainder is reported under warnings.
"""

import io
import json
import os
import re
import time
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import quote

import requests

try:
    import wmill
except ImportError:
    wmill = None


# --------------------------------------------------------------------------
# config
# --------------------------------------------------------------------------

TOKEN_PATH = "u/epham2/SMAPPLY_ACCESS_TOKEN"
BASE = "https://nnas.smapply.io"

PROGRAMS = {"ces": 295865, "tm": 295185}
PROGRAM_BY_ID = {v: k for k, v in PROGRAMS.items()}

REQUEST_TIMEOUT = 60
UPLOAD_TIMEOUT = 300
SLEEP_BETWEEN = 0.15
MAX_RETRIES = 3

TASK_STATUS_COMPLETE = 3

# Form-field names tried, in order, to tell the file upload endpoint which
# question an attachment belongs to.
FILE_QUESTION_FIELD_NAMES = ["question", "question_id", "key", "field"]

# Task names never copied. Matched as lowercase substrings.
EXCLUDE_SUBSTRINGS = [
    "payment",
    "mailroomstorage",
    "credential report",
    "credential form",
    "ien forms",
    "authenticated credentials",
    "application sign off",
    "contact us",
    "application instructions",
    "document translation",
    "terms and conditions - payment",
]

# Task names that differ between programs but mean the same thing.
# Normalized (lowercased, stripped) source name -> normalized target name.
TASK_NAME_ALIASES = {
    "application - contact us ces": "application - contact us tm",
    "application - contact us tm": "application - contact us ces",
}

# ---- eligibility ---------------------------------------------------------

# Eligibility choice questions, matched by label prefix, with their option
# order as confirmed from the eligibility form.
ELIGIBILITY_QUESTIONS: Dict[str, Tuple[str, List[str]]] = {
    "eligibility_to_apply": ("Eligibility to Apply to NNAS", ["Agree", "Disagree"]),
    "prior_report": ("Have you completed and receiv", ["Yes", "No"]),
    "ofac": ("Are any of the documents, cred", ["Yes", "No"]),
    "jurisdiction": ("Is any part of your application c", ["Yes", "No", "Unsure"]),
}

# Questions that record a sanctions connection. Changing these is always
# flagged in the report and the log.
SANCTIONS_KEYS = {"ofac", "jurisdiction"}

# Program eligibility criteria, transcribed from Program Settings > Eligibility.
# Update these if the settings change in SMApply.
PROGRAM_CRITERIA: Dict[str, Dict[str, Any]] = {
    "tm": {"match": "any", "conditions": [
        {"key": "jurisdiction", "equals": "No"},
        {"match": "all", "conditions": [
            {"key": "prior_report", "equals": "Yes"},
            {"key": "ofac", "equals": "No"},
        ]},
    ]},
    # As configured: ALL of Yes and No on the same question, which no
    # applicant can meet. CES is closed to new applications.
    "ces": {"match": "all", "conditions": [
        {"key": "jurisdiction", "equals": "Yes"},
        {"key": "jurisdiction", "equals": "No"},
    ]},
}

# Eligibility update endpoint, per the API documentation:
#   PATCH /api/users/{id}/eligibility/
#   {"is_completed": true, "responses": {"<question id>": <value>}}
# Requires the "Assume role" permission on the token's admin account.
# Choice answers are sent as option indices, the same form GET returns. The
# docs' example shows text values, so confirm with a test run first, then set
# ELIGIBILITY_UPDATE_FORMAT_CONFIRMED = True.
ELIGIBILITY_UPDATE_METHOD = "PATCH"
ELIGIBILITY_UPDATE_PATH = "/api/users/{user_id}/eligibility/"
ELIGIBILITY_UPDATE_BODY_KEY = "responses"
ELIGIBILITY_UPDATE_FORMAT_CONFIRMED = True   # confirmed by test run 2026-11-17

# ---- change log (SharePoint) ---------------------------------------------

TENANT_ID_VAR = "u/epham2/AZURE_TENANT_ID"
CLIENT_ID_VAR = "u/epham2/SHAREPOINT_CLIENT_ID"
CLIENT_SECRET_VAR = "u/epham2/SHAREPOINT_CLIENT_SECRET"
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SHAREPOINT_HOSTNAME = "nnasca.sharepoint.com"
SITE_PATH = "/sites/NNAS-Portal-Sharepoint"
DOCUMENT_LIBRARY_NAME = "Applications"

# Inside the Applications library. Restrict access to this folder in
# SharePoint: it records changes to applicants' sanctions answers.
CHANGE_LOG_FOLDER = "Eligibility Change Records"
CHANGE_LOG_FILE = "Eligibility Change Log.xlsx"
CHANGE_LOG_TABLE = "ChangeLog"
CHANGE_LOG_COLUMNS = [
    "Logged at (UTC)", "Run ID", "Event", "Application reference",
    "Application ID", "Applicant user ID", "Question ID", "Question",
    "Stored value (before)", "Answer (before)", "Stored value (after)",
    "Answer (after)", "Sanctions question", "Reason code", "Approved by",
    "Applicant confirmation", "New application ID", "Notes",
]


# --------------------------------------------------------------------------
# http
# --------------------------------------------------------------------------

def _get_token() -> str:
    if wmill is not None:
        token = wmill.get_variable(TOKEN_PATH)
        if not token:
            raise RuntimeError(f"Windmill variable {TOKEN_PATH} resolved empty.")
        return token
    tok = os.environ.get("SMAPPLY_ACCESS_TOKEN")
    if not tok:
        raise RuntimeError("wmill unavailable and SMAPPLY_ACCESS_TOKEN not set.")
    return tok


class Api:
    def __init__(self, token: str, base: str = BASE):
        self.token = token
        self.base = base.rstrip("/")
        self.calls: List[Dict[str, Any]] = []

    def _headers(self, json_body: bool = False, automations: bool = False) -> Dict[str, str]:
        h = {"Authorization": f"Bearer {self.token}", "Accept": "application/json"}
        if json_body:
            h["Content-Type"] = "application/json"
        if automations:
            h["apply-automations"] = "full"
        return h

    def request(self, method: str, path: str, *, params=None, json_body=None,
                files=None, data=None, automations=False,
                timeout=REQUEST_TIMEOUT) -> Dict[str, Any]:
        url = path if path.startswith("http") else f"{self.base}{path}"
        headers = self._headers(json_body=json_body is not None, automations=automations)
        last = None
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                resp = requests.request(
                    method, url, headers=headers, params=params, json=json_body,
                    files=files, data=data, timeout=timeout,
                )
                time.sleep(SLEEP_BETWEEN)
                if resp.status_code == 429 and attempt < MAX_RETRIES:
                    time.sleep(2 * attempt)
                    continue
                try:
                    body = resp.json()
                except Exception:
                    # 500s return HTML, not JSON
                    body = {"_non_json": True, "text": (resp.text or "")[:1500]}
                rec = {
                    "method": method, "url": url, "status": resp.status_code,
                    "ok": 200 <= resp.status_code < 300, "body": body,
                }
                self.calls.append({k: rec[k] for k in ("method", "url", "status", "ok")})
                return rec
            except Exception as exc:
                last = str(exc)
                time.sleep(1.5 * attempt)
        rec = {"method": method, "url": url, "status": None, "ok": False, "error": last}
        self.calls.append(rec)
        return rec

    def get(self, path, **kw):
        return self.request("GET", path, **kw)

    def post(self, path, **kw):
        return self.request("POST", path, **kw)

    def patch(self, path, **kw):
        return self.request("PATCH", path, **kw)


# --------------------------------------------------------------------------
# label / key helpers
# --------------------------------------------------------------------------

_SUFFIX_RE = re.compile(r"_(\d+)$")


def clean_label(s: Any) -> str:
    """SMApply field names carry the help text after a newline; keep line one."""
    first = str(s or "").split("\n")[0]
    return re.sub(r"\s+", " ", first).strip().lower()


def split_suffix(key: str) -> Tuple[str, str]:
    m = _SUFFIX_RE.search(key or "")
    if m:
        return key[: m.start()], m.group(0)
    return key, ""


def norm_task_name(name: Any) -> str:
    return re.sub(r"\s+", " ", str(name or "")).strip().lower()


def is_excluded(name: Any) -> bool:
    n = norm_task_name(name)
    return any(sub in n for sub in EXCLUDE_SUBSTRINGS)


# --------------------------------------------------------------------------
# fetching
# --------------------------------------------------------------------------

def fetch_application(api: Api, app_id: int) -> Dict[str, Any]:
    res = api.get(f"/api/applications/{app_id}/")
    if not res["ok"]:
        raise RuntimeError(f"Could not read application {app_id}: "
                           f"HTTP {res['status']} {res.get('body')}")
    return res["body"]


def fetch_task_ids(api: Api, app_id: int) -> List[int]:
    res = api.get(f"/api/applications/{app_id}/tasks/")
    if not res["ok"]:
        raise RuntimeError(f"Could not list tasks for {app_id}: HTTP {res['status']}")
    body = res["body"]
    rows = body.get("results") if isinstance(body, dict) else body
    if not isinstance(rows, list):
        rows = []
    ids = []
    for r in rows:
        if isinstance(r, dict) and r.get("id") is not None:
            ids.append(r["id"])
    return ids


def fetch_task_bundle(api: Api, app_id: int, task_id: int) -> Dict[str, Any]:
    """Instance (has data) plus definition (has field names, types, choices)."""
    inst = api.get(f"/api/applications/{app_id}/tasks/{task_id}/")
    defn = api.get(f"/api/tasks/{task_id}/")
    return {
        "task_id": task_id,
        "instance": inst["body"] if inst["ok"] else None,
        "definition": defn["body"] if defn["ok"] else None,
        "instance_status": inst["status"],
        "definition_status": defn["status"],
    }


def fetch_all_tasks(api: Api, app_id: int) -> Dict[str, Dict[str, Any]]:
    """Keyed by normalized task name."""
    out: Dict[str, Dict[str, Any]] = {}
    for tid in fetch_task_ids(api, app_id):
        b = fetch_task_bundle(api, app_id, tid)
        defn = b["definition"] or {}
        name = defn.get("name") or (b["instance"] or {}).get("name") or str(tid)
        out[norm_task_name(name)] = {**b, "name": name,
                                     "fields": defn.get("fields") or {},
                                     "type": defn.get("type")}
    return out


def pick_reference_application(api: Api, program_id: int,
                               exclude_id: Optional[int] = None) -> Optional[int]:
    """Any application in the target program, used to read its task shape."""
    res = api.get(f"/api/programs/{program_id}/applications/", params={"per_page": 20})
    if not res["ok"]:
        return None
    for row in (res["body"] or {}).get("results", []):
        if isinstance(row, dict) and row.get("id") and row["id"] != exclude_id:
            return row["id"]
    return None


# --------------------------------------------------------------------------
# field mapping
# --------------------------------------------------------------------------

def build_field_map(src_fields: Dict[str, Any], tgt_fields: Dict[str, Any],
                    policy: str) -> Dict[str, Any]:
    """
    Resolve source field key -> target field key.

    1. Identity: same key on both sides with the same label. Covers the large
       majority, because the two programs were cloned from one another.
    2. Unique label: exactly one remaining key per (label, type) on each side.
    3. Ambiguous: several remaining keys share a (label, type) on both sides.
       Position is not reliable here, so the default is to skip.
    4. Unmappable: no target counterpart at all.
    """
    mapping: Dict[str, str] = {}
    ambiguous: List[Dict[str, Any]] = []
    unmappable: List[Dict[str, Any]] = []

    src_rest = dict(src_fields)
    tgt_rest = dict(tgt_fields)

    # --- 1. identity
    for k in list(src_rest):
        if k in tgt_rest and clean_label(src_rest[k].get("name")) == clean_label(
            tgt_rest[k].get("name")
        ):
            mapping[k] = k
            src_rest.pop(k)
            tgt_rest.pop(k)

    def group(fields: Dict[str, Any]) -> Dict[Tuple[str, Any], List[str]]:
        g: Dict[Tuple[str, Any], List[str]] = {}
        for k, v in fields.items():
            g.setdefault((clean_label(v.get("name")), v.get("type")), []).append(k)
        for ks in g.values():
            ks.sort()
        return g

    gs, gt = group(src_rest), group(tgt_rest)

    for sig, sks in gs.items():
        tks = gt.get(sig)
        label, ftype = sig
        if not tks:
            for k in sks:
                unmappable.append({"source_key": k, "label": label, "type": ftype})
            continue
        # --- 2. unique on both sides
        if len(sks) == 1 and len(tks) == 1:
            mapping[sks[0]] = tks[0]
            continue
        # --- 3. ambiguous
        if policy == "positional":
            for i, sk in enumerate(sks):
                if i < len(tks):
                    mapping[sk] = tks[i]
                else:
                    unmappable.append({"source_key": sk, "label": label, "type": ftype})
            ambiguous.append({
                "label": label, "type": ftype,
                "source_keys": sks, "target_keys": tks,
                "resolution": "paired positionally (sort order, UNVERIFIED)",
            })
        else:
            ambiguous.append({
                "label": label, "type": ftype,
                "source_keys": sks, "target_keys": tks,
                "resolution": "skipped; pairing not derivable from the API",
            })

    return {"map": mapping, "ambiguous": ambiguous, "unmappable": unmappable}


def resolve_key(data_key: str, fmap: Dict[str, str]) -> Optional[str]:
    """Map a data key, preserving any _N repeat suffix."""
    if data_key in fmap:
        return fmap[data_key]
    base, suf = split_suffix(data_key)
    if base in fmap:
        return fmap[base] + suf
    if suf and (base + "_0") in fmap:
        tgt_base, _ = split_suffix(fmap[base + "_0"])
        return tgt_base + suf
    return None


def is_file_value(v: Any) -> bool:
    return isinstance(v, list) and v and isinstance(v[0], dict) and "filename" in v[0]


def extract_response(v: Any) -> Any:
    """GET returns {'label','response'}; PATCH wants the bare value."""
    if isinstance(v, dict) and "response" in v:
        return v["response"]
    return v


def is_empty(v: Any) -> bool:
    if v is None:
        return True
    if isinstance(v, str) and not v.strip():
        return True
    if isinstance(v, (list, dict)) and len(v) == 0:
        return True
    return False


# SMApply stores an unanswered single-select as -1 rather than null.
CHOICE_UNANSWERED = -1


def is_unanswered_choice(value: Any) -> bool:
    if isinstance(value, bool):
        return False
    if isinstance(value, int):
        return value == CHOICE_UNANSWERED
    if isinstance(value, list) and value:
        return all(v == CHOICE_UNANSWERED for v in value)
    return False


def out_of_range_choices(value: Any, choices: Any) -> List[int]:
    """
    Choice responses are stored as indices. The two programs share a common
    prefix but the target may have FEWER options than the source (TM carries
    entries CES does not), so an index valid on one side can be invalid on the
    other. Returns the offending indices. -1 means unanswered and is not an
    error.
    """
    if not isinstance(choices, list):
        return []
    n = len(choices)
    idxs: List[int] = []
    if isinstance(value, bool):
        return []
    if isinstance(value, int):
        idxs = [value]
    elif isinstance(value, list):
        idxs = [v for v in value if isinstance(v, int) and not isinstance(v, bool)]
    return [i for i in idxs
            if i != CHOICE_UNANSWERED and (i < 0 or i >= n)]


def choice_labels(field: Dict[str, Any]) -> List[str]:
    ch = (field or {}).get("choices")
    if not isinstance(ch, list):
        return []
    return [str(c.get("label")) for c in ch if isinstance(c, dict)]


def build_task_payload(src_task: Dict[str, Any], tgt_task: Dict[str, Any],
                       policy: str, include_empty: bool,
                       copy_signature_fields: bool = False) -> Dict[str, Any]:
    src_data = ((src_task.get("instance") or {}).get("data") or {})
    src_fields = src_task.get("fields") or {}
    tgt_fields = tgt_task.get("fields") or {}
    fm = build_field_map(src_fields, tgt_fields, policy)

    payload: Dict[str, Any] = {}
    files: List[Dict[str, Any]] = []
    skipped: List[Dict[str, Any]] = []
    invalid_choices: List[Dict[str, Any]] = []
    signature_fields: List[Dict[str, Any]] = []
    target_defs: Dict[str, Any] = {}
    written = 0

    def target_field(tkey: str) -> Dict[str, Any]:
        base, _ = split_suffix(tkey)
        return tgt_fields.get(tkey) or tgt_fields.get(base) or {}

    for skey, sval in src_data.items():
        tkey = resolve_key(skey, fm["map"])

        # --- attachment arrays
        if is_file_value(sval):
            if tkey is None:
                skipped.append({"key": skey, "reason": "file field has no target key"})
                continue
            for f in sval:
                files.append({"source_key": skey, "target_key": tkey, "file": f})
            continue

        value = extract_response(sval)
        if tkey is None:
            if not is_empty(value):
                skipped.append({"key": skey, "reason": "no target key", "has_value": True})
            continue
        if (is_empty(value) or is_unanswered_choice(value)) and not include_empty:
            continue

        tf = target_field(tkey)

        # --- file-typed fields holding a media path (e.g. the T&C signature).
        # Writing the path as text would point the new application at the old
        # application's media, so it is copied as a binary or skipped.
        if tf.get("type") == "file":
            label = clean_label(tf.get("name"))
            if isinstance(value, str) and value.strip():
                if copy_signature_fields:
                    files.append({
                        "source_key": skey, "target_key": tkey,
                        "file": {"url": value,
                                 "filename": os.path.basename(value.split("?")[0]) or "upload"},
                        "from_path_string": True,
                    })
                else:
                    signature_fields.append({"key": skey, "target_key": tkey,
                                             "label": label})
            continue

        # --- choice index validity against the TARGET option list
        bad = out_of_range_choices(value, tf.get("choices"))
        if bad:
            labels = choice_labels(src_fields.get(split_suffix(skey)[0], {}))
            invalid_choices.append({
                "key": skey, "target_key": tkey,
                "label": clean_label(tf.get("name"))[:70],
                "indices": bad,
                "source_choice": [labels[i] for i in bad if 0 <= i < len(labels)],
                "target_option_count": len(tf.get("choices") or []),
            })
            continue

        payload[tkey] = value
        target_defs[tkey] = tf
        written += 1

    return {
        "source_task_id": src_task["task_id"],
        "target_task_id": tgt_task["task_id"],
        "name": src_task.get("name"),
        "payload": payload,
        "files": files,
        "fields_written": written,
        "fields_skipped": skipped,
        "invalid_choices": invalid_choices,
        "signature_fields": signature_fields,
        "ambiguous": fm["ambiguous"],
        "unmappable": fm["unmappable"],
        "target_fields": target_defs,
    }


# --------------------------------------------------------------------------
# eligibility
# --------------------------------------------------------------------------

def _norm(s: Any) -> str:
    return re.sub(r"\s+", " ", str(s or "")).strip().casefold()


def _is_index(v: Any) -> bool:
    return isinstance(v, int) and not isinstance(v, bool)


def read_eligibility(api: Api, user_id: int) -> Dict[str, Any]:
    res = api.get(f"/api/users/{user_id}/eligibility/")
    if not res["ok"]:
        raise RuntimeError(f"Could not read eligibility for user {user_id}: "
                           f"HTTP {res['status']}. The token's admin account "
                           f"needs the View Users permission.")
    return res["body"] if isinstance(res["body"], dict) else {}


def eligibility_questions(body: Dict[str, Any]) -> List[Dict[str, Any]]:
    """{"responses": {qid: {"label", "response"}}} -> list with question keys."""
    out: List[Dict[str, Any]] = []
    responses = body.get("responses") if isinstance(body, dict) else None
    if not isinstance(responses, dict):
        return out
    for qid, item in responses.items():
        if not isinstance(item, dict):
            continue
        label = item.get("label") or item.get("name") or ""
        key = next((k for k, (prefix, _) in ELIGIBILITY_QUESTIONS.items()
                    if _norm(label).startswith(_norm(prefix))), None)
        out.append({"id": qid, "label": re.sub(r"\s+", " ", label).strip(),
                    "key": key, "response": item.get("response")})
    return out


def answer_label(key: Optional[str], value: Any) -> str:
    """Readable answer for a choice question; never returns free text."""
    choices = ELIGIBILITY_QUESTIONS[key][1] if key else None
    if _is_index(value):
        if value == -1:
            return "(unanswered)"
        if choices and 0 <= value < len(choices):
            return choices[value]
        return f"option {value}"
    if value is None or value == "":
        return "(blank)"
    return "(text or file answer)"


def eligibility_answers(questions: List[Dict[str, Any]]) -> Dict[str, str]:
    answers = {k: "(not asked)" for k in ELIGIBILITY_QUESTIONS}
    for q in questions:
        if q["key"]:
            answers[q["key"]] = answer_label(q["key"], q["response"])
    return answers


def evaluate_criteria(node: Dict[str, Any], answers: Dict[str, str],
                      failures: List[str]) -> bool:
    if "conditions" in node:
        res = [evaluate_criteria(c, answers, failures) for c in node["conditions"]]
        return any(res) if node.get("match") == "any" else all(res)
    actual = answers.get(node["key"], "(not asked)")
    ok = actual == node["equals"]
    if not ok:
        failures.append(f"{node['key']} is {actual}, needs {node['equals']}")
    return ok


def resolve_eligibility_changes(questions: List[Dict[str, Any]],
                                changes: Dict[str, str]) -> Dict[str, Dict[str, Any]]:
    """{question key: new answer} -> {qid: change details}. Raises ValueError."""
    resolved: Dict[str, Dict[str, Any]] = {}
    for key, new_answer in changes.items():
        if key not in ELIGIBILITY_QUESTIONS:
            raise ValueError(f"Unknown eligibility question '{key}'. "
                             f"Use one of {list(ELIGIBILITY_QUESTIONS)}.")
        choices = ELIGIBILITY_QUESTIONS[key][1]
        if new_answer not in choices:
            raise ValueError(f"'{new_answer}' is not an option for {key}: {choices}.")
        matches = [q for q in questions if q["key"] == key]
        if len(matches) != 1:
            raise ValueError(f"Expected one '{key}' question on the applicant's "
                             f"eligibility form, found {len(matches)}.")
        q = matches[0]
        new_value = choices.index(new_answer)
        if q["response"] == new_value:
            continue
        resolved[q["id"]] = {
            "key": key, "question": q["label"][:120],
            "before_value": q["response"], "before_answer": answer_label(key, q["response"]),
            "after_value": new_value, "after_answer": new_answer,
            "sanctions": key in SANCTIONS_KEYS,
        }
    return resolved


def write_eligibility(api: Api, user_id: int, values: Dict[str, Any],
                      is_completed: Optional[bool]) -> Dict[str, Any]:
    """Send only the answers being changed, keeping is_completed as it was."""
    path = ELIGIBILITY_UPDATE_PATH.format(user_id=user_id)
    body: Dict[str, Any] = {ELIGIBILITY_UPDATE_BODY_KEY: values}
    if isinstance(is_completed, bool):
        body["is_completed"] = is_completed
    res = api.request(ELIGIBILITY_UPDATE_METHOD, path, json_body=body)
    return {"status": res["status"], "ok": res["ok"],
            "body": None if res["ok"] else res.get("body")}


def verify_eligibility(api: Api, user_id: int, expected: Dict[str, Any],
                       untouched: Dict[str, Any],
                       is_completed: Optional[bool] = None) -> List[str]:
    """Re-read; return mismatches. Empty list = exactly as intended."""
    body = read_eligibility(api, user_id)
    questions = eligibility_questions(body)
    current = {q["id"]: q["response"] for q in questions}
    problems = []
    if isinstance(is_completed, bool) and body.get("is_completed") != is_completed:
        problems.append(f"is_completed changed to {body.get('is_completed')!r}")
    for qid, value in expected.items():
        if current.get(qid) != value:
            problems.append(f"{qid}: expected {value}, found {current.get(qid)!r}")
    for qid, value in untouched.items():
        if current.get(qid) != value:
            problems.append(f"{qid}: should be unchanged but differs "
                            f"({'missing' if qid not in current else 'changed'})")
    return problems


# --------------------------------------------------------------------------
# change log (one workbook, one table, rows appended via Graph)
# --------------------------------------------------------------------------

class ChangeLog:
    def __init__(self, folder: str = CHANGE_LOG_FOLDER, filename: str = CHANGE_LOG_FILE,
                 table: str = CHANGE_LOG_TABLE):
        self.folder, self.filename, self.table = folder, filename, table
        self.token: Optional[str] = None
        self.drive_id: Optional[str] = None
        self.item_id: Optional[str] = None
        self.web_url = ""

    # -- graph plumbing
    def _token(self) -> str:
        if self.token:
            return self.token
        tenant = (wmill.get_variable(TENANT_ID_VAR) or "").strip()
        client_id = (wmill.get_variable(CLIENT_ID_VAR) or "").strip()
        secret = (wmill.get_variable(CLIENT_SECRET_VAR) or "").strip()
        r = requests.post(
            f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token",
            data={"client_id": client_id, "client_secret": secret,
                  "scope": "https://graph.microsoft.com/.default",
                  "grant_type": "client_credentials"},
            timeout=60)
        if r.status_code != 200:
            raise RuntimeError(f"Graph token request failed {r.status_code}: {r.text[:500]}")
        self.token = r.json()["access_token"]
        return self.token

    def _req(self, method: str, endpoint: str, ok=(200,), **kw) -> Any:
        headers = kw.pop("headers", {})
        headers["Authorization"] = f"Bearer {self._token()}"
        r = requests.request(method, f"{GRAPH_BASE}{endpoint}",
                             headers=headers, timeout=120, **kw)
        if r.status_code not in ok:
            raise RuntimeError(f"Graph {method} {endpoint} failed {r.status_code}: "
                               f"{r.text[:500]}")
        return r.json() if r.content else None

    def _get_item(self, path: str) -> Optional[Dict[str, Any]]:
        r = requests.get(f"{GRAPH_BASE}/drives/{self.drive_id}/root:/{quote(path)}",
                         headers={"Authorization": f"Bearer {self._token()}"}, timeout=60)
        if r.status_code == 404:
            return None
        if r.status_code != 200:
            raise RuntimeError(f"Graph lookup failed {r.status_code}: {r.text[:500]}")
        return r.json()

    def _init(self) -> None:
        if self.item_id:
            return
        site = self._req("GET", f"/sites/{SHAREPOINT_HOSTNAME}:{SITE_PATH}")
        drives = self._req("GET", f"/sites/{site['id']}/drives")["value"]
        drive = next((d for d in drives if d.get("name") == DOCUMENT_LIBRARY_NAME), None)
        if not drive:
            raise RuntimeError(f"Document library '{DOCUMENT_LIBRARY_NAME}' not found.")
        self.drive_id = drive["id"]

        if self._get_item(self.folder) is None:
            self._req("POST", f"/drives/{self.drive_id}/root/children", ok=(201, 409),
                      headers={"Content-Type": "application/json"},
                      json={"name": self.folder, "folder": {},
                            "@microsoft.graph.conflictBehavior": "fail"})

        path = f"{self.folder}/{self.filename}"
        item = self._get_item(path)
        if item is None:
            r = requests.put(
                f"{GRAPH_BASE}/drives/{self.drive_id}/root:/{quote(path)}:/content"
                f"?@microsoft.graph.conflictBehavior=fail",
                headers={"Authorization": f"Bearer {self._token()}",
                         "Content-Type": "application/vnd.openxmlformats-"
                                         "officedocument.spreadsheetml.sheet"},
                data=self._new_workbook(), timeout=120)
            if r.status_code in (200, 201):
                item = r.json()
            elif r.status_code == 409:      # created by a concurrent run
                item = self._get_item(path)
            else:
                raise RuntimeError(f"Could not create change log {r.status_code}: "
                                   f"{r.text[:500]}")
        self.item_id = item["id"]
        self.web_url = item.get("webUrl", "")

    def _new_workbook(self) -> bytes:
        from openpyxl import Workbook
        from openpyxl.styles import Font
        from openpyxl.utils import get_column_letter
        from openpyxl.worksheet.table import Table, TableStyleInfo

        wb = Workbook()
        ws = wb.active
        ws.title = "Change Log"
        ws.append(CHANGE_LOG_COLUMNS)
        first = [""] * len(CHANGE_LOG_COLUMNS)
        first[0] = _utc_now()
        first[2] = "log_created"
        first[-1] = ("Append-only. Rows are added by the cross-stream clone "
                     "script; do not edit or delete rows.")
        ws.append(first)
        for row in ws.iter_rows():
            for cell in row:
                cell.font = Font(name="Arial", size=10, bold=cell.row == 1)
        last_col = get_column_letter(len(CHANGE_LOG_COLUMNS))
        table = Table(displayName=self.table, ref=f"A1:{last_col}2")
        table.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
        ws.add_table(table)
        for i, name in enumerate(CHANGE_LOG_COLUMNS, start=1):
            ws.column_dimensions[get_column_letter(i)].width = max(14, min(len(name) + 4, 40))
        ws.freeze_panes = "A2"
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    # -- public
    def append(self, rows: List[List[Any]]) -> None:
        if not rows:
            return
        self._init()
        values = [["" if v is None else v for v in r] for r in rows]
        self._req("POST",
                  f"/drives/{self.drive_id}/items/{self.item_id}/workbook/tables/"
                  f"{self.table}/rows/add",
                  ok=(200, 201), headers={"Content-Type": "application/json"},
                  json={"index": None, "values": values})


def _utc_now() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")


def log_rows(event: str, ctx: Dict[str, Any],
             changes: Optional[Dict[str, Dict[str, Any]]] = None,
             use_after: bool = True, new_app_id: Any = "", notes: str = "") -> List[List[Any]]:
    """One row per changed question, or a single row when there are none."""
    def row(qid="", c=None):
        c = c or {}
        return [
            _utc_now(), ctx["run_id"], event, ctx["reference"], ctx["application_id"],
            ctx["applicant_id"], qid, c.get("question", ""),
            c.get("before_value", ""), c.get("before_answer", ""),
            c.get("after_value", "") if use_after else c.get("before_value", ""),
            c.get("after_answer", "") if use_after else c.get("before_answer", ""),
            ("Yes" if c.get("sanctions") else "No") if c else "",
            ctx["reason_code"], ctx["approved_by"], ctx["applicant_confirmation"],
            new_app_id or "", notes,
        ]
    if not changes:
        return [row()]
    return [row(qid, c) for qid, c in changes.items()]


# --------------------------------------------------------------------------
# eligibility gate (runs before any create)
# --------------------------------------------------------------------------

def eligibility_gate(api: Api, target_stream: str, applicant_id: int,
                     ctx: Dict[str, Any], changes: Optional[Dict[str, str]],
                     dry_run: bool, log: ChangeLog) -> Dict[str, Any]:
    """
    Returns {"status": "eligible" | "planned" | "updated" | "blocked" | "error",
             "report": {...}, "message": str, "applied": {...} or None}
    """
    criteria = PROGRAM_CRITERIA[target_stream]
    elig_body = read_eligibility(api, applicant_id)
    completed = elig_body.get("is_completed")
    questions = eligibility_questions(elig_body)
    answers = eligibility_answers(questions)
    failures: List[str] = []
    eligible = evaluate_criteria(criteria, answers, failures)
    rep: Dict[str, Any] = {"answers": answers, "meets_target_criteria": eligible}

    if eligible:
        if changes:
            rep["note"] = "Already eligible; eligibility_changes were not needed or applied."
        return {"status": "eligible", "report": rep, "applied": None}

    rep["failing_conditions"] = failures
    if not changes:
        return {"status": "blocked", "report": rep, "applied": None,
                "message": (f"The applicant does not meet the {target_stream.upper()} "
                            f"eligibility criteria, so SMApply would refuse the "
                            f"create. Nothing was changed. If the applicant has "
                            f"confirmed different answers, pass eligibility_changes "
                            f"with approved_by and applicant_confirmation.")}
    if not ctx["approved_by"] or not ctx["applicant_confirmation"]:
        return {"status": "error", "report": rep, "applied": None,
                "message": "eligibility_changes requires approved_by and "
                           "applicant_confirmation."}
    try:
        resolved = resolve_eligibility_changes(questions, changes)
    except ValueError as exc:
        return {"status": "error", "report": rep, "applied": None, "message": str(exc)}

    after = dict(answers)
    for c in resolved.values():
        after[c["key"]] = c["after_answer"]
    after_failures: List[str] = []
    eligible_after = evaluate_criteria(criteria, after, after_failures)
    rep["planned_changes"] = [
        {"question_id": qid, "question": c["question"], "before": c["before_answer"],
         "after": c["after_answer"], "sanctions_question": c["sanctions"]}
        for qid, c in resolved.items()]
    rep["meets_target_criteria_after_changes"] = eligible_after
    if any(c["sanctions"] for c in resolved.values()):
        rep["sanctions_answer_change"] = (
            "This changes the applicant's sanctions answer. It is recorded with "
            "the approver and the applicant's confirmation.")
    if not resolved:
        return {"status": "error", "report": rep, "applied": None,
                "message": "eligibility_changes match the current answers; nothing to change."}
    if not eligible_after:
        rep["failing_conditions_after_changes"] = after_failures
        return {"status": "error", "report": rep, "applied": None,
                "message": (f"These changes would still not meet the "
                            f"{target_stream.upper()} criteria. Nothing was changed.")}

    if dry_run:
        rep["log_rows_to_append"] = len(resolved)
        return {"status": "planned", "report": rep, "applied": None}

    if not ELIGIBILITY_UPDATE_FORMAT_CONFIRMED:
        return {"status": "error", "report": rep, "applied": None,
                "message": ("The eligibility update format has not been confirmed. "
                            "Run mode='test_eligibility_update' on a TEST applicant, "
                            "then set ELIGIBILITY_UPDATE_FORMAT_CONFIRMED = True. "
                            "Nothing was changed.")}

    # 1. log first; no log, no change
    try:
        log.append(log_rows("planned", ctx, resolved))
    except Exception as exc:
        return {"status": "error", "report": rep, "applied": None,
                "message": f"Could not write the change log, so nothing was changed: {exc}"}
    rep["change_log"] = log.web_url

    # 2. update and verify
    originals = {q["id"]: q["response"] for q in questions}
    expected = {qid: c["after_value"] for qid, c in resolved.items()}
    untouched = {qid: v for qid, v in originals.items() if qid not in resolved}
    upd = write_eligibility(api, applicant_id, expected, completed)
    problems = [] if upd["ok"] else [f"update returned HTTP {upd['status']}"]
    if upd["ok"]:
        problems = verify_eligibility(api, applicant_id, expected, untouched, completed)
    if problems:
        # Put back the changed answers AND anything else the update disturbed.
        current = {q["id"]: q["response"]
                   for q in eligibility_questions(read_eligibility(api, applicant_id))}
        disturbed = {qid: v for qid, v in untouched.items() if current.get(qid) != v}
        restore = restore_eligibility(api, applicant_id, resolved, ctx, log,
                                      reason="update did not verify",
                                      also_restore=disturbed, is_completed=completed)
        rep["update"] = {"status": upd["status"], "problems": problems, "restore": restore}
        _safe_log(log, log_rows("update_failed", ctx, resolved,
                                notes="; ".join(problems)[:500]))
        return {"status": "error", "report": rep, "applied": None,
                "message": "The eligibility update did not verify; see report.eligibility.update."}

    _safe_log(log, log_rows("eligibility_updated", ctx, resolved))
    rep["update"] = {"status": upd["status"], "verified": True}
    return {"status": "updated", "report": rep,
            "applied": {"resolved": resolved, "untouched": untouched,
                        "is_completed": completed}}


def restore_eligibility(api: Api, applicant_id: int, resolved: Dict[str, Dict[str, Any]],
                        ctx: Dict[str, Any], log: ChangeLog, reason: str,
                        also_restore: Optional[Dict[str, Any]] = None,
                        is_completed: Optional[bool] = None) -> Dict[str, Any]:
    originals = {qid: c["before_value"] for qid, c in resolved.items()}
    originals.update(also_restore or {})
    res = write_eligibility(api, applicant_id, originals, is_completed)
    problems = verify_eligibility(api, applicant_id, originals, {}, is_completed) if res["ok"] \
        else [f"restore returned HTTP {res['status']}"]
    event = "restored" if not problems else "restore_failed"
    _safe_log(log, log_rows(event, ctx, resolved, use_after=False,
                            notes=reason + ("" if not problems else
                                            " | " + "; ".join(problems)[:400])))
    return {"restored": not problems, "problems": problems}


def _safe_log(log: ChangeLog, rows: List[List[Any]]) -> Optional[str]:
    try:
        log.append(rows)
        return None
    except Exception as exc:
        return str(exc)[:300]


# --------------------------------------------------------------------------
# test mode: confirm the eligibility update format on a TEST applicant
# --------------------------------------------------------------------------

def run_update_test(api: Api, test_user_id: int, changes: Optional[Dict[str, str]],
                    log_to_sharepoint: bool) -> Dict[str, Any]:
    user = api.get(f"/api/users/{test_user_id}/")
    if not user["ok"]:
        return {"status": "error", "message": f"Could not read user {test_user_id}: "
                                              f"HTTP {user['status']}"}
    u = user["body"] if isinstance(user["body"], dict) else {}
    marker = " ".join(str(u.get(k) or "") for k in ("first_name", "last_name", "email")).lower()
    if "test" not in marker:
        return {"status": "error",
                "message": "This user is not marked as a test account (no 'test' in "
                           "name or email). The test was not run."}

    elig_body = read_eligibility(api, test_user_id)
    completed = elig_body.get("is_completed")
    questions = eligibility_questions(elig_body)
    changes = changes or {"jurisdiction": "No"}
    try:
        resolved = resolve_eligibility_changes(questions, changes)
    except ValueError as exc:
        return {"status": "error", "message": str(exc)}
    if not resolved:
        return {"status": "error",
                "message": "The test change matches the current answers; choose a "
                           "different answer in eligibility_changes."}

    ctx = {"run_id": f"TEST-{test_user_id}-{datetime.now(timezone.utc):%Y%m%d%H%M%S}",
           "reference": "TEST", "application_id": "", "applicant_id": test_user_id,
           "reason_code": "format_test", "approved_by": "format test",
           "applicant_confirmation": "test account"}
    log = ChangeLog()
    log_error = _safe_log(log, log_rows("TEST planned", ctx, resolved)) \
        if log_to_sharepoint else "skipped"

    originals = {q["id"]: q["response"] for q in questions}
    expected = {qid: c["after_value"] for qid, c in resolved.items()}
    untouched = {qid: v for qid, v in originals.items() if qid not in resolved}

    upd = write_eligibility(api, test_user_id, expected, completed)
    problems = verify_eligibility(api, test_user_id, expected, untouched,
                                  completed) if upd["ok"] \
        else [f"update returned HTTP {upd['status']}"]

    back = write_eligibility(api, test_user_id, {qid: c["before_value"]
                                                 for qid, c in resolved.items()},
                             completed)
    restore_problems = verify_eligibility(
        api, test_user_id, {qid: c["before_value"] for qid, c in resolved.items()},
        untouched, completed) if back["ok"] else [f"restore returned HTTP {back['status']}"]

    if log_to_sharepoint:
        _safe_log(log, log_rows("TEST result", ctx, resolved, notes=(
            "update verified" if not problems else "; ".join(problems)[:300])))
        _safe_log(log, log_rows("TEST restored", ctx, resolved, use_after=False, notes=(
            "restore verified" if not restore_problems else
            "; ".join(restore_problems)[:300])))

    passed = upd["ok"] and not problems and back["ok"] and not restore_problems
    return {
        "status": "passed" if passed else "failed",
        "update": {"http_status": upd["status"], "problems": problems,
                   "error_body": upd["body"]},
        "restore": {"http_status": back["status"], "problems": restore_problems,
                    "error_body": back["body"]},
        "change_log": log.web_url if log_to_sharepoint else "",
        "change_log_error": log_error,
        "next_step": ("Set ELIGIBILITY_UPDATE_FORMAT_CONFIRMED = True." if passed else
                      "Do not enable eligibility changes. A 403 means the token's "
                      "admin account lacks the Assume role permission. A 400, or "
                      "an update that did not verify, may mean the endpoint wants "
                      "option labels instead of indices; share error_body."),
    }


# --------------------------------------------------------------------------
# creation
# --------------------------------------------------------------------------

def find_existing_target_application(api: Api, program_id: int,
                                     applicant_id: int,
                                     max_pages: int = 6) -> Dict[str, Any]:
    """
    Report any application the applicant already holds in the target program.

    The filter parameters are undocumented, so every row is re-checked
    client-side; a silently ignored filter cannot produce a false positive.
    """
    found: List[Dict[str, Any]] = []
    res = api.get("/api/applications/",
                  params={"program": program_id, "applicant": applicant_id,
                          "per_page": 100})
    pages_read = 0
    body = res["body"] if res["ok"] else {}

    while isinstance(body, dict) and pages_read < max_pages:
        pages_read += 1
        for row in body.get("results", []) or []:
            if not isinstance(row, dict):
                continue
            owner = row.get("applicant")
            oid = owner.get("id") if isinstance(owner, dict) else owner
            prog = row.get("program")
            pid = prog.get("id") if isinstance(prog, dict) else prog
            # Require an explicit program match. Accepting rows with no
            # program field lets legacy-portal applications through.
            if oid == applicant_id and pid == program_id:
                found.append({"id": row.get("id"),
                              "reference_id": row.get("reference_id"),
                              "status": (row.get("status") or {}).get("name")
                              if isinstance(row.get("status"), dict) else row.get("status")})
        nxt = body.get("next")
        if not nxt or found:
            break
        res = api.get(nxt)
        body = res["body"] if res["ok"] else {}

    return {"found": found, "pages_read": pages_read,
            "lookup_ok": bool(res.get("ok"))}


def _app_summary(b: Any) -> Any:
    """Application fields safe to keep in job output (no applicant details)."""
    if not isinstance(b, dict):
        return b
    prog = b.get("program")
    stage = b.get("current_stage")
    return {"id": b.get("id"), "reference_id": b.get("reference_id"),
            "program": prog.get("name") if isinstance(prog, dict) else prog,
            "current_stage": stage.get("title") if isinstance(stage, dict) else stage,
            "created_at": b.get("created_at")}


def create_application(api: Api, program_id: int, applicant_id: int,
                       title: Optional[str] = None,
                       override: Optional[dict] = None) -> Dict[str, Any]:
    """
    POST /api/applications/ per the API documentation. `applicant` and
    `program` are the required fields; sending `user` instead returns
    400 "applicant: user is required".

    The admin account behind the token needs the Create Applications
    permission, and the applicant must meet the program's eligibility criteria.
    """
    body = override or {"applicant": applicant_id, "program": program_id}
    if title and not override:
        body["title"] = title

    res = api.post("/api/applications/", json_body=body)
    attempt = {"path": "/api/applications/", "payload": body, "status": res["status"],
               "body": _app_summary(res.get("body")) if res["ok"] else res.get("body")}

    if res["ok"]:
        new_id = (res["body"] or {}).get("id")
        if new_id:
            return {"created": True, "application_id": new_id,
                    "reference_id": (res["body"] or {}).get("reference_id"),
                    "attempt": attempt}
        return {"created": False, "attempt": attempt,
                "reason": "create returned 2xx but no application id"}

    reason = None
    text = ""
    b = res.get("body") or {}
    if isinstance(b, dict):
        text = str(b.get("text") or json.dumps(b))
    if "not eligible for this program" in text:
        reason = ("SMApply rejected the applicant as not eligible for the target "
                  "program. Their eligibility answers do not meet its criteria, "
                  "or the program's application limit has been reached.")
    elif res["status"] == 403:
        reason = ("403 from SMApply. Check that the token's admin account has "
                  "the Create Applications permission.")
    elif res["status"] == 400:
        reason = "SMApply rejected the payload; see attempt.body for the field errors."

    return {"created": False, "attempt": attempt, "reason": reason}


# --------------------------------------------------------------------------
# writing
# --------------------------------------------------------------------------

SINGLE_CHOICE_ERROR = "single-choice value must be an integer index"


def coerce_single_choice(value: Any, target_field: Dict[str, Any]) -> Optional[int]:
    """
    A field can be multi-choice in one program and single-choice in the other,
    so the source may hold [1] where the target wants 1, or a label where the
    target wants its index. Returns the index, or None if it cannot be derived.
    """
    if isinstance(value, list):
        if len(value) == 1 and _is_index(value[0]) and value[0] != CHOICE_UNANSWERED:
            return value[0]
        return None
    if isinstance(value, str):
        s = value.strip()
        if s.isdigit():
            return int(s)
        for i, label in enumerate(choice_labels(target_field)):
            if _norm(label) == _norm(s):
                return i
    return None


def write_task_data(api: Api, app_id: int, task_id: int, payload: Dict[str, Any],
                    target_fields: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """
    PATCH the task's answers. SMApply rejects the WHOLE request if any one
    field is invalid, so on a 400 that names fields, each named field is
    repaired (single-choice shape) or left out, and the rest is sent again.
    """
    if not payload:
        return {"skipped": "empty payload"}
    payload = dict(payload)
    repaired: List[Dict[str, Any]] = []
    dropped: List[Dict[str, Any]] = []
    res: Dict[str, Any] = {}
    for _ in range(4):
        res = api.patch(f"/api/applications/{app_id}/tasks/{task_id}/",
                        json_body={"data": payload})
        if res["ok"]:
            out = {"status": res["status"], "ok": True, "body": None}
            if repaired:
                out["repaired"] = repaired
            if dropped:
                out["dropped"] = dropped
            return out
        body = res.get("body")
        bad = {k: v for k, v in body.items() if k in payload} \
            if res["status"] == 400 and isinstance(body, dict) else {}
        if not bad:
            break
        for key, msg in bad.items():
            fixed = None
            if SINGLE_CHOICE_ERROR in str(msg):
                fixed = coerce_single_choice(payload[key], (target_fields or {}).get(key, {}))
            if fixed is not None:
                repaired.append({"key": key, "from_type": type(payload[key]).__name__,
                                 "to_index": fixed})
                payload[key] = fixed
            else:
                dropped.append({"key": key, "error": str(msg)[:200]})
                payload.pop(key)
        if not payload:
            break
    return {"status": res.get("status"), "ok": False, "body": res.get("body"),
            "repaired": repaired, "dropped": dropped}


def complete_task(api: Api, app_id: int, task_id: int,
                  automations: bool) -> Dict[str, Any]:
    """Status must be a SEPARATE patch; combining data and status returns 500."""
    res = api.patch(f"/api/applications/{app_id}/tasks/{task_id}/",
                    json_body={"status": TASK_STATUS_COMPLETE},
                    automations=automations)
    return {"status": res["status"], "ok": res["ok"],
            "body": None if res["ok"] else res.get("body")}


def copy_file(api: Api, src_url: str, filename: str, app_id: int, task_id: int,
              target_key: Optional[str], max_bytes: int) -> Dict[str, Any]:
    """Download from the source application, re-post to the target task."""
    url = src_url if str(src_url).startswith("http") else f"{BASE}{src_url}"
    try:
        r = requests.get(url, headers={"Authorization": f"Bearer {api.token}"},
                         timeout=UPLOAD_TIMEOUT, allow_redirects=True)
    except Exception as exc:
        return {"filename": filename, "ok": False, "error": f"download failed: {exc}"}
    if r.status_code != 200:
        return {"filename": filename, "ok": False,
                "error": f"download HTTP {r.status_code}"}
    blob = r.content
    if len(blob) > max_bytes:
        return {"filename": filename, "ok": False,
                "error": f"{len(blob)} bytes exceeds max_file_mb"}

    path = f"/api/applications/{app_id}/tasks/{task_id}/files/"
    # SMApply answers an upload without a recognised question reference with
    # 400 "invalid question id". The form-field name is undocumented here, so
    # the likely names are tried in turn; a rejected attempt stores nothing.
    tries: List[Optional[Dict[str, Any]]] = []
    if target_key:
        base_key, _ = split_suffix(target_key)
        for name in FILE_QUESTION_FIELD_NAMES:
            tries.append({name: target_key})
            if base_key != target_key:
                tries.append({name: base_key})
    tries.append(None)

    attempts: List[Dict[str, Any]] = []
    for extra in tries:
        res = api.post(path,
                       files={"file": (filename, io.BytesIO(blob),
                                       "application/octet-stream")},
                       data=extra, timeout=UPLOAD_TIMEOUT)
        if res["ok"]:
            return {"filename": filename, "ok": True, "bytes": len(blob),
                    "form_data": extra, "response": res.get("body")}
        attempts.append({"form_data": extra, "status": res["status"],
                         "body": res.get("body")})
    return {"filename": filename, "ok": False, "attempts": attempts}


# --------------------------------------------------------------------------
# main
# --------------------------------------------------------------------------

def main(
    source_application_id: int = 0,
    target_program: str = "auto",
    dry_run: bool = True,
    copy_files: bool = True,
    mark_tasks_complete: bool = False,
    fire_automations: bool = False,
    ambiguous_field_policy: str = "skip",
    include_empty_fields: bool = False,
    copy_signature_fields: bool = False,
    target_application_id: int = 0,
    reuse_existing_target: bool = False,
    reference_target_application_id: int = 0,
    create_payload_override: Optional[dict] = None,
    max_file_mb: int = 50,
    extra_exclude_task_names: Optional[List[str]] = None,
    only_task_names: Optional[List[str]] = None,
    eligibility_changes: Optional[dict] = None,
    approved_by: str = "",
    applicant_confirmation: str = "",
    reason_code: str = "partner_contract_termination",
    restore_eligibility_on_failure: bool = True,
    mode: str = "transfer",
    test_user_id: int = 0,
    test_log_to_sharepoint: bool = True,
) -> dict:
    """
    source_application_id   The application to copy FROM.
    target_program          "auto" (the other stream), "tm", or "ces".
    dry_run                 True = check and plan only, no writes at all.
    copy_files              Re-upload attachments on copied tasks.
    mark_tasks_complete     Also mark each written task complete.
    fire_automations        Send apply-automations: full on the completion
                            PATCH. Only meaningful with mark_tasks_complete.
    ambiguous_field_policy  "skip" (default) or "positional". See module docs.
    include_empty_fields    Write blanks as well as answered fields.
    target_application_id   Populate THIS existing application instead of
                            creating one (skips the eligibility check).
    reuse_existing_target   If the applicant already has a target-stream
                            application, populate it rather than creating
                            another. Off by default, because it overwrites
                            answers.
    copy_signature_fields   Re-upload file-typed fields such as the Terms and
                            Conditions signature. Off by default: an attestation
                            signature given for one application should normally
                            be collected again for the new one.
    reference_target_application_id
                            Dry runs need an existing application in the target
                            program to read its field keys. Left at 0, one is
                            picked automatically.
    create_payload_override Exact POST body for the create, if ever needed.
    only_task_names         Copy only these tasks (exact names), e.g. to finish
                            tasks that failed, with target_application_id set
                            to the application already created.
    eligibility_changes     Only when the applicant does not meet the target
                            criteria and has CONFIRMED different answers,
                            e.g. {"ofac": "No"}. Keys: eligibility_to_apply,
                            prior_report, ofac, jurisdiction.
    approved_by             Required with eligibility_changes.
    applicant_confirmation  Required with eligibility_changes: where the
                            applicant's confirmation is recorded.
    reason_code             Recorded in the change log.
    restore_eligibility_on_failure
                            If the create fails after an eligibility change,
                            put the original answers back (logged).
    mode                    "transfer" (default) or "test_eligibility_update".
    test_user_id            For mode="test_eligibility_update": a TEST
                            applicant's user id. Changes one answer, verifies,
                            and restores it.
    test_log_to_sharepoint  In test mode, also test appending to the change log.
    """
    started = datetime.now(timezone.utc)
    api = Api(_get_token())

    if mode == "test_eligibility_update":
        if not test_user_id:
            return {"status": "error", "message": "Pass test_user_id."}
        return run_update_test(api, test_user_id, eligibility_changes,
                               test_log_to_sharepoint)
    if mode != "transfer":
        return {"status": "error", "message": "mode must be 'transfer' or "
                                              "'test_eligibility_update'."}
    if not source_application_id:
        return {"status": "error", "message": "Pass source_application_id."}

    if ambiguous_field_policy not in ("skip", "positional"):
        return {"status": "error",
                "message": "ambiguous_field_policy must be 'skip' or 'positional'."}

    exclude_extra = [norm_task_name(n) for n in (extra_exclude_task_names or [])]
    only = {norm_task_name(n) for n in (only_task_names or [])}

    def excluded(name: Any) -> bool:
        n = norm_task_name(name)
        return is_excluded(name) or any(sub in n for sub in exclude_extra)

    # ---- source
    src_app = fetch_application(api, source_application_id)
    src_program_id = ((src_app.get("program") or {}).get("id"))
    applicant = src_app.get("applicant") or {}
    applicant_id = applicant.get("id")
    src_stream = PROGRAM_BY_ID.get(src_program_id)

    if src_stream is None:
        return {"status": "error",
                "message": (f"Application {source_application_id} is in program "
                            f"{src_program_id}, which is neither CES ({PROGRAMS['ces']}) "
                            f"nor TM ({PROGRAMS['tm']}).")}
    if not applicant_id:
        return {"status": "error",
                "message": "Could not resolve the applicant id on the source application."}

    # ---- direction
    tgt_stream = ("tm" if src_stream == "ces" else "ces") \
        if target_program == "auto" else target_program.strip().lower()
    if tgt_stream not in PROGRAMS:
        return {"status": "error",
                "message": f"target_program must be auto/tm/ces (got '{target_program}')."}
    if tgt_stream == src_stream:
        return {"status": "error",
                "message": (f"Source is already in the {src_stream.upper()} program. "
                            f"This script only clones cross-stream.")}
    tgt_program_id = PROGRAMS[tgt_stream]

    report: Dict[str, Any] = {
        "source_application_id": source_application_id,
        "source_reference_id": src_app.get("reference_id"),
        "source_stream": src_stream.upper(),
        "target_stream": tgt_stream.upper(),
        "target_program_id": tgt_program_id,
        "applicant_id": applicant_id,
        "dry_run": dry_run,
        "ambiguous_field_policy": ambiguous_field_policy,
    }
    warnings: List[str] = []

    ctx = {
        "run_id": f"{src_app.get('reference_id') or source_application_id}-"
                  f"{started:%Y%m%d%H%M%S}",
        "reference": src_app.get("reference_id") or "",
        "application_id": source_application_id,
        "applicant_id": applicant_id,
        "reason_code": (reason_code or "").strip(),
        "approved_by": (approved_by or "").strip(),
        "applicant_confirmation": (applicant_confirmation or "").strip(),
    }
    log = ChangeLog()
    applied: Optional[Dict[str, Any]] = None

    # ---- eligibility, existing applications, create
    new_app_id: Optional[int] = None
    existing: Dict[str, Any] = {"found": []}
    if not dry_run:
        existing = find_existing_target_application(api, tgt_program_id, applicant_id)
        report["existing_target_applications"] = existing["found"]

    needs_create = not target_application_id and not (reuse_existing_target and existing["found"])
    if needs_create:
        gate = eligibility_gate(api, tgt_stream, applicant_id, ctx,
                                eligibility_changes, dry_run, log)
        report["eligibility"] = gate["report"]
        if gate["status"] in ("blocked", "error"):
            return {"status": f"eligibility_{gate['status']}",
                    "message": gate.get("message"), "report": report}
        applied = gate.get("applied")

    # ---- source tasks
    src_tasks = fetch_all_tasks(api, source_application_id)
    report["source_task_count"] = len(src_tasks)

    if dry_run:
        ref_id = reference_target_application_id or pick_reference_application(
            api, tgt_program_id, exclude_id=source_application_id)
        if not ref_id:
            return {"status": "error",
                    "message": (f"Dry run needs an existing {tgt_stream.upper()} "
                                f"application to read the target field keys, and "
                                f"none could be found. Pass "
                                f"reference_target_application_id."),
                    "report": report}
        report["reference_target_application_id"] = ref_id
        tgt_tasks = fetch_all_tasks(api, ref_id)
    else:
        if target_application_id:
            new_app_id = target_application_id
            report["mode"] = "populate_supplied_application"
        elif reuse_existing_target and existing["found"]:
            new_app_id = existing["found"][0]["id"]
            report["mode"] = "populate_existing_application"
            warnings.append(
                f"Populating existing {tgt_stream.upper()} application "
                f"{new_app_id}; any answers already there will be overwritten.")
        else:
            if existing["found"]:
                warnings.append(
                    f"Applicant already holds {len(existing['found'])} "
                    f"{tgt_stream.upper()} application(s): "
                    f"{[a['reference_id'] for a in existing['found']]}. "
                    f"Creating another anyway.")
            creation = create_application(
                api, tgt_program_id, applicant_id,
                title=src_app.get("title"), override=create_payload_override)
            report["creation"] = creation
            if not creation.get("created"):
                msg = creation.get("reason") or "SMApply rejected the create."
                if applied:
                    if restore_eligibility_on_failure:
                        report["eligibility"]["restore"] = restore_eligibility(
                            api, applicant_id, applied["resolved"], ctx, log,
                            reason="create failed after eligibility change",
                            is_completed=applied.get("is_completed"))
                        msg += " The applicant's original eligibility answers were restored."
                    else:
                        _safe_log(log, log_rows("transfer_failed", ctx, applied["resolved"],
                                                notes="eligibility change left in place"))
                        msg += " The eligibility change was left in place (logged)."
                return {"status": "create_failed", "message": msg, "report": report}
            new_app_id = creation["application_id"]
            report["mode"] = "created_new_application"
            if applied:
                err = _safe_log(log, log_rows("transfer_created", ctx, applied["resolved"],
                                              new_app_id=new_app_id))
                if err:
                    warnings.append(f"Transfer succeeded but the change log could not "
                                    f"record it: {err}")

        report["new_application_id"] = new_app_id
        tgt_tasks = fetch_all_tasks(api, new_app_id)

    report["target_task_count"] = len(tgt_tasks)

    # ---- plan each task
    plans: List[Dict[str, Any]] = []
    unmatched_tasks: List[str] = []

    for sname, stask in src_tasks.items():
        if excluded(stask.get("name")):
            continue
        if only and norm_task_name(stask.get("name")) not in only:
            continue
        tname = sname if sname in tgt_tasks else TASK_NAME_ALIASES.get(sname, "")
        ttask = tgt_tasks.get(tname) if tname else None
        if ttask is None:
            if ((stask.get("instance") or {}).get("data")):
                unmatched_tasks.append(stask.get("name"))
            continue
        plan = build_task_payload(stask, ttask, ambiguous_field_policy,
                                  include_empty_fields, copy_signature_fields)
        if (plan["fields_written"] or plan["files"] or plan["invalid_choices"]
                or plan["signature_fields"]):
            plans.append(plan)

    for name in unmatched_tasks:
        warnings.append(f"Task '{name}' has data but no counterpart in "
                        f"{tgt_stream.upper()}; not copied.")
    manual_followup: List[Dict[str, Any]] = []

    for p in plans:
        for a in p["ambiguous"]:
            warnings.append(
                f"{p['name']}: {len(a['source_keys'])} fields share the label "
                f"'{a['label'][:60]}' on both sides; {a['resolution']}.")
            manual_followup.append({"task": p["name"], "reason": "ambiguous repeat block",
                                    "label": a["label"][:70]})
        n_drop = len(p["unmappable"])
        if n_drop:
            warnings.append(
                f"{p['name']}: {n_drop} field(s) exist only in "
                f"{src_stream.upper()} and were dropped "
                f"(e.g. {[u['label'][:40] for u in p['unmappable'][:3]]}).")
        for ic in p["invalid_choices"]:
            warnings.append(
                f"{p['name']}: '{ic['label']}' answer {ic['source_choice'] or ic['indices']} "
                f"has no equivalent option in {tgt_stream.upper()} "
                f"({ic['target_option_count']} options); left blank.")
            manual_followup.append({"task": p["name"], "reason": "option missing in target",
                                    "label": ic["label"],
                                    "source_answer": ic["source_choice"] or ic["indices"]})
        for sf in p["signature_fields"]:
            warnings.append(
                f"{p['name']}: '{sf['label'][:60]}' is a signature/file field; not "
                f"copied. The applicant must sign again on the new application.")
            manual_followup.append({"task": p["name"], "reason": "signature not copied",
                                    "label": sf["label"][:70]})

    report["manual_followup"] = manual_followup

    report["planned_tasks"] = [
        {"name": p["name"], "source_task_id": p["source_task_id"],
         "target_task_id": p["target_task_id"],
         "fields_written": p["fields_written"],
         "files": len(p["files"]),
         "fields_dropped": len(p["unmappable"]),
         "fields_ambiguous": sum(len(a["source_keys"]) for a in p["ambiguous"]),
         "invalid_choices": len(p["invalid_choices"]),
         "signature_fields": len(p["signature_fields"])}
        for p in plans
    ]
    report["totals"] = {
        "tasks": len(plans),
        "fields": sum(p["fields_written"] for p in plans),
        "files": sum(len(p["files"]) for p in plans),
        "dropped": sum(len(p["unmappable"]) for p in plans),
        "ambiguous": sum(sum(len(a["source_keys"]) for a in p["ambiguous"]) for p in plans),
        "invalid_choices": sum(len(p["invalid_choices"]) for p in plans),
        "signatures_skipped": sum(len(p["signature_fields"]) for p in plans),
    }

    if dry_run:
        report["warnings"] = warnings
        report["sample_payload_keys"] = {p["name"]: list(p["payload"])[:5] for p in plans[:3]}
        report["elapsed_s"] = round(
            (datetime.now(timezone.utc) - started).total_seconds(), 1)
        return {"status": "dry_run", "report": report}

    # ---- write
    results: List[Dict[str, Any]] = []
    max_bytes = max_file_mb * 1024 * 1024

    for p in plans:
        entry = {"name": p["name"], "target_task_id": p["target_task_id"]}
        entry["data"] = write_task_data(api, new_app_id, p["target_task_id"], p["payload"],
                                        p.get("target_fields"))
        if not entry["data"].get("ok", True):
            warnings.append(f"{p['name']}: data PATCH returned "
                            f"HTTP {entry['data'].get('status')}; task not written.")
        for d in entry["data"].get("dropped") or []:
            warnings.append(f"{p['name']}: field {d['key']} rejected by SMApply "
                            f"({d['error']}); left blank for manual entry.")
            manual_followup.append({"task": p["name"], "reason": "field rejected by SMApply",
                                    "key": d["key"]})

        if copy_files and p["files"]:
            entry["files"] = []
            for f in p["files"]:
                meta = f["file"]
                entry["files"].append(copy_file(
                    api, meta.get("url", ""), meta.get("filename", "file"),
                    new_app_id, p["target_task_id"], f.get("target_key"), max_bytes))
            failed = [x for x in entry["files"] if not x.get("ok")]
            if failed:
                warnings.append(f"{p['name']}: {len(failed)} of "
                                f"{len(entry['files'])} file(s) failed to copy.")

        if mark_tasks_complete and entry["data"].get("ok", False):
            entry["complete"] = complete_task(api, new_app_id, p["target_task_id"],
                                              fire_automations)
        results.append(entry)

    report["results"] = results
    report["warnings"] = warnings
    report["elapsed_s"] = round((datetime.now(timezone.utc) - started).total_seconds(), 1)

    failures = [r for r in results if not r.get("data", {}).get("ok", True)]
    return {"status": "ok" if not failures else "partial",
            "new_application_id": new_app_id,
            "report": report}