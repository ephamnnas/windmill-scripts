# py: 3.12
# requirements:
# wmill
# requests

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
dry_run=True (the default) performs NO writes. It resolves the full task and
field crosswalk, builds every payload, and returns exactly what it would send,
including a per-task count of fields that would be written, skipped as
ambiguous, or dropped as unmappable. Run it dry first and read `warnings`.

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
order instead -- this WILL scramble block order if the sort happens not to
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
    }


# --------------------------------------------------------------------------
# creation
# --------------------------------------------------------------------------

def find_existing_target_application(api: Api, program_id: int,
                                     applicant_id: int,
                                     max_pages: int = 6) -> Dict[str, Any]:
    """
    Both expedited programs are max_applications=1, so an applicant who already
    holds an application in the target stream cannot receive another. Find it
    rather than letting the create fail with an opaque 403.

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
            "filter_trusted": bool(res.get("ok"))}


def create_application(api: Api, program_id: int, applicant_id: int,
                       title: Optional[str] = None,
                       override: Optional[dict] = None) -> Dict[str, Any]:
    """
    POST /api/applications/ per the API documentation. `applicant` and
    `program` are the required fields; sending `user` instead returns
    400 "applicant: user is required".

    The admin account behind the token needs the Create Applications
    permission, and the applicant must already be a member of the site.
    """
    body = override or {"applicant": applicant_id, "program": program_id}
    if title and not override:
        body["title"] = title

    res = api.post("/api/applications/", json_body=body)
    attempt = {"path": "/api/applications/", "payload": body,
               "status": res["status"], "body": res.get("body")}

    if res["ok"]:
        new_id = (res["body"] or {}).get("id")
        if new_id:
            return {"created": True, "application_id": new_id,
                    "attempt": attempt, "response": res["body"]}
        return {"created": False, "attempt": attempt,
                "reason": "create returned 2xx but no application id"}

    reason = None
    text = ""
    b = res.get("body") or {}
    if isinstance(b, dict):
        text = str(b.get("text") or json.dumps(b))
    if "not eligible for this program" in text:
        reason = ("SMApply rejected the applicant as not eligible. Both "
                  "expedited programs are max_applications=1, so the usual "
                  "cause is that this applicant already holds an application "
                  "in the target stream.")
    elif res["status"] == 403:
        reason = ("403 from SMApply. Check that the token's admin account has "
                  "the Create Applications permission.")
    elif res["status"] == 400:
        reason = "SMApply rejected the payload; see attempt.body for the field errors."

    return {"created": False, "attempt": attempt, "reason": reason}


# --------------------------------------------------------------------------
# writing
# --------------------------------------------------------------------------

def write_task_data(api: Api, app_id: int, task_id: int,
                    payload: Dict[str, Any]) -> Dict[str, Any]:
    if not payload:
        return {"skipped": "empty payload"}
    res = api.patch(f"/api/applications/{app_id}/tasks/{task_id}/",
                    json_body={"data": payload})
    return {"status": res["status"], "ok": res["ok"],
            "body": None if res["ok"] else res.get("body")}


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
    # Field-scoped uploads may need the field key; the association is
    # undocumented, so try with it and fall back to a plain upload.
    tries = []
    if target_key:
        tries.append({"field": target_key})
    tries.append(None)

    for extra in tries:
        res = api.post(path,
                       files={"file": (filename, io.BytesIO(blob),
                                       "application/octet-stream")},
                       data=extra, timeout=UPLOAD_TIMEOUT)
        if res["ok"]:
            return {"filename": filename, "ok": True, "bytes": len(blob),
                    "form_data": extra, "response": res.get("body")}
        tries_err = {"filename": filename, "ok": False, "status": res["status"],
                     "form_data": extra, "body": res.get("body")}
    return tries_err


# --------------------------------------------------------------------------
# main
# --------------------------------------------------------------------------

def main(
    source_application_id: int,
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
    **kwargs: Any,
) -> dict:
    """
    source_application_id   The application to copy FROM.
    target_program          "auto" (the other stream), "tm", or "ces".
    dry_run                 True = resolve and plan only, no writes at all.
    copy_files              Re-upload attachments on copied tasks.
    mark_tasks_complete     Also mark each written task complete.
    fire_automations        Send apply-automations: full on the completion
                            PATCH. Only meaningful with mark_tasks_complete.
    ambiguous_field_policy  "skip" (default) or "positional". See module docs.
    include_empty_fields    Write blanks as well as answered fields.
    target_application_id   Populate THIS existing application instead of
                            creating one. Both programs are max_applications=1,
                            so this is the only route for an applicant who
                            already holds an application in the target stream.
    reuse_existing_target   If the applicant already has a target-stream
                            application, populate it rather than failing.
                            Off by default, because it overwrites answers.
    copy_signature_fields   Re-upload file-typed fields such as the Terms and
                            Conditions signature. Off by default: an attestation
                            signature given for one application should normally
                            be collected again for the new one.
    reference_target_application_id
                            Dry runs need an existing application in the target
                            program to read its field keys. Left at 0, one is
                            picked automatically.
    create_payload_override Bypass the create-candidate ladder with an exact
                            POST body, once the working shape is known.
    """
    started = datetime.now(timezone.utc)
    api = Api(_get_token())

    if ambiguous_field_policy not in ("skip", "positional"):
        return {"status": "error",
                "message": "ambiguous_field_policy must be 'skip' or 'positional'."}

    if extra_exclude_task_names:
        EXCLUDE_SUBSTRINGS.extend(norm_task_name(n) for n in extra_exclude_task_names)

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

    # ---- source tasks
    src_tasks = fetch_all_tasks(api, source_application_id)
    report["source_task_count"] = len(src_tasks)

    # ---- create (or resolve a reference for planning)
    new_app_id: Optional[int] = None
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
        existing = find_existing_target_application(api, tgt_program_id, applicant_id)
        report["existing_target_applications"] = existing["found"]

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
            # An existing application is reported but does not block the
            # create; SMApply decides eligibility, not this script.
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
                if existing["found"]:
                    msg += (f" The applicant already has "
                            f"{len(existing['found'])} application(s) in this "
                            f"program, which is above its max_applications=1 "
                            f"setting, so that is the likely trigger. Either "
                            f"raise the per-applicant limit in the program's "
                            f"SMApply settings, or pass "
                            f"target_application_id / reuse_existing_target "
                            f"to populate one that already exists.")
                return {"status": "create_failed", "message": msg,
                        "report": report}
            new_app_id = creation["application_id"]
            report["mode"] = "created_new_application"

        report["new_application_id"] = new_app_id
        tgt_tasks = fetch_all_tasks(api, new_app_id)

    report["target_task_count"] = len(tgt_tasks)

    # ---- plan each task
    plans: List[Dict[str, Any]] = []
    unmatched_tasks: List[str] = []

    for sname, stask in src_tasks.items():
        if is_excluded(stask.get("name")):
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
                f"'{a['label'][:60]}' on both sides -- {a['resolution']}.")
            manual_followup.append({"task": p["name"], "reason": "ambiguous repeat block",
                                    "label": a["label"][:70]})
        n_drop = sum(1 for u in p["unmappable"])
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
        report["sample_payload"] = (
            {p["name"]: dict(list(p["payload"].items())[:5]) for p in plans[:3]}
        )
        report["elapsed_s"] = round(
            (datetime.now(timezone.utc) - started).total_seconds(), 1)
        return {"status": "dry_run", "report": report}

    # ---- write
    results: List[Dict[str, Any]] = []
    max_bytes = max_file_mb * 1024 * 1024

    for p in plans:
        entry = {"name": p["name"], "target_task_id": p["target_task_id"]}
        entry["data"] = write_task_data(api, new_app_id, p["target_task_id"], p["payload"])
        if not entry["data"].get("ok", True):
            warnings.append(f"{p['name']}: data PATCH returned "
                            f"HTTP {entry['data'].get('status')}.")

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
    return {"status": "ok" if not failures and not warnings else "partial",
            "new_application_id": new_app_id,
            "report": report}