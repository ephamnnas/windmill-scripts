# py: 3.12
# requirements:
# requests
# wmill

"""
Windmill script: screen CES program (ID 295865) applications in SM Apply for
links to restricted countries/regions across residential address, education,
registration/licensing, examination, employment, and issuing bank fields.

Country data lives in each application's tasks:
    GET /api/applications/{id}/tasks/
Each task has a "data" dict of {field_key: {"label": ..., "response": ...}}.
Country questions are either free text ("Iran") or dropdowns stored as an
integer option index (e.g. 30), where -1 means unanswered. Dropdown indexes
are translated through choice_maps (see mode="discover").

Auth uses the same Windmill variable as the credential report script.

Run order:
    1. mode="discover"  -> lists every country dropdown and the indexes in use
    2. Fill in CHOICE_MAPS below with each dropdown's option order
    3. mode="screen"    -> returns flagged applications

The optional choice_maps parameter uses the same format as CHOICE_MAPS and,
when passed, is merged over it (useful for testing a mapping before saving).
"""

import csv
import io
import json
import re
import threading
import time
import unicodedata
from datetime import datetime, timezone
from urllib.parse import quote
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor
from typing import Any, Dict, List, Optional

import requests
import wmill

PROGRAM_ID = 295865
SMAPPLY_BASE = "https://nnas.smapply.io/api"
SMAPPLY_TOKEN_VAR = "u/epham2/SMAPPLY_ACCESS_TOKEN"

# --------------------------------------------------------------------------
# Dropdown index -> country name mappings
# --------------------------------------------------------------------------
# Primary source: at runtime the script reads each task's settings
# (TASK_SETTINGS_PATH) and builds index maps from every country field's
# "choices". Indexes are 0-based and match the stored response values.
#
# Fallback: CHOICE_MAPS below, used for any field whose settings could not be
# loaded. Personal Information (2053922), Registration/Licensing History
# (2053917) and Examination (2053918) are filled in from their settings.
# Other tasks can be added the same way once their settings are confirmed.
TASK_SETTINGS_PATH = "/tasks/{task_id}/"   # GET /api/tasks/{id}/ (confirmed)

# Personal Information address country dropdowns (197 options, index 0 first)
PI_ADDRESS_COUNTRIES: List[str] = [
    'Afghanistan', 'Albania', 'Algeria', 'Andorra', 'Angola', 'Antigua & Deps',
    'Argentina', 'Armenia', 'Australia', 'Austria', 'Azerbaijan', 'Bahamas', 'Bahrain',
    'Bangladesh', 'Barbados', 'Belarus', 'Belgium', 'Belize', 'Benin', 'Bhutan',
    'Bolivia', 'Bosnia Herzegovina', 'Botswana', 'Brazil', 'Brunei', 'Bulgaria',
    'Burkina', 'Burundi', 'Cambodia', 'Cameroon', 'Canada', 'Cape Verde',
    'Central African Rep', 'Chad', 'Chile', 'China', 'Colombia', 'Comoros', 'Congo',
    'Congo {Democratic Rep}', 'Costa Rica', 'Croatia', 'Cuba', 'Cyprus',
    'Czech Republic', 'Denmark', 'Djibouti', 'Dominica', 'Dominican Republic',
    'East Timor', 'Ecuador', 'Egypt', 'El Salvador', 'Equatorial Guinea', 'Eritrea',
    'Estonia', 'Ethiopia', 'Fiji', 'Finland', 'France', 'Gabon', 'Gambia', 'Georgia',
    'Germany', 'Ghana', 'Greece', 'Grenada', 'Guatemala', 'Guinea', 'Guinea-Bissau',
    'Guyana', 'Haiti', 'Honduras', 'Hungary', 'Iceland', 'India', 'Indonesia', 'Iran',
    'Iraq', 'Ireland {Republic}', 'Israel', 'Italy', 'Ivory Coast', 'Jamaica', 'Japan',
    'Jordan', 'Kazakhstan', 'Kenya', 'Kiribati', 'Korea North', 'Korea South',
    'Kosovo', 'Kuwait', 'Kyrgyzstan', 'Laos', 'Latvia', 'Lebanon', 'Lesotho',
    'Liberia', 'Libya', 'Liechtenstein', 'Lithuania', 'Luxembourg', 'Macedonia',
    'Madagascar', 'Malawi', 'Malaysia', 'Maldives', 'Mali', 'Malta',
    'Marshall Islands', 'Mauritania', 'Mauritius', 'Mexico', 'Micronesia', 'Moldova',
    'Monaco', 'Mongolia', 'Montenegro', 'Morocco', 'Mozambique', 'Myanmar, {Burma}',
    'Namibia', 'Nauru', 'Nepal', 'Netherlands', 'New Zealand', 'Nicaragua', 'Niger',
    'Nigeria', 'Norway', 'Oman', 'Pakistan', 'Palau', 'Panama', 'Papua New Guinea',
    'Paraguay', 'Peru', 'Philippines', 'Poland', 'Portugal', 'Qatar', 'Romania',
    'Russian Federation', 'Rwanda', 'St Kitts & Nevis', 'St Lucia',
    'Saint Vincent & the Grenadines', 'Samoa', 'San Marino', 'Sao Tome & Principe',
    'Saudi Arabia', 'Senegal', 'Serbia', 'Seychelles', 'Sierra Leone', 'Singapore',
    'Slovakia', 'Slovenia', 'Solomon Islands', 'Somalia', 'South Africa',
    'South Sudan', 'Spain', 'Sri Lanka', 'Sudan', 'Suriname', 'Swaziland', 'Sweden',
    'Switzerland', 'Syria', 'Taiwan', 'Tajikistan', 'Tanzania', 'Thailand', 'Togo',
    'Tonga', 'Trinidad & Tobago', 'Tunisia', 'Turkey', 'Turkmenistan', 'Tuvalu',
    'Uganda', 'Ukraine', 'United Arab Emirates', 'United Kingdom', 'United States',
    'Uruguay', 'Uzbekistan', 'Vanuatu', 'Vatican City', 'Venezuela', 'Vietnam',
    'Yemen', 'Zambia', 'Zimbabwe', 'Not found',
]

# Initial nursing education country: same list plus "Palestine" at index 197
PI_EDUCATION_COUNTRIES: List[str] = PI_ADDRESS_COUNTRIES + ["Palestine"]

# Registration/Licensing History country dropdowns (task 2053917): a different
# list (191 options, no "Not found"), shared by the current registration
# country and all six "previously registered" repeat blocks.
REG_COUNTRIES: List[str] = [
    'Afghanistan', 'Albania', 'Algeria', 'Andorra', 'Angola', 'Antarctica',
    'Antigua and Barbuda', 'Argentina', 'Armenia', 'Australia', 'Austria',
    'Azerbaijan', 'Bahamas', 'Bahrain', 'Bangladesh', 'Barbados', 'Belarus', 'Belgium',
    'Belize', 'Benin', 'Bermuda', 'Bhutan', 'Bolivia', 'Bosnia and Herzegovina',
    'Botswana', 'Brazil', 'Brunei', 'Bulgaria', 'Burkina Faso', 'Burma', 'Burundi',
    'Cambodia', 'Cameroon', 'Canada', 'Cape Verde', 'Central African Republic', 'Chad',
    'Chile', 'China', 'Colombia', 'Comoros', 'Congo, Democratic Republic',
    'Congo, Republic of the', 'Costa Rica', "Cote d'Ivoire", 'Croatia', 'Cuba',
    'Cyprus', 'Czech Republic', 'Denmark', 'Djibouti', 'Dominica',
    'Dominican Republic', 'East Timor', 'Ecuador', 'Egypt', 'El Salvador',
    'Equatorial Guinea', 'Eritrea', 'Estonia', 'Ethiopia', 'Fiji', 'Finland', 'France',
    'Gabon', 'Gambia', 'Georgia', 'Germany', 'Ghana', 'Greece', 'Greenland', 'Grenada',
    'Guatemala', 'Guinea', 'Guinea-Bissau', 'Guyana', 'Haiti', 'Honduras', 'Hong Kong',
    'Hungary', 'Iceland', 'India', 'Indonesia', 'Iran', 'Iraq', 'Ireland', 'Israel',
    'Italy', 'Jamaica', 'Japan', 'Jordan', 'Kazakhstan', 'Kenya', 'Kiribati',
    'Korea, North', 'Korea, South', 'Kuwait', 'Kyrgyzstan', 'Laos', 'Latvia',
    'Lebanon', 'Lesotho', 'Liberia', 'Libya', 'Liechtenstein', 'Lithuania',
    'Luxembourg', 'Macedonia', 'Madagascar', 'Malawi', 'Malaysia', 'Maldives', 'Mali',
    'Malta', 'Marshall Islands', 'Mauritania', 'Mauritius', 'Mexico', 'Micronesia',
    'Moldova', 'Mongolia', 'Morocco', 'Monaco', 'Mozambique', 'Namibia', 'Nauru',
    'Nepal', 'Netherlands', 'New Zealand', 'Nicaragua', 'Niger', 'Nigeria', 'Norway',
    'Oman', 'Pakistan', 'Panama', 'Papua New Guinea', 'Paraguay', 'Peru',
    'Philippines', 'Poland', 'Portugal', 'Qatar', 'Romania', 'Russia', 'Rwanda',
    'Samoa', 'San Marino', 'Sao Tome', 'Saudi Arabia', 'Senegal',
    'Serbia and Montenegro', 'Seychelles', 'Sierra Leone', 'Singapore', 'Slovakia',
    'Slovenia', 'Solomon Islands', 'Somalia', 'South Africa', 'Spain', 'Sri Lanka',
    'Sudan', 'Suriname', 'Swaziland', 'Sweden', 'Switzerland', 'Syria', 'Taiwan',
    'Tajikistan', 'Tanzania', 'Thailand', 'Togo', 'Tonga', 'Trinidad and Tobago',
    'Tunisia', 'Turkey', 'Turkmenistan', 'Uganda', 'Ukraine', 'United Arab Emirates',
    'United Kingdom', 'United States', 'Uruguay', 'Uzbekistan', 'Vanuatu', 'Venezuela',
    'Vietnam', 'Yemen', 'Zambia', 'Zimbabwe',
]

# Examination country dropdown (task 2053918): same list and order as the
# Personal Information address list, without the trailing "Not found" option
# (196 options, index 0-195).
EXAM_COUNTRIES: List[str] = PI_ADDRESS_COUNTRIES[:196]

CHOICE_MAPS: Dict[str, Any] = {
    "lists": {
        "pi_address_countries":   PI_ADDRESS_COUNTRIES,
        "pi_education_countries": PI_EDUCATION_COUNTRIES,
        "reg_countries":          REG_COUNTRIES,
        "exam_countries":         EXAM_COUNTRIES,
    },
    "fields": {
        "odQBnoDDlJ_5": "pi_address_countries",    # Current Address | Country
        "Esldo5mQWc_5": "pi_address_countries",    # Mailing Address | Country
        "KOs3HEE6sr":   "pi_education_countries",  # Country of initial nursing education

        # Registration/Licensing History (task 2053917)
        "GNviXZh7zt":   "reg_countries",           # Country registered/licensed as a nurse
        "78nblFoLys":   "reg_countries",           # Previously registered country, block 1
        "0gqFv2wuRp":   "reg_countries",           # block 2
        "TqaiTTrVqd":   "reg_countries",           # block 3
        "pkmuzT9P75":   "reg_countries",           # block 4
        "Je7rLX84r5":   "reg_countries",           # block 5
        "pr11MJ2OXs":   "reg_countries",           # block 6

        # Examination (task 2053918)
        "tagqtsheVy":   "exam_countries",          # Country where NCLEX-RN was passed
    },
}

# Dropdown selections that are not restricted but need a person to look at:
# Ukraine (occupied regions may not be named in the free-text fields) and
# "Not found" (applicant could not find their country in the list).
REVIEW_COUNTRIES = {"ukraine": "Ukraine (confirm region)", "not found": "Country not in list"}

# --------------------------------------------------------------------------
# Restricted list and spelling variants
# --------------------------------------------------------------------------
COUNTRY_ALIASES = {
    "Cuba": ["cuba", "republic of cuba"],
    "Iran": ["iran", "iran, islamic republic of", "islamic republic of iran", "persia"],
    "North Korea": [
        "north korea", "korea north", "korea, north", "dprk", "korea dpr", "korea, dpr",
        "democratic people's republic of korea",
        "korea, democratic people's republic of",
    ],
    "Venezuela": ["venezuela", "venezuela, bolivarian republic of", "bolivarian republic of venezuela"],
}

# Nationality adjectives, used only for free-text names (regulators, banks)
DEMONYMS = {
    "Cuba": ["cuban"],
    "Iran": ["iranian"],
    "North Korea": ["north korean"],
    "Venezuela": ["venezuelan"],
}

# Occupied regions are usually entered under "Ukraine", so these are matched
# against province/state and city fields.
REGION_ALIASES = {
    "Crimea (Ukraine)": [
        "crimea", "krym", "republic of crimea", "autonomous republic of crimea",
        "sevastopol", "simferopol", "kerch", "yalta", "yevpatoria", "feodosia",
    ],
    "Donetsk People's Republic": [
        "donetsk", "donetsk people's republic", "dnr", "dpr",
        "makiivka", "horlivka", "gorlovka", "mariupol",
    ],
    "Luhansk People's Republic": [
        "luhansk", "lugansk", "luhansk people's republic", "lugansk people's republic",
        "lnr", "lpr", "alchevsk", "sievierodonetsk", "severodonetsk",
    ],
}


def normalize(text):
    if text is None:
        return ""
    text = str(text).replace("\u2019", "'").replace("\u2018", "'")
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode()
    return " ".join(text.lower().split())


def build_matcher(*alias_dicts):
    pairs = []
    for d in alias_dicts:
        for label, aliases in d.items():
            for a in aliases:
                pairs.append((normalize(a), label))
    pairs.sort(key=lambda p: -len(p[0]))  # longest alias wins
    compiled = [(re.compile(r"(?<![a-z])" + re.escape(a) + r"(?![a-z])"), label) for a, label in pairs]

    def match(value):
        v = normalize(value)
        if not v:
            return None
        for pattern, label in compiled:
            if pattern.search(v):
                return label
        return None

    return match


match_country = build_matcher(COUNTRY_ALIASES)
match_country_or_demonym = build_matcher(COUNTRY_ALIASES, DEMONYMS)
match_region = build_matcher(REGION_ALIASES)

# --------------------------------------------------------------------------
# Field classification
# --------------------------------------------------------------------------
def first_line(label):
    for line in (label or "").split("\n"):
        line = line.strip()
        if line:
            return line.rstrip(".").strip()
    return ""


def field_kind(label):
    """Return 'country', 'region', 'regulator', 'bank' or None."""
    l = normalize(first_line(label))
    if "bank" in l:
        return "bank"
    if (l.endswith("| country") or l.startswith("country")
            or l.startswith(("which country", "in which country", "what country"))):
        return "country"
    if (l.endswith(("| province/state", "| city")) or l in ("province/state/territory", "city/town")):
        return "region"
    if l.startswith(("what is the name of the regulatory body", "what is the name of the nursing regulatory")):
        return "regulator"
    return None


def category(task_name, label):
    t, l = normalize(task_name), normalize(first_line(label))
    if "bank" in l or "payment" in t:
        return "issuing_bank"
    if "education" in t or "educ" in l or "school" in l:
        return "education"
    if "registration" in t or "licens" in l or "regulatory" in l:
        return "registration"
    if "examination" in t or "exam" in l:
        return "examination"
    if "employment" in t:
        return "employment"
    if "address" in l:
        return "residential"
    return "other"


def scalar_values(response):
    """Yield plain values from a response, skipping file uploads."""
    if isinstance(response, dict):
        for k in ("value", "other"):
            yield response.get(k)
    elif isinstance(response, list):
        for item in response:
            if isinstance(item, dict) and "filename" in item:
                continue
            yield item
    else:
        yield response


def iter_fields(task):
    data = task.get("data")
    if not isinstance(data, dict):
        return
    for key, field in data.items():
        if isinstance(field, dict) and "label" in field:
            yield key, field.get("label"), field.get("response")


# --------------------------------------------------------------------------
# Dropdown index maps
# --------------------------------------------------------------------------
class ChoiceMaps:
    def __init__(self, cfg: Optional[Dict[str, Any]]):
        self.fields: Dict[str, Dict[str, str]] = {}
        cfg = cfg or {}
        lists = cfg.get("lists", {})
        for key, spec in cfg.get("fields", {}).items():
            if isinstance(spec, str):
                self.fields[key] = {str(i): name for i, name in enumerate(lists.get(spec, []))}
            elif isinstance(spec, dict):
                self.fields[key] = {str(k): v for k, v in spec.items()}

    def lookup(self, field_key, index):
        return self.fields.get(field_key, {}).get(str(index))


def load_choice_maps(override: Optional[Any] = None) -> ChoiceMaps:
    """Use CHOICE_MAPS from this script, with an optional parameter merged over it."""
    if isinstance(override, str) and override.strip():
        override = json.loads(override)
    cfg = {
        "lists":  dict(CHOICE_MAPS.get("lists", {})),
        "fields": dict(CHOICE_MAPS.get("fields", {})),
    }
    if isinstance(override, dict):
        cfg["lists"].update(override.get("lists", {}))
        cfg["fields"].update(override.get("fields", {}))
    return ChoiceMaps(cfg)


# --------------------------------------------------------------------------
# SM Apply helpers (same auth pattern as the credential report script)
# --------------------------------------------------------------------------
_TOKEN: Optional[str] = None
_TOKEN_LOCK = threading.Lock()


def _smapply_headers() -> Dict[str, str]:
    global _TOKEN
    if _TOKEN is None:
        with _TOKEN_LOCK:
            if _TOKEN is None:
                _TOKEN = wmill.get_variable(SMAPPLY_TOKEN_VAR)
                if not _TOKEN:
                    raise RuntimeError(f"Missing Windmill variable: {SMAPPLY_TOKEN_VAR}")
    return {"Authorization": f"Bearer {_TOKEN}", "Accept": "application/json"}


def _smapply_get(url: str, params: Optional[dict] = None, retries: int = 6) -> Dict[str, Any]:
    if url.startswith("/"):
        url = SMAPPLY_BASE + url
    for attempt in range(retries):
        resp = requests.get(url, headers=_smapply_headers(), params=params, timeout=120)
        if resp.status_code == 429 or resp.status_code >= 500:
            time.sleep(int(resp.headers.get("Retry-After", 2 ** attempt)))
            continue
        if resp.status_code == 401:
            raise RuntimeError(
                f"SMApply returned 401 for {url}. The token in {SMAPPLY_TOKEN_VAR} "
                "may have expired."
            )
        if not resp.ok:
            raise RuntimeError(f"SMApply GET failed {resp.status_code} {url}: {resp.text[:2000]}")
        return resp.json()
    raise RuntimeError(f"SMApply GET gave up after {retries} attempts: {url}")


def _paginate(path: str, params: Optional[dict] = None):
    """Follow DRF pagination via `next`, as in fetch_all_reviews."""
    url = path
    while url:
        data = _smapply_get(url, params=params)
        yield from data.get("results", [])
        url, params = data.get("next"), None   # `next` already carries the query


def fetch_applications(program_id: int, limit: int = 0) -> List[Dict[str, Any]]:
    apps = []
    for app in _paginate("/applications/", {"program": program_id}):
        apps.append(app)
        if limit and len(apps) >= limit:
            break
    return apps


def fetch_tasks(application_id: int) -> List[Dict[str, Any]]:
    return list(_paginate(f"/applications/{application_id}/tasks/"))


def fetch_task_choice_maps(task_ids):
    """Read each task's settings once.
    Returns (country_maps, warnings, extra) where:
      country_maps = {field_key: {"index": "Country"}} for country dropdowns
      extra = {"choices": {field_key: {"index": "Label"}} for every choice field,
               "names":   {field_key: full question text}}
    Failures fall back to the static CHOICE_MAPS."""
    maps: Dict[str, Dict[str, str]] = {}
    extra: Dict[str, Dict[str, Any]] = {"choices": {}, "names": {}}
    warnings: List[Dict[str, Any]] = []
    for tid in sorted(task_ids):
        try:
            settings = _smapply_get(TASK_SETTINGS_PATH.format(task_id=tid), retries=3)
        except Exception as exc:
            warnings.append({"task_id": tid, "warning": f"settings not loaded: {str(exc)[:200]}"})
            continue
        for key, field in (settings.get("fields") or {}).items():
            if not isinstance(field, dict):
                continue
            extra["names"][key] = field.get("name") or ""
            choices = field.get("choices")
            if not choices:
                continue
            idx_map = {str(c["index"]): c.get("label", "") for c in choices if "index" in c}
            extra["choices"][key] = idx_map
            if field_kind(field.get("name")) == "country":
                maps[key] = idx_map
    return maps, warnings, extra


# --------------------------------------------------------------------------
# Screening
# --------------------------------------------------------------------------
def screen_application(app_id, tasks, choice_maps):
    hits, unmapped = [], []
    for task in tasks:
        task_name = task.get("name", "")
        for key, label, response in iter_fields(task):
            kind = field_kind(label)
            if not kind:
                continue
            cat = category(task_name, label)
            for value in scalar_values(response):
                if value is None or value == "":
                    continue
                matched, match_type = None, None
                if kind == "country":
                    if isinstance(value, bool):
                        continue
                    if isinstance(value, int):
                        if value < 0:
                            continue
                        name = choice_maps.lookup(key, value)
                        if name is None:
                            unmapped.append({"application_id": app_id, "task": task_name,
                                             "field_key": key, "field_label": first_line(label)[:90],
                                             "index": value})
                            continue
                        review = REVIEW_COUNTRIES.get(normalize(name))
                        if review:
                            hits.append({"application_id": app_id, "category": cat,
                                         "matched_restriction": review, "match_type": "review_dropdown",
                                         "task": task_name, "field_key": key,
                                         "field_label": first_line(label)[:90]})
                            continue
                        matched, match_type = match_country(name), "country_dropdown"
                    else:
                        matched, match_type = match_country(value), "country_text"
                elif kind == "region" and isinstance(value, str):
                    matched, match_type = match_region(value), "region_text"
                elif kind in ("regulator", "bank") and isinstance(value, str):
                    matched, match_type = match_country_or_demonym(value), f"{kind}_name"
                if matched:
                    hits.append({"application_id": app_id, "category": cat,
                                 "matched_restriction": matched, "match_type": match_type,
                                 "task": task_name, "field_key": key,
                                 "field_label": first_line(label)[:90]})
    return hits, unmapped


def _discover(app_tasks: Dict[int, List[Dict[str, Any]]]) -> List[Dict[str, Any]]:
    seen = defaultdict(lambda: {"task": "", "label": "", "values": defaultdict(int)})
    for tasks in app_tasks.values():
        for task in tasks:
            for key, label, response in iter_fields(task):
                if field_kind(label) != "country":
                    continue
                entry = seen[key]
                entry["task"], entry["label"] = task.get("name", ""), first_line(label)[:90]
                for v in scalar_values(response):
                    if isinstance(v, int) and not isinstance(v, bool) and v >= 0:
                        entry["values"][str(v)] += 1
                    elif isinstance(v, str) and v:
                        entry["values"]["(free text)"] += 1
    return [
        {"field_key": k, "task": e["task"], "field_label": e["label"],
         "values_seen": dict(sorted(e["values"].items(), key=lambda kv: kv[0]))}
        for k, e in sorted(seen.items(), key=lambda kv: (kv[1]["task"], kv[0]))
    ]


# --------------------------------------------------------------------------
# Applicant profile (contact, countries, application state)
# Built ONLY for flagged and review applications.
# --------------------------------------------------------------------------
# Personal Information (task 2053922) field keys, from its task settings
PI_FIELDS = {
    "first_name":        "3XWPTimXqR",
    "middle_name":       "CpyEDlPTUL",
    "last_name":         "Oo8AH3TITd",
    "email":             "pbwrR7HpWH",
    "phone":             "IqGLTd9NHd",
    "alt_phone":         "ElX5xth1Sp",
    "current_line1":     "odQBnoDDlJ_0",
    "current_line2":     "odQBnoDDlJ_1",
    "current_city":      "odQBnoDDlJ_2",
    "current_province":  "odQBnoDDlJ_3",
    "current_postal":    "odQBnoDDlJ_4",
    "current_country":   "odQBnoDDlJ_5",
    "mailing_same":      "e9LnXaQVCM",     # 0 = Yes, 1 = No
    "mailing_line1":     "Esldo5mQWc_0",
    "mailing_line2":     "Esldo5mQWc_1",
    "mailing_city":      "Esldo5mQWc_2",
    "mailing_province":  "Esldo5mQWc_3",
    "mailing_postal":    "Esldo5mQWc_4",
    "mailing_country":   "Esldo5mQWc_5",
}
DROPDOWN_KEYS = {"current_country", "mailing_country"}

# Task status codes seen in this program (3 = complete, as in the report script)
TASK_STATUS_LABELS = {1: "Not started", 2: "In progress", 3: "Complete"}

# Application-level keys to pass through if SM Apply returns them
APP_STATE_KEYS = [
    "status", "current_stage", "stage", "decision", "archived",
    "created_at", "submitted_at", "completed_at", "updated_at", "tags",
]


_TAG_RE = re.compile(r"<[^>]+>")


def _clean_text(text: str) -> str:
    """Strip HTML list/line markup SM Apply stores in custom fields, and trim."""
    if "<" in text and ">" in text:
        items = re.findall(r"<li[^>]*>(.*?)</li>", text, flags=re.I | re.S)
        parts = items if items else [text]
        cleaned = []
        for part in parts:
            part = re.sub(r"<br\s*/?>", " ", part, flags=re.I)
            part = _TAG_RE.sub("", part)
            part = " ".join(part.split())
            if part:
                cleaned.append(part)
        text = "; ".join(cleaned)
    text = text.strip()
    # Values like ", , , , ," carry no information
    if text and not text.replace(",", "").strip():
        return ""
    return text


def _plain(v: Any) -> str:
    """Flatten SM Apply values (dicts, lists, HTML) into readable text."""
    if v is None:
        return ""
    if isinstance(v, dict):
        for k in ("title", "name", "label", "value"):
            if v.get(k) not in (None, ""):
                return _clean_text(str(v[k]))
        return json.dumps(v, default=str)
    if isinstance(v, list):
        return "; ".join(x for x in (_plain(i) for i in v) if x)
    return _clean_text(str(v))


PLACEHOLDER_VALUES = {"", "-", "--", "n/a", "na", "none", "null"}


def parse_custom_fields(app: Dict[str, Any]) -> Dict[str, str]:
    """custom_fields list -> {field name: value}."""
    out: Dict[str, str] = {}
    for cf in app.get("custom_fields") or []:
        if not isinstance(cf, dict):
            continue
        name = cf.get("name") or cf.get("label") or str(cf.get("id", ""))
        value = cf.get("value", cf.get("response", cf.get("data")))
        out[name] = _plain(value)
    return out


def parse_notes(app: Dict[str, Any]) -> str:
    """notes list -> 'YYYY-MM-DD [private]: content' lines, newest first."""
    notes = [n for n in (app.get("notes") or []) if isinstance(n, dict)]
    notes.sort(key=lambda n: n.get("created_at") or "", reverse=True)
    lines = []
    for n in notes:
        date = (n.get("created_at") or "")[:10]
        tag = "private" if n.get("is_private") else "public"
        lines.append(f"{date} [{tag}]: {(n.get('content') or '').strip()}")
    return " || ".join(lines)


def payment_summary(tasks: List[Dict[str, Any]], custom: Dict[str, str],
                    labels: str = "") -> Dict[str, str]:
    """Paid status, from strongest evidence to weakest:
    1. an 'Expedited Service Payment...' task marked Complete
    2. a 'Payment date' custom field value
    3. an 'Invoiced' label (older applications without payment tasks)"""
    pay_tasks = [t for t in tasks if normalize(t.get("name")).startswith("expedited service payment")]
    done = [t.get("name", "") for t in pay_tasks if t.get("status") == 3]
    payment_date = custom.get("Payment date", "")
    invoiced = "invoiced" in normalize(labels)
    if done:
        paid, basis = "Yes", "; ".join(done)
    elif payment_date:
        paid, basis = "Yes", "Payment date field"
    elif invoiced:
        paid, basis = "Invoiced", "Invoiced label (no completed payment task)"
    elif not pay_tasks:
        paid, basis = "Unknown", "No payment task on application"
    else:
        paid, basis = "No", "Payment tasks not completed"
    return {"paid": paid, "paid_via": basis, "payment_date": payment_date}


def build_contact(tasks: List[Dict[str, Any]], maps: "ChoiceMaps") -> Dict[str, str]:
    pi = next((t for t in tasks if normalize(t.get("name")) == "personal information"), None)
    data = (pi or {}).get("data") or {}
    out: Dict[str, str] = {}
    for name, key in PI_FIELDS.items():
        resp = (data.get(key) or {}).get("response")
        if name in DROPDOWN_KEYS and isinstance(resp, int) and not isinstance(resp, bool):
            out[name] = (maps.lookup(key, resp) or "") if resp >= 0 else ""
        elif name == "mailing_same":
            out[name] = {0: "Yes", 1: "No"}.get(resp, _plain(resp))
        else:
            out[name] = _plain(resp)
    if out.get("mailing_same") == "Yes":
        for part in ("line1", "line2", "city", "province", "postal", "country"):
            out[f"mailing_{part}"] = out.get(f"current_{part}", "")
    return out


def build_countries(tasks: List[Dict[str, Any]], maps: "ChoiceMaps") -> Dict[str, List[str]]:
    """Every country answer, translated, grouped by category."""
    found: Dict[str, set] = defaultdict(set)
    for task in tasks:
        for key, label, response in iter_fields(task):
            if field_kind(label) != "country":
                continue
            for v in scalar_values(response):
                if isinstance(v, int) and not isinstance(v, bool):
                    name = maps.lookup(key, v) if v >= 0 else None
                elif isinstance(v, str):
                    name = v.strip()
                else:
                    name = None
                if name and normalize(name) not in PLACEHOLDER_VALUES:
                    bucket = found[category(task.get("name", ""), label)]
                    key_norm = re.sub(r"^the ", "", normalize(name))
                    if not any(re.sub(r"^the ", "", normalize(x)) == key_norm for x in bucket):
                        bucket.add(name)
    return {k: sorted(v) for k, v in sorted(found.items())}


def build_state(app: Dict[str, Any], tasks: List[Dict[str, Any]]) -> Dict[str, Any]:
    state: Dict[str, Any] = {k: _plain(app.get(k)) for k in APP_STATE_KEYS if k in app}
    # Any payment-related application fields (e.g. paid, payment_status)
    for k, v in app.items():
        if ("pay" in k.lower() or "paid" in k.lower()) and k not in state:
            state[k] = _plain(v)
    # Payment-related tasks (e.g. "Terms and Conditions - Payment Confirmation")
    for t in tasks:
        if "pay" in normalize(t.get("name")):
            state[f"task: {t.get('name')}"] = TASK_STATUS_LABELS.get(t.get("status"), _plain(t.get("status")))
    state.update(payment_summary(tasks, parse_custom_fields(app), _plain(app.get("labels"))))
    state["task_statuses"] = {
        t.get("name", str(t.get("id"))): TASK_STATUS_LABELS.get(t.get("status"), _plain(t.get("status")))
        for t in tasks
    }
    return state


def _choice_text(key: str, resp: Any, extra: Dict[str, Any]) -> str:
    """Translate a choice answer (int or list of ints) using task settings."""
    labels = extra.get("choices", {}).get(key, {})
    vals = resp if isinstance(resp, list) else [resp]
    out = []
    for v in vals:
        if isinstance(v, bool) or v in (None, "", -1):
            continue
        if isinstance(v, int):
            out.append(labels.get(str(v), f"[option {v}]"))
        elif isinstance(v, dict):
            out.append(_plain(v))
        else:
            out.append(str(v))
    return "; ".join(x for x in out if x)


def build_application_info(app: Dict[str, Any], tasks: List[Dict[str, Any]],
                           extra: Dict[str, Any]) -> Dict[str, Any]:
    """Application-level fields plus the nursing discipline / regulatory body selections."""
    skip = set(APP_STATE_KEYS) | {"applicant", "custom_fields", "notes", "state"}
    info: Dict[str, Any] = {k: _plain(v) for k, v in app.items() if k not in skip}
    applicant = app.get("applicant") or {}
    if isinstance(applicant, dict):
        info["applicant_id"] = _plain(applicant.get("id"))
        info["applicant_name"] = " ".join(
            p for p in [_plain(applicant.get("first_name")), _plain(applicant.get("last_name"))] if p)
        info["applicant_email"] = _plain(applicant.get("email"))
    info["custom_fields"] = parse_custom_fields(app)
    info["notes"] = parse_notes(app)

    selections: Dict[str, str] = {}
    for task in tasks:
        if "discipline" not in normalize(task.get("name")):
            continue
        for key, label, response in iter_fields(task):
            answer = _choice_text(key, response, extra)
            if not answer:
                continue
            question = first_line(extra.get("names", {}).get(key) or label)[:120].rstrip(": ").strip()
            if question in selections:
                question = f"{question} ({key})"
            selections[question] = answer
    info["nursing_discipline"] = selections
    return info


def build_csv(rows: List[Dict[str, Any]]) -> str:
    """One row per flagged/review application, for the compliance team."""
    contact_cols = list(PI_FIELDS)
    custom_names: List[str] = []
    for r in rows:
        for k, v in (r.get("application_info", {}).get("custom_fields") or {}).items():
            if v and k not in custom_names:
                custom_names.append(k)
    custom_names.sort(key=lambda n: n.lower())
    cols = (["result", "application_id", "reference_id", "applicant_id", "program",
             "restrictions", "match_categories",
             "application_status", "current_stage", "paid", "paid_via", "payment_date",
             "labels", "nursing_discipline", "created_at", "last_submitted_at", "updated_at"]
            + contact_cols
            + ["countries_residential", "countries_education", "countries_registration",
               "countries_examination", "countries_employment", "countries_other"]
            + ["notes"]
            + [f"custom: {n}" for n in custom_names])
    buf = io.StringIO()
    w = csv.DictWriter(buf, fieldnames=cols, extrasaction="ignore")
    w.writeheader()
    for r in rows:
        c, ctry, st = r.get("contact", {}), r.get("countries", {}), r.get("state", {})
        ai = r.get("application_info", {})
        custom = ai.get("custom_fields") or {}
        line = {
            "result": r["result"],
            "application_id": r["application_id"],
            "reference_id": ai.get("reference_id") or ai.get("title", ""),
            "applicant_id": ai.get("applicant_id", ""),
            "program": ai.get("program", ""),
            "restrictions": "; ".join(r["restrictions"]),
            "match_categories": "; ".join(r["categories"]),
            "application_status": st.get("status", ""),
            "current_stage": st.get("current_stage", ""),
            "paid": st.get("paid", ""),
            "paid_via": st.get("paid_via", ""),
            "payment_date": st.get("payment_date", ""),
            "labels": ai.get("labels", ""),
            "nursing_discipline": " | ".join(f"{q}: {a}" for q, a in
                                             (ai.get("nursing_discipline") or {}).items()),
            "created_at": st.get("created_at", ""),
            "last_submitted_at": ai.get("last_submitted_at", ""),
            "updated_at": st.get("updated_at", ""),
            **{k: c.get(k, "") for k in contact_cols},
            **{f"countries_{k}": "; ".join(ctry.get(k, [])) for k in
               ("residential", "education", "registration", "examination", "employment", "other")},
            "notes": ai.get("notes", ""),
            **{f"custom: {n}": custom.get(n, "") for n in custom_names},
        }
        w.writerow(line)
    return buf.getvalue()


# --------------------------------------------------------------------------
# SharePoint upload (Microsoft Graph, same auth as the credential report script)
# --------------------------------------------------------------------------
TENANT_ID_VAR     = "u/epham2/AZURE_TENANT_ID"
CLIENT_ID_VAR     = "u/epham2/SHAREPOINT_CLIENT_ID"
CLIENT_SECRET_VAR = "u/epham2/SHAREPOINT_CLIENT_SECRET"

GRAPH_BASE            = "https://graph.microsoft.com/v1.0"
SHAREPOINT_HOSTNAME   = "nnasca.sharepoint.com"
SITE_PATH             = "/sites/NNAS-Portal-Sharepoint"
SITE_WEB_URL          = "https://nnasca.sharepoint.com/sites/NNAS-Portal-Sharepoint"
DOCUMENT_LIBRARY_NAME = "Applications"

# Screening results folder, same path style as the credential report script.
# The file contains applicant contact details, so break permission inheritance
# on this folder in SharePoint and limit it to the people who need it.
SCREEN_FOLDER    = "Applications/Sanctions Screening"
SCREEN_FILE_TYPE = "Report"      # FileType column value, as the report script sets

_GRAPH_TOKEN: Optional[str] = None
_SP_SITE_ID: Optional[str] = None
_SP_DRIVE_ID: Optional[str] = None
_SP_LIST_ID: Optional[str] = None


def _get_graph_token() -> str:
    global _GRAPH_TOKEN
    if _GRAPH_TOKEN:
        return _GRAPH_TOKEN
    tenant = (wmill.get_variable(TENANT_ID_VAR) or "").strip()
    client_id = (wmill.get_variable(CLIENT_ID_VAR) or "").strip()
    secret = (wmill.get_variable(CLIENT_SECRET_VAR) or "").strip()
    if not tenant:
        raise RuntimeError(f"Missing Windmill variable: {TENANT_ID_VAR}")
    if not client_id:
        raise RuntimeError(f"Missing Windmill variable: {CLIENT_ID_VAR}")
    if not secret:
        raise RuntimeError(f"Missing Windmill variable: {CLIENT_SECRET_VAR}")
    resp = requests.post(
        f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token",
        data={"client_id": client_id, "client_secret": secret,
              "scope": "https://graph.microsoft.com/.default",
              "grant_type": "client_credentials"},
        timeout=60,
    )
    if resp.status_code != 200:
        raise RuntimeError(f"Microsoft Graph token request failed {resp.status_code}: {resp.text[:1000]}")
    _GRAPH_TOKEN = resp.json()["access_token"]
    return _GRAPH_TOKEN


def _graph_request(method: str, endpoint: str, expected_statuses=(200,), **kwargs):
    headers = kwargs.pop("headers", {})
    headers["Authorization"] = f"Bearer {_get_graph_token()}"
    resp = requests.request(method, f"{GRAPH_BASE}{endpoint}", headers=headers, timeout=180, **kwargs)
    if resp.status_code not in expected_statuses:
        raise RuntimeError(f"Microsoft Graph {method} {endpoint} failed {resp.status_code}: {resp.text[:2000]}")
    return resp.json() if resp.content else None


def _library_relative_path(path: str) -> str:
    """Accepts 'sites/NNAS-Portal-Sharepoint/Applications/x', 'Applications/x' or 'x'."""
    p = (path or "").strip().strip("/")
    if p.lower().startswith("sites/nnas-portal-sharepoint/"):
        p = p[len("sites/NNAS-Portal-Sharepoint/"):]
    if p.lower().startswith("applications/"):
        p = p[len("Applications/"):]
    return p.strip("/")


def _encode_graph_path(path: str) -> str:
    return quote(path.strip("/"), safe="/")


def _init_sharepoint_graph() -> None:
    global _SP_SITE_ID, _SP_DRIVE_ID, _SP_LIST_ID
    if _SP_SITE_ID and _SP_DRIVE_ID and _SP_LIST_ID:
        return
    site = _graph_request("GET", f"/sites/{SHAREPOINT_HOSTNAME}:{SITE_PATH}")
    _SP_SITE_ID = site["id"]
    drives = _graph_request("GET", f"/sites/{_SP_SITE_ID}/drives")["value"]
    drive = next((d for d in drives if d.get("name") == DOCUMENT_LIBRARY_NAME), None)
    if not drive:
        raise RuntimeError(f"Document library '{DOCUMENT_LIBRARY_NAME}' not found. "
                           f"Available: {[d.get('name') for d in drives]}")
    _SP_DRIVE_ID = drive["id"]
    lists = _graph_request("GET", f"/sites/{_SP_SITE_ID}/lists?$select=id,displayName")["value"]
    sp_list = next((l for l in lists if l.get("displayName") == DOCUMENT_LIBRARY_NAME), None)
    if not sp_list:
        raise RuntimeError(f"SharePoint list '{DOCUMENT_LIBRARY_NAME}' not found.")
    _SP_LIST_ID = sp_list["id"]


def _ensure_sp_folder_tree(folder_path: str) -> None:
    _init_sharepoint_graph()
    current = ""
    for part in [p for p in _library_relative_path(folder_path).split("/") if p]:
        target = f"{current}/{part}".strip("/")
        resp = requests.get(f"{GRAPH_BASE}/drives/{_SP_DRIVE_ID}/root:/{_encode_graph_path(target)}",
                            headers={"Authorization": f"Bearer {_get_graph_token()}"}, timeout=60)
        if resp.status_code == 404:
            parent = (f"/drives/{_SP_DRIVE_ID}/root:/{_encode_graph_path(current)}:/children"
                      if current else f"/drives/{_SP_DRIVE_ID}/root/children")
            _graph_request("POST", parent, expected_statuses=(201, 409),
                           headers={"Content-Type": "application/json"},
                           json={"name": part, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"})
        elif resp.status_code != 200:
            raise RuntimeError(f"Graph folder lookup failed {resp.status_code}: {resp.text[:1000]}")
        current = target


def _patch_driveitem_fields(drive_item_id: str, fields: dict) -> None:
    item = _graph_request("GET", f"/drives/{_SP_DRIVE_ID}/items/{drive_item_id}?$expand=listItem")
    list_item = item.get("listItem") or {}
    if not list_item.get("id"):
        raise RuntimeError(f"Drive item {drive_item_id} has no listItem information.")
    _graph_request("PATCH", f"/sites/{_SP_SITE_ID}/lists/{_SP_LIST_ID}/items/{list_item['id']}/fields",
                   expected_statuses=(200, 204), headers={"Content-Type": "application/json"},
                   json=fields)


def upload_csv_to_sharepoint(csv_text: str, filename: str,
                             folder_path: str = SCREEN_FOLDER,
                             file_type: str = SCREEN_FILE_TYPE) -> str:
    """Upload the CSV (UTF-8 with BOM so Excel opens it cleanly). Returns the file's web URL."""
    _ensure_sp_folder_tree(folder_path)
    upload_path = f"{_library_relative_path(folder_path)}/{filename}".strip("/")
    item = _graph_request("PUT", f"/drives/{_SP_DRIVE_ID}/root:/{_encode_graph_path(upload_path)}:/content",
                          expected_statuses=(200, 201), headers={"Content-Type": "text/csv"},
                          data=csv_text.encode("utf-8-sig"))
    if file_type:
        try:
            _patch_driveitem_fields(item["id"], {"FileType": file_type})
        except Exception:
            pass  # metadata is optional; the file itself is saved
    return item.get("webUrl") or f"{SITE_WEB_URL}/Applications/{_encode_graph_path(upload_path)}"


def main(mode: str = "screen",
         program_id: int = PROGRAM_ID,
         limit: int = 0,
         max_workers: int = 4,
         choice_maps: Optional[dict] = None,
         include_contact: bool = True,
         upload_to_sharepoint: bool = True,
         sharepoint_folder: str = SCREEN_FOLDER):
    """
    mode="discover": list country dropdown fields and index values in use.
    mode="screen":   flag applications linked to a restricted country/region.
    limit:           process only the first N applications (0 = all).
    max_workers:     parallel task fetches; keep low to respect rate limits.
    upload_to_sharepoint: save the full CSV (with contact details) to SharePoint
                     and keep personal information out of the Windmill job result.
    """
    started = time.time()
    apps = fetch_applications(program_id, limit)
    app_meta = {a["id"]: a for a in apps}
    app_ids = list(app_meta)

    app_tasks: Dict[int, List[Dict[str, Any]]] = {}
    errors: List[Dict[str, Any]] = []

    def _load(aid):
        try:
            return aid, fetch_tasks(aid), None
        except Exception as exc:
            return aid, [], str(exc)[:500]

    with ThreadPoolExecutor(max_workers=max(1, max_workers)) as pool:
        for aid, tasks, err in pool.map(_load, app_ids):
            if err:
                errors.append({"application_id": aid, "error": err})
            else:
                app_tasks[aid] = tasks

    if mode == "discover":
        return {
            "mode": "discover",
            "applications_scanned": len(app_tasks),
            "country_fields": _discover(app_tasks),
            "errors": errors,
        }

    maps = load_choice_maps(choice_maps)
    task_ids = {t["id"] for tasks in app_tasks.values() for t in tasks if "id" in t}
    live_maps, settings_warnings, settings_extra = fetch_task_choice_maps(task_ids)
    maps.fields.update(live_maps)          # task settings win over static fallback
    hits_by_app: Dict[int, List[Dict[str, Any]]] = {}
    unmapped: List[Dict[str, Any]] = []
    for aid in app_ids:
        if aid not in app_tasks:
            continue
        hits, um = screen_application(aid, app_tasks[aid], maps)
        if hits:
            hits_by_app[aid] = hits
        unmapped.extend(um)

    flagged, review = [], []
    for aid, hits in hits_by_app.items():
        restricted = [h for h in hits if h["match_type"] != "review_dropdown"]
        target = flagged if restricted else review
        entry = {
            "application_id": aid,
            "result": "flagged" if restricted else "review",
            "restrictions": sorted({h["matched_restriction"] for h in (restricted or hits)}),
            "categories": sorted({h["category"] for h in (restricted or hits)}),
            "countries": build_countries(app_tasks[aid], maps),
            "application_info": build_application_info(app_meta.get(aid, {}), app_tasks[aid],
                                                        settings_extra),
            "state": build_state(app_meta.get(aid, {}), app_tasks[aid]),
            "matches": hits,
        }
        if include_contact:
            entry["contact"] = build_contact(app_tasks[aid], maps)
        target.append(entry)

    csv_text = build_csv(flagged + review)
    sharepoint_file, sharepoint_error = "", ""
    if upload_to_sharepoint:
        stamp = datetime.now(timezone.utc).strftime("%Y-%m-%d_%H%M")
        filename = f"CES_sanctions_screen_{stamp}.csv"
        try:
            sharepoint_file = upload_csv_to_sharepoint(csv_text, filename, sharepoint_folder)
        except Exception as exc:
            sharepoint_error = str(exc)[:1000]

    # Keep personal information out of Windmill job history once the file is safely saved.
    if sharepoint_file:
        for entry in flagged + review:
            entry.pop("contact", None)
            entry.get("application_info", {}).pop("notes", None)
            for k in ("applicant_name", "applicant_email"):
                entry.get("application_info", {}).pop(k, None)

    return {
        "mode": "screen",
        "program_id": program_id,
        "applications_scanned": len(app_tasks),
        "sharepoint_file": sharepoint_file,
        "sharepoint_error": sharepoint_error,
        "flagged_application_ids": sorted(f["application_id"] for f in flagged),
        "review_application_ids": sorted(r["application_id"] for r in review),
        "error_application_ids": sorted(e["application_id"] for e in errors),
        # Full CSV stays in the result only if it could not be saved to SharePoint
        "csv": "" if sharepoint_file else csv_text,
        "application_fields_available": sorted(next(iter(app_meta.values()), {}).keys()),
        "flagged_count": len(flagged),
        "flagged": flagged,
        "review_count": len(review),
        "review": review,
        "unmapped_dropdown_count": len(unmapped),
        "unmapped_dropdowns": unmapped,
        "choice_map_fields_from_task_settings": len(live_maps),
        "choice_map_fields_with_options": sum(1 for m in maps.fields.values() if m),
        "settings_warnings": settings_warnings,
        "choice_map_fields_empty": sorted(k for k, m in maps.fields.items() if not m),
        "errors": errors,
        "elapsed_seconds": round(time.time() - started, 1),
    }