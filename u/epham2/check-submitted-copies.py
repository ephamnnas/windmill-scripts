# py: 3.12
# requirements:
# requests
# wmill
# pypdf
# pdf2image
# pytesseract

"""
READ-ONLY: check the 27 ambiguous registration country fields on the 16
submitted applications against the copy of the application saved to
SharePoint at submission.

For each field there are only two possible answers:
  * untouched:   the stored value read against Countries (Registration),
                 correct if the applicant did not change the field after the
                 Oct 30 switch
  * re-selected: the stored value read against Countries (Custom), correct if
                 they re-picked the country while Countries (Custom) was active

The submitted copy settles it when it was saved BEFORE the switch
(2026-10-30 21:20 UTC): it shows the country the applicant originally chose.
  * copy matches "untouched"   -> no change needed
  * copy matches "re-selected" -> change to that country
  * copy shows something else  -> manual check (applicant chose a different
                                  country after the switch, or text unclear)
A copy saved DURING the affected window shows whatever was displayed at the
time, which is always the Countries (Custom) reading, so it cannot decide
the question; those fields are marked inconclusive.

Text is read with pypdf. If a page has no text layer (a scanned image), OCR
is tried with pytesseract, which needs the tesseract and poppler system
packages on the Windmill worker; if they are missing the field is marked
"needs OCR" rather than failing.

Nothing is changed in SM Apply or SharePoint apart from uploading the result
CSV to Applications/Country list backups.

Run discover=True first with one application: it lists the files found for
that application (names, paths, dates) and shows the text around each
country question, so FILE_PATTERN and the label patterns can be confirmed.
"""

import csv
import io
import json
import re
import unicodedata
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import quote

import requests
import wmill
from pypdf import PdfReader


# =========================================================
# CONFIG
# =========================================================

SMAPPLY_BASE = "https://nnas.smapply.io/api"
SMAPPLY_TOKEN_VAR = "u/epham2/SMAPPLY_ACCESS_TOKEN"

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SHAREPOINT_HOSTNAME = "nnasca.sharepoint.com"
SITE_PATH = "/sites/NNAS-Portal-Sharepoint"
DOCUMENT_LIBRARY_NAME = "Applications"
# The report script's "applications/tm" is relative to the site; inside the
# Applications library it is just "TM".
APPLICATIONS_BASE = "TM"
RESULT_FOLDER = "Country list backups"

# Which files count as the submitted application copy (case-insensitive
# regex on the file name). Confirm with discover mode and adjust.
# Submitted copy: <reference>_<YYYYMMDD>_<HHMMSS>.pdf, e.g.
# NNAS-T-0000000947_20260906_130322.pdf (the ID1, ID2, IENForm and
# DocumentIndex files are excluded by this pattern).
FILE_PATTERN = r"^NNAS-T-\d+_\d{8}_\d{6}\.pdf$"
SKIP_FOLDERS = {"credential reports", "communications", "international credentials"}
SKIP_FOLDER_PREFIXES = ("sftp_export",)

SWITCH_UTC = datetime(2026, 10, 30, 21, 20, tzinfo=timezone.utc)
SWITCH_BACK_UTC = datetime(2026, 11, 5, 16, 20, tzinfo=timezone.utc)

# Question labels in the submitted copy, English and French.
CURRENT_LABEL = re.compile(
    r"(which country are you currently registered|dans quel pays .{0,40}actuellement)", re.I)
PREVIOUS_LABEL = re.compile(
    r"(which country were you previously registered|dans quel pays .{0,40}(anterieurement|precedemment|auparavant))", re.I)

# Fields to check: application, field position (0 = current, 1-6 = previous),
# stored value, and its index in Countries (Registration) if it was chosen
# from Countries (Custom).
FIELDS = [
    {"app": 35589355, "field": "Current Registration", "pos": 0, "stored": 198, "custom_idx": 78},
    {"app": 35590885, "field": "Current Registration", "pos": 0, "stored": 138, "custom_idx": 139},
    {"app": 35604818, "field": "Current Registration", "pos": 0, "stored": 174, "custom_idx": 171},
    {"app": 35625530, "field": "Current Registration", "pos": 0, "stored": 185, "custom_idx": 181},
    {"app": 35650362, "field": "Current Registration", "pos": 0, "stored": 13, "custom_idx": 14},
    {"app": 35671271, "field": "Current Registration", "pos": 0, "stored": 29, "custom_idx": 32},
    {"app": 35671271, "field": "Previous Registration 1", "pos": 1, "stored": 29, "custom_idx": 32},
    {"app": 35716321, "field": "Current Registration", "pos": 0, "stored": 186, "custom_idx": 182},
    {"app": 35853913, "field": "Current Registration", "pos": 0, "stored": 184, "custom_idx": 180},
    {"app": 35864060, "field": "Current Registration", "pos": 0, "stored": 75, "custom_idx": 81},
    {"app": 35879414, "field": "Current Registration", "pos": 0, "stored": 56, "custom_idx": 60},
    {"app": 35879414, "field": "Previous Registration 1", "pos": 1, "stored": 56, "custom_idx": 60},
    {"app": 35879414, "field": "Previous Registration 2", "pos": 2, "stored": 56, "custom_idx": 60},
    {"app": 35879414, "field": "Previous Registration 3", "pos": 3, "stored": 56, "custom_idx": 60},
    {"app": 35879414, "field": "Previous Registration 4", "pos": 4, "stored": 56, "custom_idx": 60},
    {"app": 35908317, "field": "Current Registration", "pos": 0, "stored": 182, "custom_idx": 178},
    {"app": 35908317, "field": "Previous Registration 1", "pos": 1, "stored": 186, "custom_idx": 182},
    {"app": 35908317, "field": "Previous Registration 2", "pos": 2, "stored": 186, "custom_idx": 182},
    {"app": 35937949, "field": "Current Registration", "pos": 0, "stored": 121, "custom_idx": 29},
    {"app": 35945656, "field": "Current Registration", "pos": 0, "stored": 138, "custom_idx": 139},
    {"app": 35945656, "field": "Previous Registration 1", "pos": 1, "stored": 151, "custom_idx": 149},
    {"app": 35945662, "field": "Current Registration", "pos": 0, "stored": 75, "custom_idx": 81},
    {"app": 35945662, "field": "Previous Registration 1", "pos": 1, "stored": 63, "custom_idx": 67},
    {"app": 35945662, "field": "Previous Registration 2", "pos": 2, "stored": 186, "custom_idx": 182},
    {"app": 35945662, "field": "Previous Registration 3", "pos": 3, "stored": 30, "custom_idx": 33},
    {"app": 35962741, "field": "Current Registration", "pos": 0, "stored": 138, "custom_idx": 139},
    {"app": 35970115, "field": "Current Registration", "pos": 0, "stored": 138, "custom_idx": 139},
]

# Every name each Countries (Registration) index can appear as:
# Registration name, French name, and the Countries (Custom) name.
NAMES = {
    "0": ["Afghanistan", "Afghanistan"],
    "1": ["Albania", "Albanie"],
    "2": ["Algeria", "Algérie"],
    "3": ["Andorra", "Andorre"],
    "4": ["Angola", "Angola"],
    "5": ["Antarctica", "Antarctique"],
    "6": ["Antigua and Barbuda", "Antigua-et-Barbuda", "Antigua & Deps"],
    "7": ["Argentina", "Argentine"],
    "8": ["Armenia", "Arménie"],
    "9": ["Australia", "Australie"],
    "10": ["Austria", "Autriche"],
    "11": ["Azerbaijan", "Azerbaïdjan"],
    "12": ["Bahamas", "Bahamas"],
    "13": ["Bahrain", "Bahreïn"],
    "14": ["Bangladesh", "Bangladesh"],
    "15": ["Barbados", "Barbade"],
    "16": ["Belarus", "Bélarus"],
    "17": ["Belgium", "Belgique"],
    "18": ["Belize", "Belize"],
    "19": ["Benin", "Bénin"],
    "20": ["Bermuda", "Bermudes"],
    "21": ["Bhutan", "Bhoutan"],
    "22": ["Bolivia", "Bolivie"],
    "23": ["Bosnia and Herzegovina", "Bosnie-Herzégovine", "Bosnia Herzegovina"],
    "24": ["Botswana", "Botswana"],
    "25": ["Brazil", "Brésil"],
    "26": ["Brunei", "Brunéi Darussalam"],
    "27": ["Bulgaria", "Bulgarie"],
    "28": ["Burkina Faso", "Burkina Faso", "Burkina"],
    "29": ["Burma", "Myanmar", "Myanmar, {Burma}"],
    "30": ["Burundi", "Burundi"],
    "31": ["Cambodia", "Cambodge"],
    "32": ["Cameroon", "Cameroun"],
    "33": ["Canada", "Canada"],
    "34": ["Cape Verde", "Cabo Verde"],
    "35": ["Central African Republic", "République centrafricaine", "Central African Rep"],
    "36": ["Chad", "Tchad"],
    "37": ["Chile", "Chili"],
    "38": ["China", "Chine"],
    "39": ["Colombia", "Colombie"],
    "40": ["Comoros", "Comores"],
    "41": ["Congo, Democratic Republic", "Congo (République démocratique)", "Congo {Democratic Rep}"],
    "42": ["Congo, Republic of the", "Congo (République du)", "Congo"],
    "43": ["Costa Rica", "Costa Rica"],
    "44": ["Cote d'Ivoire", "Côte d'Ivoire", "Ivory Coast"],
    "45": ["Croatia", "Croatie"],
    "46": ["Cuba", "Cuba"],
    "47": ["Cyprus", "Chypre"],
    "48": ["Czech Republic", "République tchèque"],
    "49": ["Denmark", "Danemark"],
    "50": ["Djibouti", "Djibouti"],
    "51": ["Dominica", "Dominique"],
    "52": ["Dominican Republic", "République dominicaine"],
    "53": ["East Timor", "Timor-Leste"],
    "54": ["Ecuador", "Équateur"],
    "55": ["Egypt", "Égypte"],
    "56": ["El Salvador", "El Salvador"],
    "57": ["Equatorial Guinea", "Guinée équatoriale"],
    "58": ["Eritrea", "Érythrée"],
    "59": ["Estonia", "Estonie"],
    "60": ["Ethiopia", "Éthiopie"],
    "61": ["Fiji", "Fidji"],
    "62": ["Finland", "Finlande"],
    "63": ["France", "France"],
    "64": ["Gabon", "Gabon"],
    "65": ["Gambia", "Gambie"],
    "66": ["Georgia", "Géorgie"],
    "67": ["Germany", "Allemagne"],
    "68": ["Ghana", "Ghana"],
    "69": ["Greece", "Grèce"],
    "70": ["Greenland", "Groenland"],
    "71": ["Grenada", "Grenade"],
    "72": ["Guatemala", "Guatemala"],
    "73": ["Guinea", "Guinée"],
    "74": ["Guinea-Bissau", "Guinée-Bissau"],
    "75": ["Guyana", "Guyana"],
    "76": ["Haiti", "Haïti"],
    "77": ["Honduras", "Honduras"],
    "78": ["Hong Kong", "Hong Kong"],
    "79": ["Hungary", "Hongrie"],
    "80": ["Iceland", "Islande"],
    "81": ["India", "Inde"],
    "82": ["Indonesia", "Indonésie"],
    "83": ["Iran", "Iran"],
    "84": ["Iraq", "Irak"],
    "85": ["Ireland", "Irlande", "Ireland {Republic}"],
    "86": ["Israel", "Israël"],
    "87": ["Italy", "Italie"],
    "88": ["Jamaica", "Jamaïque"],
    "89": ["Japan", "Japon"],
    "90": ["Jordan", "Jordanie"],
    "91": ["Kazakhstan", "Kazakhstan"],
    "92": ["Kenya", "Kenya"],
    "93": ["Kiribati", "Kiribati"],
    "94": ["Korea, North", "Corée du Nord", "Korea North"],
    "95": ["Korea, South", "Corée du Sud", "Korea South"],
    "96": ["Kuwait", "Koweït"],
    "97": ["Kyrgyzstan", "Kirghizistan"],
    "98": ["Laos", "Laos"],
    "99": ["Latvia", "Lettonie"],
    "100": ["Lebanon", "Liban"],
    "101": ["Lesotho", "Lesotho"],
    "102": ["Liberia", "Libéria"],
    "103": ["Libya", "Libye"],
    "104": ["Liechtenstein", "Liechtenstein"],
    "105": ["Lithuania", "Lituanie"],
    "106": ["Luxembourg", "Luxembourg"],
    "107": ["Macedonia", "Macédoine du Nord"],
    "108": ["Madagascar", "Madagascar"],
    "109": ["Malawi", "Malawi"],
    "110": ["Malaysia", "Malaisie"],
    "111": ["Maldives", "Maldives"],
    "112": ["Mali", "Mali"],
    "113": ["Malta", "Malte"],
    "114": ["Marshall Islands", "Îles Marshall"],
    "115": ["Mauritania", "Mauritanie"],
    "116": ["Mauritius", "Maurice"],
    "117": ["Mexico", "Mexique"],
    "118": ["Micronesia", "Micronésie"],
    "119": ["Moldova", "Moldavie"],
    "120": ["Mongolia", "Mongolie"],
    "121": ["Morocco", "Maroc"],
    "122": ["Monaco", "Monaco"],
    "123": ["Mozambique", "Mozambique"],
    "124": ["Namibia", "Namibie"],
    "125": ["Nauru", "Nauru"],
    "126": ["Nepal", "Népal"],
    "127": ["Netherlands", "Pays-Bas"],
    "128": ["New Zealand", "Nouvelle-Zélande"],
    "129": ["Nicaragua", "Nicaragua"],
    "130": ["Niger", "Niger"],
    "131": ["Nigeria", "Nigéria"],
    "132": ["Norway", "Norvège"],
    "133": ["Oman", "Oman"],
    "134": ["Pakistan", "Pakistan"],
    "135": ["Panama", "Panama"],
    "136": ["Papua New Guinea", "Papouasie-Nouvelle-Guinée"],
    "137": ["Paraguay", "Paraguay"],
    "138": ["Peru", "Pérou"],
    "139": ["Philippines", "Philippines"],
    "140": ["Poland", "Pologne"],
    "141": ["Portugal", "Portugal"],
    "142": ["Qatar", "Qatar"],
    "143": ["Romania", "Roumanie"],
    "144": ["Russia", "Russie", "Russian Federation"],
    "145": ["Rwanda", "Rwanda"],
    "146": ["Samoa", "Samoa"],
    "147": ["San Marino", "Saint-Marin"],
    "148": ["Sao Tome", "Sao Tomé-et-Principe", "Sao Tome & Principe"],
    "149": ["Saudi Arabia", "Arabie saoudite"],
    "150": ["Senegal", "Sénégal"],
    "151": ["Serbia and Montenegro", "Serbie-et-Monténégro"],
    "152": ["Seychelles", "Seychelles"],
    "153": ["Sierra Leone", "Sierra Leone"],
    "154": ["Singapore", "Singapour"],
    "155": ["Slovakia", "Slovaquie"],
    "156": ["Slovenia", "Slovénie"],
    "157": ["Solomon Islands", "Îles Salomon"],
    "158": ["Somalia", "Somalie"],
    "159": ["South Africa", "Afrique du Sud"],
    "160": ["Spain", "Espagne"],
    "161": ["Sri Lanka", "Sri Lanka"],
    "162": ["Sudan", "Soudan"],
    "163": ["Suriname", "Suriname"],
    "164": ["Swaziland", "Eswatini"],
    "165": ["Sweden", "Suède"],
    "166": ["Switzerland", "Suisse"],
    "167": ["Syria", "Syrie"],
    "168": ["Taiwan", "Taïwan"],
    "169": ["Tajikistan", "Tadjikistan"],
    "170": ["Tanzania", "Tanzanie"],
    "171": ["Thailand", "Thaïlande"],
    "172": ["Togo", "Togo"],
    "173": ["Tonga", "Tonga"],
    "174": ["Trinidad and Tobago", "Trinité-et-Tobago", "Trinidad & Tobago"],
    "175": ["Tunisia", "Tunisie"],
    "176": ["Turkey", "Türkiye"],
    "177": ["Turkmenistan", "Turkménistan"],
    "178": ["Uganda", "Ouganda"],
    "179": ["Ukraine", "Ukraine"],
    "180": ["United Arab Emirates", "Émirats arabes unis"],
    "181": ["United Kingdom", "Royaume-Uni"],
    "182": ["United States", "États-Unis"],
    "183": ["Uruguay", "Uruguay"],
    "184": ["Uzbekistan", "Ouzbékistan"],
    "185": ["Vanuatu", "Vanuatu"],
    "186": ["Venezuela", "Venezuela"],
    "187": ["Vietnam", "Viêt Nam"],
    "188": ["Yemen", "Yémen"],
    "189": ["Zambia", "Zambie"],
    "190": ["Zimbabwe", "Zimbabwe"],
    "191": ["Palestine", "Palestine"],
    "192": ["St. Lucia", "Sainte-Lucie", "St Lucia"],
    "193": ["Kosovo", "Kosovo"],
    "194": ["Montenegro", "Monténégro"],
    "195": ["Palau", "Palaos"],
    "196": ["Serbia", "Serbie"],
    "197": ["South Sudan", "Soudan du Sud"],
    "198": ["St. Kitts and Nevis", "Saint-Kitts-et-Nevis", "St Kitts & Nevis"],
    "199": ["St. Vincent and the Grenadines", "Saint-Vincent-et-les Grenadines", "Saint Vincent & the Grenadines"],
    "200": ["Tuvalu", "Tuvalu"],
    "201": ["Vatican City", "Cité du Vatican"],
    "202": ["Not found", "Non trouvé"],
}


# =========================================================
# HELPERS
# =========================================================

def _norm(s: str) -> str:
    s = unicodedata.normalize("NFKD", s or "")
    s = "".join(c for c in s if not unicodedata.combining(c)).lower()
    s = re.sub(r"[{}(),.'’\-&]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


NAME_INDEX: List[Tuple[str, int]] = sorted(
    ((_norm(n), int(i)) for i, ns in NAMES.items() for n in ns),
    key=lambda t: -len(t[0]))     # longest first, so "niger" never beats "nigeria"


def match_country(text: str) -> Tuple[Optional[int], str]:
    """Return (index, matched name) for the first country name at the start
    of the answer text."""
    t = _norm(text)
    for name, idx in NAME_INDEX:
        if t == name or t.startswith(name + " "):
            return idx, name
    for name, idx in NAME_INDEX:              # looser: name anywhere in the line
        if re.search(rf"(^| ){re.escape(name)}( |$)", t):
            return idx, name
    return None, ""


def _parse_time(v: Any) -> Optional[datetime]:
    if not v:
        return None
    dt = datetime.fromisoformat(str(v).replace("Z", "+00:00"))
    return dt if dt.tzinfo else dt.replace(tzinfo=timezone.utc)


# =========================================================
# SM APPLY (read-only): application title = reference number
# =========================================================

def app_reference(aid: int) -> str:
    token = wmill.get_variable(SMAPPLY_TOKEN_VAR)
    r = requests.get(f"{SMAPPLY_BASE}/applications/{aid}/",
                     headers={"Authorization": f"Bearer {token}"}, timeout=60)
    if not r.ok:
        return ""
    return str(r.json().get("title") or "").strip()


# =========================================================
# SHAREPOINT (Graph, same credentials as the report script)
# =========================================================

class SharePoint:
    def __init__(self):
        tenant = wmill.get_variable("u/epham2/AZURE_TENANT_ID").strip()
        tok = requests.post(
            f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token",
            data={
                "client_id": wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_ID").strip(),
                "client_secret": wmill.get_variable("u/epham2/SHAREPOINT_CLIENT_SECRET").strip(),
                "scope": "https://graph.microsoft.com/.default",
                "grant_type": "client_credentials",
            },
            timeout=60,
        )
        if tok.status_code != 200:
            raise RuntimeError(f"Graph token request failed {tok.status_code}: {tok.text[:500]}")
        self.h = {"Authorization": f"Bearer {tok.json()['access_token']}"}
        site = self._get(f"/sites/{SHAREPOINT_HOSTNAME}:{SITE_PATH}")
        drives = self._get(f"/sites/{site['id']}/drives")["value"]
        self.drive = next(d for d in drives if d.get("name") == DOCUMENT_LIBRARY_NAME)["id"]

    def _get(self, endpoint: str) -> Any:
        r = requests.get(endpoint if endpoint.startswith("http") else f"{GRAPH_BASE}{endpoint}",
                         headers=self.h, timeout=180)
        if r.status_code == 404:
            return None
        r.raise_for_status()
        return r.json()

    def item(self, path: str) -> Optional[Dict[str, Any]]:
        if not path.strip("/"):
            return self._get(f"/drives/{self.drive}/root")
        return self._get(f"/drives/{self.drive}/root:/{quote(path.strip('/'), safe='/')}")

    def files_under(self, item_id: str, rel: str = "") -> List[Dict[str, Any]]:
        out: List[Dict[str, Any]] = []
        url = (f"{GRAPH_BASE}/drives/{self.drive}/items/{item_id}/children"
               "?$top=200&$select=id,name,file,folder,createdDateTime,lastModifiedDateTime")
        while url:
            data = self._get(url)
            for c in data.get("value", []):
                path = f"{rel}/{c['name']}".strip("/")
                if c.get("folder"):
                    low = c["name"].strip().lower()
                    if low not in SKIP_FOLDERS and not low.startswith(SKIP_FOLDER_PREFIXES):
                        out.extend(self.files_under(c["id"], path))
                elif c.get("file"):
                    stamp = re.search(r"_(\d{8})_(\d{6})\.pdf$", c["name"], re.I)
                    named = (datetime.strptime(stamp.group(1) + stamp.group(2), "%Y%m%d%H%M%S")
                             .replace(tzinfo=timezone.utc) if stamp else None)
                    out.append({"id": c["id"], "path": path,
                                # time in the file name (UTC) when present, else upload time
                                "created": named or _parse_time(c.get("createdDateTime")),
                                "modified": _parse_time(c.get("lastModifiedDateTime"))})
            url = data.get("@odata.nextLink")
        return out

    def download(self, item_id: str) -> bytes:
        r = requests.get(f"{GRAPH_BASE}/drives/{self.drive}/items/{item_id}/content",
                         headers=self.h, timeout=300)
        r.raise_for_status()
        return r.content

    def upload(self, name: str, content: bytes) -> str:
        path = quote(f"{RESULT_FOLDER}/{name}", safe="/")
        r = requests.put(f"{GRAPH_BASE}/drives/{self.drive}/root:/{path}:/content",
                         headers={**self.h, "Content-Type": "application/octet-stream"},
                         data=content, timeout=300)
        r.raise_for_status()
        return r.json().get("webUrl", "")


def folder_candidates(aid: int, reference: str) -> List[str]:
    """Possible folder names: the full reference (NNAS-T-0000000947), its
    number with and without leading zeros (0000000947, 947), and the SM Apply
    application ID."""
    names = [reference]
    m = re.search(r"(\d+)$", reference or "")
    if m:
        names += [m.group(1), str(int(m.group(1)))]
    names.append(str(aid))
    seen, out = set(), []
    for n in names:
        if n and n not in seen:
            seen.add(n)
            out.append(f"{APPLICATIONS_BASE}/{n}")
    return out


def find_app_folder(sp: SharePoint, aid: int, reference: str) -> Tuple[str, Optional[Dict[str, Any]]]:
    for path in folder_candidates(aid, reference):
        it = sp.item(path)
        if it and it.get("folder"):
            return path, it
    return "", None


def sample_folder_names(sp: SharePoint, path: str, limit: int = 15) -> List[str]:
    it = sp.item(path)
    if not it:
        return [f"'{path}' not found in the {DOCUMENT_LIBRARY_NAME} library"]
    data = sp._get(f"/drives/{sp.drive}/items/{it['id']}/children?$top={limit}&$select=name,folder")
    return [c["name"] for c in (data or {}).get("value", [])]


# =========================================================
# PDF TEXT AND ANSWER EXTRACTION
# =========================================================

def pdf_text(content: bytes) -> Tuple[str, str]:
    """Return (text, method). Falls back to OCR for pages without text."""
    reader = PdfReader(io.BytesIO(content))
    pages = [(p.extract_text() or "") for p in reader.pages]
    if sum(len(p.strip()) for p in pages) > 200:
        return "\n".join(pages), "text layer"
    try:
        import pytesseract
        from pdf2image import convert_from_bytes
        images = convert_from_bytes(content, dpi=300)
        return "\n".join(pytesseract.image_to_string(im, lang="eng+fra") for im in images), "OCR"
    except Exception as exc:
        return "", f"needs OCR ({type(exc).__name__}: {str(exc)[:120]})"


NAME_SET: Dict[str, int] = {}
for _n, _i in NAME_INDEX:
    NAME_SET.setdefault(_n, _i)


def match_exact(line: str) -> Optional[int]:
    """Index if the whole line is a country name (ignores a trailing period)."""
    return NAME_SET.get(_norm(line))


def _label_kind(lines: List[str], i: int) -> Optional[str]:
    """'current' or 'previous' if a registration country question starts on
    line i. Looks at the next two lines too, since PDF text often wraps the
    question across lines. Accepts any wording that mentions a country and
    registration/licensing, e.g. "In which country are you currently
    registered..." or "Country of current registration"."""
    if HELP_TEXT.search(lines[i]):
        return None
    head = _norm(lines[i])
    if not re.search(r"\bcountry\b|\bpays\b", head):
        return None
    window = _norm(" ".join(x for x in lines[i:i + 3] if not HELP_TEXT.search(x)))
    if not re.search(r"regist|licen|inscri|autoris|permis", window):
        return None
    if re.search(r"previous|anterieur|precedem|auparavant", window):
        return "previous"
    if re.search(r"current|actuel|presently|\bnow\b", window):
        return "current"
    return None


HELP_TEXT = re.compile(r"choose a country|choisissez un pays|dropdown|liste deroulante", re.I)


SECTION_HEADING = re.compile(
    r"details for your (current|previous) nursing registration\s*(\d*)|"
    r"(actuelle|anterieure|precedente).{0,40}inscription\s*(\d*)", re.I)


def _section(line: str) -> Optional[int]:
    """Field position from a section heading: 0 for "Details for Your
    Current Nursing Registration", N for "...Previous Nursing Registration N"."""
    m = SECTION_HEADING.search(_norm(line))
    if not m:
        return None
    kind = (m.group(1) or m.group(3) or "").lower()
    num = m.group(2) or m.group(4) or ""
    if kind in ("current", "actuelle"):
        return 0
    return int(num) if num.isdigit() else None


def _is_country_question(lines: List[str], i: int) -> bool:
    if HELP_TEXT.search(lines[i]):
        return False
    head = _norm(lines[i])
    if not re.search(r"\bcountry\b|\bpays\b", head):
        return False
    window = _norm(" ".join(x for x in lines[i:i + 3] if not HELP_TEXT.search(x)))
    return bool(re.search(r"regist|licen|inscri|autoris|permis", window))


def extract_answers(text: str) -> Dict[int, str]:
    """Map field position (0 = current, 1-6 = previous) to the country shown
    under each registration country question.

    The position comes from the section heading above the question ("Details
    for Your Current Nursing Registration" / "...Previous Nursing
    Registration N"), because the current question's wording ("In which
    country were you registered or licensed as a nurse?") does not itself say
    current or previous. If no heading has been seen, the question wording is
    used instead. The answer is the first line below the question that is
    exactly a country name, skipping the dropdown help line and stopping at
    the next question or heading."""
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    answers: Dict[int, str] = {}
    section: Optional[int] = None
    prev_n = 0
    for i in range(len(lines)):
        sec = _section(lines[i])
        if sec is not None:
            section = sec
            continue
        if not _is_country_question(lines, i):
            continue
        pos = section
        if pos is None:
            kind = _label_kind(lines, i)
            if kind == "current":
                pos = 0
            elif kind == "previous":
                prev_n += 1
                pos = prev_n
        if pos is None or pos in answers:
            continue
        answer = ""
        tail = lines[i].split("?", 1)
        if len(tail) > 1 and match_exact(tail[1].strip(" :-\t")) is not None:
            answer = tail[1].strip(" :-\t")
        j = i + 1
        while not answer and j < min(len(lines), i + 12):
            if _is_country_question(lines, j) or _section(lines[j]) is not None:
                break
            if not HELP_TEXT.search(lines[j]) and match_exact(lines[j]) is not None:
                answer = lines[j]
            j += 1
        answers[pos] = answer
        section = None          # each heading covers one question
    return answers


def snippets(text: str) -> List[str]:
    """Lines around every line mentioning a country question, for checking
    the layout. Only question text, help text and country names are shown."""
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    out = []
    for i, ln in enumerate(lines):
        n = _norm(ln)
        if "country" in n or "pays" in n:
            ctx = []
            for x in lines[i:i + 6]:
                nx = _norm(x)
                keep = ("country" in nx or "pays" in nx or HELP_TEXT.search(x)
                        or match_exact(x) is not None or re.search(r"regist|licen|inscri", nx))
                ctx.append(x if keep else "[other text]")
            out.append(" | ".join(ctx))
    return out[:30]


# =========================================================
# MAIN
# =========================================================

def verdict(f: Dict[str, Any], shown_idx: Optional[int], copy_time: Optional[datetime]) -> Tuple[str, Optional[int]]:
    untouched, custom = f["stored"], f["custom_idx"]
    if copy_time is None:
        return "Manual: no submitted copy found", None
    if shown_idx is None:
        return "Manual: answer not read from copy", None
    if copy_time >= SWITCH_UTC:
        if copy_time < SWITCH_BACK_UTC and shown_idx == custom:
            return "Inconclusive: submitted during the affected window while the form showed this country", None
        return "Inconclusive: copy saved after the switch", None
    if shown_idx == untouched and shown_idx == custom:
        return "No change", untouched
    if shown_idx == untouched:
        return "No change", untouched
    if shown_idx == custom:
        return "Change", custom
    return "Manual: copy shows a different country", None


def main(discover: bool = False, sample_appid: int = 35650362, upload: bool = True):
    sp = SharePoint()
    by_app: Dict[int, List[Dict[str, Any]]] = {}
    for f in FIELDS:
        by_app.setdefault(f["app"], []).append(f)
    apps = [int(sample_appid)] if discover else sorted(by_app)

    rows, info = [], []
    for aid in apps:
        ref = app_reference(aid)
        folder, it = find_app_folder(sp, aid, ref)
        files = sp.files_under(it["id"]) if it else []
        copies = [x for x in files if x["path"].lower().endswith(".pdf")
                  and re.search(FILE_PATTERN, x["path"].split("/")[-1], re.I)]
        pre = [c for c in copies if c["created"] and c["created"] < SWITCH_UTC]
        chosen = max(pre, key=lambda c: c["created"]) if pre else (
            min(copies, key=lambda c: c["created"] or SWITCH_BACK_UTC) if copies else None)

        text, method, answers = "", "", {}
        if chosen:
            text, method = pdf_text(sp.download(chosen["id"]))
            answers = extract_answers(text)

        if discover:
            return {
                "application": aid, "reference": ref, "folder": folder or "not found",
                "paths_tried": folder_candidates(aid, ref),
                "sample_folders_in_base": [] if folder else sample_folder_names(sp, APPLICATIONS_BASE),
                "library_root_sample": [] if folder else sample_folder_names(sp, ""),
                "all_files": [{"path": x["path"], "created_utc": str(x["created"])} for x in files],
                "matching_copies": [x["path"] for x in copies],
                "chosen_copy": chosen["path"] if chosen else None,
                "text_method": method,
                "question_snippets": snippets(text)[:10],
                "answers_by_position": answers,
            }

        info.append({"application_id": aid, "reference": ref, "folder": folder,
                     "copies_found": len(copies), "chosen": chosen["path"] if chosen else "",
                     "method": method})
        for f in by_app[aid]:
            shown_text = answers.get(f["pos"], "")
            shown_idx, _ = match_country(shown_text) if shown_text else (None, "")
            if method.startswith("needs OCR"):
                v, idx = "Manual: needs OCR", None
            else:
                v, idx = verdict(f, shown_idx, chosen["created"] if chosen else None)
            rows.append({
                "application_id": aid, "reference": ref, "field": f["field"],
                "stored_value": f["stored"],
                "untouched_reading": NAMES[str(f["stored"])][0],
                "reselected_reading": NAMES[str(f["custom_idx"])][0],
                "copy_file": chosen["path"] if chosen else "",
                "copy_saved_utc": chosen["created"].strftime("%Y-%m-%d %H:%M") if chosen and chosen["created"] else "",
                "copy_shows": shown_text[:80],
                "copy_shows_matched": NAMES[str(shown_idx)][0] if shown_idx is not None else "",
                "verdict": v,
                "index_to_store": idx if idx is not None else "",
                "country_to_store": NAMES[str(idx)][0] if idx is not None else "",
                "text_method": method,
            })

    counts: Dict[str, int] = {}
    for r in rows:
        k = r["verdict"].split(":")[0]
        counts[k] = counts.get(k, 0) + 1
    result: Dict[str, Any] = {"fields_checked": len(rows), "verdicts": counts, "applications": info}

    if upload:
        buf = io.StringIO()
        w = csv.DictWriter(buf, fieldnames=list(rows[0].keys()) if rows else ["application_id"])
        w.writeheader()
        w.writerows(rows)
        ts = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
        result["sharepoint_file"] = sp.upload(f"submitted_copy_check_{ts}.csv", buf.getvalue().encode("utf-8"))
    else:
        result["rows"] = rows
    return result