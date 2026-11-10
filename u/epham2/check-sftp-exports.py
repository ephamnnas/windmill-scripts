# py: 3.12
# requirements:
# requests
# wmill

"""
READ-ONLY: list every SFTP export package in SharePoint for the 72
applications in registration_country_final.xlsx, and flag the ones created
after Friday's switch (2026-10-30 21:20 UTC), which may contain incorrect
registration countries and need regenerating and resending.

Export folders look like TM/<reference>/SFTP_Export_<YYYYMMDD>_<HHMMSS>/<RB>/<reference>.zip
(for example SFTP_Export_20261104_202512/CNO/NNAS-T-0000000947.zip), where
<RB> is the regulatory body the package went to.

Reference numbers missing from the workbook are looked up in SM Apply (the
application title). Nothing is changed; only a results CSV is uploaded to
Applications/Country list backups.
"""

import csv
import io
import re
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional
from urllib.parse import quote

import requests
import wmill

SMAPPLY_BASE = "https://nnas.smapply.io/api"
SMAPPLY_TOKEN_VAR = "u/epham2/SMAPPLY_ACCESS_TOKEN"
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SHAREPOINT_HOSTNAME = "nnasca.sharepoint.com"
SITE_PATH = "/sites/NNAS-Portal-Sharepoint"
DOCUMENT_LIBRARY_NAME = "Applications"
APPLICATIONS_BASE = "TM"
RESULT_FOLDER = "Country list backups"

SWITCH_UTC = datetime(2026, 10, 30, 21, 20, tzinfo=timezone.utc)
SWITCH_BACK_UTC = datetime(2026, 11, 5, 16, 20, tzinfo=timezone.utc)

# application ID: (reference, overall status in the final workbook, stage)
APPS = {
    35500617: ("", "Contact applicant", "Application Stage"),
    35589355: ("NNAS-T-0000000797", "Changed, regenerate", "Document Intake & Review"),
    35590885: ("NNAS-T-0000000803", "Changed, regenerate", "Document Intake & Review"),
    35596873: ("", "Contact applicant", "Application Stage"),
    35604818: ("NNAS-T-0000000837", "No change needed", "Application Review"),
    35625530: ("NNAS-T-0000000891", "Changed, regenerate", "Document Intake & Review"),
    35647178: ("", "Changed", "Application Stage"),
    35650362: ("NNAS-T-0000000947", "Changed, regenerate", "Credential Report Issued"),
    35671256: ("", "Contact applicant", "Application Stage"),
    35671271: ("NNAS-T-0000000998", "Changed, regenerate", "Document Intake & Review"),
    35690631: ("", "Contact applicant", "Application Stage"),
    35703034: ("", "Changed, regenerate", "Document Intake & Review"),
    35716321: ("NNAS-T-0000001084", "Changed, regenerate", "Document Intake & Review"),
    35768338: ("", "Changed, regenerate", "Document Intake & Review"),
    35791245: ("", "Contact applicant", "Application Stage"),
    35801406: ("", "Contact applicant", "Application Stage"),
    35817358: ("", "Changed", "Application Stage"),
    35853913: ("NNAS-T-0000001445", "Changed, regenerate", "Credential Report Validation - Finalize Application"),
    35862552: ("", "Contact applicant", "Application Stage"),
    35864060: ("NNAS-T-0000001487", "Changed, regenerate", "Application Review"),
    35874317: ("", "Changed", "Application Stage"),
    35879414: ("NNAS-T-0000001542", "Changed, regenerate", "Document Intake & Review"),
    35879876: ("", "Contact applicant", "Application Stage"),
    35894301: ("", "Changed", "Application Stage"),
    35896712: ("", "Changed", "Application Stage"),
    35904275: ("", "Contact applicant", "Application Stage"),
    35908317: ("NNAS-T-0000001627", "Changed, regenerate", "Document Intake & Review"),
    35914759: ("", "Changed", "Application Stage"),
    35925873: ("", "Changed", "Application Stage"),
    35934846: ("", "Contact applicant", "Application Stage"),
    35937949: ("NNAS-T-0000001733", "No action", "Document Intake & Review"),
    35942863: ("", "Changed, regenerate", "Application Review"),
    35945656: ("NNAS-T-0000001776", "Changed, regenerate", "Document Intake & Review"),
    35945662: ("NNAS-T-0000001777", "Changed, regenerate", "Credential Report Issued"),
    35949706: ("", "Changed, regenerate", "Document Intake & Review"),
    35952334: ("", "Changed", "Application Stage"),
    35962741: ("NNAS-T-0000001816", "Changed, regenerate", "Document Intake & Review"),
    35968951: ("", "Changed, regenerate", "Document Intake & Review"),
    35969748: ("", "Contact applicant", "Application Stage"),
    35969773: ("", "Changed", "Application Stage"),
    35970115: ("NNAS-T-0000001839", "Changed, regenerate", "Document Intake & Review"),
    35971544: ("", "Contact applicant", "Application Stage"),
    35975343: ("", "Changed", "Application Stage"),
    35976737: ("", "Changed", "Application Stage"),
    35978096: ("", "Changed, regenerate", "Document Intake & Review"),
    35979241: ("", "Changed", "Application Stage"),
    35980266: ("", "Changed, regenerate", "Document Intake & Review"),
    35981909: ("", "Changed, regenerate", "Document Intake & Review"),
    35982316: ("", "Changed", "Application Stage"),
    35984792: ("", "Changed, regenerate", "Document Intake & Review"),
    35985498: ("", "Changed, regenerate", "Document Intake & Review"),
    35986821: ("", "Changed", "Application Stage"),
    35988943: ("", "Changed", "Application Stage"),
    35989640: ("", "Changed", "Application Stage"),
    35990640: ("", "Changed", "Application Stage"),
    35990665: ("", "Changed, regenerate", "Document Intake & Review"),
    35991279: ("", "Changed", "Application Stage"),
    35991913: ("", "Changed", "Application Stage"),
    35992469: ("", "Changed, regenerate", "Document Intake & Review"),
    35994869: ("", "Changed", "Application Stage"),
    35996425: ("", "Changed, regenerate", "Document Intake & Review"),
    35997088: ("", "Changed, regenerate", "Document Intake & Review"),
    35997252: ("", "Changed", "Application Stage"),
    35997774: ("", "Changed", "Application Stage"),
    35999153: ("", "Changed", "Application Stage"),
    36000810: ("", "Changed", "Application Stage"),
    36002836: ("", "Changed", "Application Stage"),
    36003360: ("", "Changed", "Application Stage"),
    36004071: ("", "Changed", "Application Stage"),
    36007610: ("", "Changed", "Application Stage"),
    36008235: ("", "Changed", "Application Stage"),
    36009809: ("", "Changed", "Application Stage"),
}


def app_reference(aid: int) -> str:
    token = wmill.get_variable(SMAPPLY_TOKEN_VAR)
    r = requests.get(f"{SMAPPLY_BASE}/applications/{aid}/",
                     headers={"Authorization": f"Bearer {token}"}, timeout=60)
    return str(r.json().get("title") or "").strip() if r.ok else ""


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
        tok.raise_for_status()
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
        return self._get(f"/drives/{self.drive}/root:/{quote(path.strip('/'), safe='/')}")

    def children(self, item_id: str) -> List[Dict[str, Any]]:
        out, url = [], (f"{GRAPH_BASE}/drives/{self.drive}/items/{item_id}/children"
                        "?$top=200&$select=id,name,file,folder,createdDateTime")
        while url:
            d = self._get(url)
            out.extend(d.get("value", []))
            url = d.get("@odata.nextLink")
        return out

    def upload(self, name: str, content: bytes) -> str:
        path = quote(f"{RESULT_FOLDER}/{name}", safe="/")
        r = requests.put(f"{GRAPH_BASE}/drives/{self.drive}/root:/{path}:/content",
                         headers={**self.h, "Content-Type": "application/octet-stream"},
                         data=content, timeout=300)
        r.raise_for_status()
        return r.json().get("webUrl", "")


def _export_time(folder_name: str) -> Optional[datetime]:
    m = re.match(r"SFTP_Export_(\d{8})_(\d{6})", folder_name, re.I)
    if not m:
        return None
    return datetime.strptime(m.group(1) + m.group(2), "%Y%m%d%H%M%S").replace(tzinfo=timezone.utc)


def main(upload: bool = True):
    sp = SharePoint()
    rows: List[Dict[str, Any]] = []
    for aid, (ref, status, stage) in APPS.items():
        ref = ref or app_reference(aid)
        folder = sp.item(f"{APPLICATIONS_BASE}/{ref}") if ref else None
        base = {"application_id": aid, "reference": ref, "country_status": status, "stage": stage}
        if not folder:
            rows.append({**base, "export_folder": "", "note": "application folder not found"})
            continue
        exports = [c for c in sp.children(folder["id"])
                   if c.get("folder") and c["name"].lower().startswith("sftp_export")]
        if not exports:
            rows.append({**base, "export_folder": "", "note": "no SFTP exports"})
            continue
        for ex in exports:
            when = _export_time(ex["name"])
            for rb in [c for c in sp.children(ex["id"]) if c.get("folder")] or [None]:
                files = [f["name"] for f in sp.children(rb["id"]) if f.get("file")] if rb else []
                if when is None:
                    period = "unknown"
                elif when < SWITCH_UTC:
                    period = "before switch"
                elif when < SWITCH_BACK_UTC:
                    period = "during affected window"
                else:
                    period = "after switch back"
                needs = (period == "during affected window"
                         or (period == "after switch back" and status in ("Changed, regenerate", "Verify")))
                rows.append({**base,
                             "export_folder": ex["name"],
                             "export_time_utc": when.strftime("%Y-%m-%d %H:%M") if when else "",
                             "regulatory_body": rb["name"] if rb else "",
                             "files": "; ".join(files),
                             "period": period,
                             "resend_needed": "Yes" if needs and status != "No change needed" else "Check" if needs else "No",
                             "note": ""})

    flagged = [r for r in rows if r.get("resend_needed") in ("Yes", "Check")]
    result = {
        "applications_checked": len(APPS),
        "exports_found": sum(1 for r in rows if r.get("export_folder")),
        "exports_needing_resend": len(flagged),
        "applications_needing_resend": sorted({r["reference"] or r["application_id"] for r in flagged}),
        "by_regulatory_body": {},
        "folders_not_found": [r["application_id"] for r in rows if r.get("note") == "application folder not found"],
    }
    for r in flagged:
        result["by_regulatory_body"][r["regulatory_body"]] = result["by_regulatory_body"].get(r["regulatory_body"], 0) + 1

    if upload:
        fields = ["application_id", "reference", "country_status", "stage", "export_folder",
                  "export_time_utc", "period", "regulatory_body", "files", "resend_needed", "note"]
        buf = io.StringIO()
        w = csv.DictWriter(buf, fieldnames=fields, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)
        ts = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
        result["sharepoint_file"] = sp.upload(f"sftp_export_check_{ts}.csv", buf.getvalue().encode("utf-8"))
    return result