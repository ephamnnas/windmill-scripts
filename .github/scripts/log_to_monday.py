"""
Log Windmill deployments to Monday.com, with item names written by Claude.

Runs inside a GitHub Action on each push from Windmill Git Sync. For every
commit, it groups the changed files by Windmill object (script, flow or app),
asks Claude to summarise each object's code changes in one changelog line, and
creates one Monday.com item per object using that line as the item name.

What leaves GitHub:
  * To Claude: the code diff for each changed object (lock files and
    package files are skipped) and the commit message.
  * To Monday.com: the summary line plus metadata (path, type, action,
    author, timestamp, commit message and links).
Job arguments and job results are never read, as they are not in the repo.

If Claude is unavailable or no API key is set, the item name falls back to
"<path> (<type>) <action>" so logging never stops.
"""

import json
import os
import subprocess
import sys
import urllib.request
from datetime import datetime, timezone

MONDAY_URL = "https://api.monday.com/v2"
CLAUDE_URL = "https://api.anthropic.com/v1/messages"
DEFAULT_MODEL = "claude-haiku-5-5"
MAX_DIFF_CHARS = 20000

SCRIPT_SUFFIXES = (
    ".script.yaml", ".script.json", ".script.lock",
    ".bun.ts", ".deno.ts", ".fetch.ts",
    ".py", ".ts", ".js", ".go", ".sh", ".sql", ".ps1", ".php", ".rs", ".cs",
)
IGNORED_PREFIXES = (".github/", "wmill.yaml", "wmill-lock.yaml", "settings.yaml", "README")
NOISY_FILES = (".lock", "package.json", "package-lock.json")

SUMMARY_PROMPT = """You write entries for a change log of automation scripts that run an \
online application portal (the scripts connect to SMApply and SharePoint).

Write ONE line describing what changed in the Windmill {obj_type} "{path}", based on the \
diff below. Follow the style of these existing entries:
- Added Urdu to the Languages bulk preset
- Updated licensure-form-tm script to first delete current form if existed
- Fixed a bug in credential-form script where expiry date and exam language could be skipped
- Unified TM/CES applicant-docs script

Rules:
- Past tense, starting with a verb such as Added, Updated, Fixed, Removed or Renamed.
- Under 120 characters. Plain text only: no quotes, no markdown, no em-dashes.
- Describe the behaviour change, not the file or line numbers.
- Never include personal information, email addresses, tokens or secrets, even if they appear in the diff.
- If the diff is too small or unclear to tell, write a short neutral line such as "Updated <name> {obj_type_lower}".
- Reply with the line only.

Commit message (may include the developer's own description):
{message}

Diff:
{diff}"""


def env(name, required=False):
    value = os.environ.get(name, "").strip()
    if required and not value:
        sys.exit(f"Missing required setting: {name}")
    return value or None


def object_from_file(file_path):
    """Map a file in the repository back to the Windmill object it belongs to."""
    for marker, obj_type in ((".flow/", "Flow"), ("__flow/", "Flow"),
                             (".app/", "App"), ("__app/", "App"),
                             (".raw_app/", "App"), ("__raw_app/", "App")):
        if marker in file_path:
            return obj_type, file_path.split(marker)[0]
    for suffix in SCRIPT_SUFFIXES:
        if file_path.endswith(suffix):
            return "Script", file_path[: -len(suffix)]
    return "Other", file_path


def changed_files(sha):
    """Ask Git which files a commit changed. Returns a list of (action, path).

    GitHub's push details do not always list changed files (Windmill's pushes
    arrive without them), so the repository history is the reliable source.
    """
    try:
        out = subprocess.run(
            ["git", "diff-tree", "--root", "--no-commit-id", "--name-status",
             "--no-renames", "-r", sha],
            capture_output=True, text=True, check=True,
        ).stdout
    except subprocess.CalledProcessError as exc:
        print(f"Could not list files for {sha}: {exc.stderr.strip()}")
        return []
    actions = {"A": "added", "M": "modified", "D": "deleted"}
    files = []
    for line in out.splitlines():
        parts = line.split("\t", 1)
        if len(parts) == 2:
            files.append((actions.get(parts[0][:1], "modified"), parts[1]))
    return files


def group_objects(commit):
    """Return a list of dicts: type, path, action and the files that changed."""
    objects = {}
    for action, file_path in changed_files(commit["id"]):
        if file_path.startswith(IGNORED_PREFIXES):
            continue
        obj_type, path = object_from_file(file_path)
        obj = objects.setdefault((obj_type, path), {
            "type": obj_type, "path": path, "actions": set(), "files": []})
        obj["actions"].add(action)
        obj["files"].append(file_path)
    result = []
    for obj in objects.values():
        actions = obj.pop("actions")
        if actions == {"deleted"}:
            obj["action"] = "deleted"
        elif actions == {"added"}:
            obj["action"] = "added"
        else:
            obj["action"] = "modified"
        result.append(obj)
    return result


def get_diff(sha, files):
    files = [f for f in files if not f.endswith(NOISY_FILES)]
    if not files:
        return ""
    try:
        out = subprocess.run(
            ["git", "show", "--format=", "--unified=3", sha, "--", *files],
            capture_output=True, text=True, check=True,
        ).stdout
    except subprocess.CalledProcessError as exc:
        print(f"Could not read diff for {sha}: {exc.stderr.strip()}")
        return ""
    if len(out) > MAX_DIFF_CHARS:
        out = out[:MAX_DIFF_CHARS] + "\n[diff truncated]"
    return out


def post_json(url, payload, headers):
    request = urllib.request.Request(
        url, data=json.dumps(payload).encode(),
        headers={"Content-Type": "application/json", **headers}, method="POST")
    with urllib.request.urlopen(request, timeout=60) as response:
        return json.loads(response.read())


def summarise(obj, commit_message, diff):
    """Ask Claude for a one-line change log entry. Returns None on any failure."""
    api_key = env("ANTHROPIC_API_KEY")
    if not api_key or not diff or obj["action"] == "deleted":
        return None
    prompt = SUMMARY_PROMPT.format(
        obj_type=obj["type"].lower(), obj_type_lower=obj["type"].lower(),
        path=obj["path"], message=(commit_message or "(none)")[:1000], diff=diff)
    try:
        body = post_json(CLAUDE_URL, {
            "model": env("CLAUDE_MODEL") or DEFAULT_MODEL,
            "max_tokens": 200,
            "messages": [{"role": "user", "content": prompt}],
        }, {"x-api-key": api_key, "anthropic-version": "2023-06-01"})
        text = "".join(block.get("text", "") for block in body.get("content", []))
        line = text.strip().splitlines()[0].strip().strip('"') if text.strip() else ""
        return line.replace("\u2014", ",") or None
    except Exception as exc:  # Never let a summary failure stop the logging
        print(f"Claude summary failed for {obj['path']}: {exc}")
        return None


def windmill_link(obj):
    base, workspace = env("WINDMILL_BASE_URL"), env("WINDMILL_WORKSPACE")
    route = {"Script": "scripts", "Flow": "flows", "App": "apps"}.get(obj["type"])
    if not base or not workspace or not route or obj["action"] == "deleted":
        return None
    return f"{base.rstrip('/')}/{route}/get/{obj['path']}?workspace={workspace}"


def build_columns(commit, obj):
    cols = {}
    ts = datetime.fromisoformat(commit["timestamp"]).astimezone(timezone.utc)
    setters = {
        "MONDAY_COL_PATH": obj["path"],
        "MONDAY_COL_AUTHOR": commit.get("author", {}).get("name", "unknown"),
        "MONDAY_COL_DATE": {"date": ts.strftime("%Y-%m-%d"), "time": ts.strftime("%H:%M:%S")},
        "MONDAY_COL_TYPE": obj["type"],
        "MONDAY_COL_ACTION": {"label": obj["action"].capitalize()},
        "MONDAY_COL_MESSAGE": {"text": (commit.get("message") or "")[:2000]},
        "MONDAY_COL_COMMIT_LINK": {"url": commit["url"], "text": commit["id"][:7]},
    }
    for setting, value in setters.items():
        if env(setting):
            cols[env(setting)] = value
    link = windmill_link(obj)
    if env("MONDAY_COL_WINDMILL_LINK") and link:
        cols[env("MONDAY_COL_WINDMILL_LINK")] = {"url": link, "text": "Open in Windmill"}
    return cols


def create_item(name, columns):
    group_id = env("MONDAY_GROUP_ID")
    group_param = ", $group: String" if group_id else ""
    group_arg = ", group_id: $group" if group_id else ""
    query = f"""
    mutation ($board: ID!, $name: String!, $cols: JSON{group_param}) {{
      create_item(board_id: $board, item_name: $name, column_values: $cols,
                  create_labels_if_missing: true{group_arg}) {{ id }}
    }}"""
    variables = {"board": env("MONDAY_BOARD_ID", required=True),
                 "name": name[:255], "cols": json.dumps(columns)}
    if group_id:
        variables["group"] = group_id
    body = post_json(MONDAY_URL, {"query": query, "variables": variables},
                     {"Authorization": env("MONDAY_API_TOKEN", required=True)})
    # Monday.com reports most errors with HTTP 200, so check the body too.
    if body.get("errors") or body.get("error_message"):
        raise RuntimeError(f"Monday.com API error: {body}")
    return body["data"]["create_item"]["id"]


def main():
    with open(os.environ["GITHUB_EVENT_PATH"]) as f:
        event = json.load(f)

    commits = event.get("commits", [])
    if not commits:
        print("No commits in this push. Nothing to log.")
        return

    created = 0
    for commit in commits:
        objects = group_objects(commit)
        print(f"Commit {commit['id'][:7]}: {len(objects)} Windmill object(s) changed")
        for obj in objects:
            diff = get_diff(commit["id"], obj["files"])
            summary = summarise(obj, commit.get("message"), diff)
            name = summary or f"{obj['path']} ({obj['type']}) {obj['action']}"
            item_id = create_item(name, build_columns(commit, obj))
            print(f"Created Monday.com item {item_id}: {name}")
            created += 1

    print(f"Done. Created {created} item(s) from {len(commits)} commit(s).")


if __name__ == "__main__":
    main()
