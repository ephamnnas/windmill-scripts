import requests
from datetime import datetime

SCRIPT_NAMES = {
    "testScript": "u/epham2/test-script",
    "applicationTM": "u/epham2/application-tm",
    "applicationCES": "u/epham2/application-ces",
    "applicantDocsTM": "u/epham2/applicant-docs-tm",
    "applicantDocsCES": "u/epham2/applicant-docs-ces",
    "licensureFormTM": "u/epham2/licensure-form-tm",
    "licensureFormCES": "u/epham2/licensure-form-ces",
    "credentialFormTM": "u/epham2/credential-form-tm",
    "credentialFormCES": "u/epham2/credential-form-ces",
}

SCRIPT_TO_FLOW = {
    "testScript": "Test Flow",
    "applicationTM": "Application Flow (TM)",
    "applicationCES": "Application Flow (CES)",
    "applicantDocsTM": "u/epham2/applicant-docs-tm",
    "applicantDocsCES": "u/epham2/applicant-docs-ces",
    "licensureFormTM": "Licensure Form Flow (TM)",
    "licensureFormCES": "Licensure Form Flow (CES)",
    "credentialFormTM": "Credential Form Flow (TM)",
    "credentialFormCES": "Credential Form Flow (CES)"
}

def main(
    error: dict,
    ms_teams_webhook: dict,
    schedule_path: str = "",
):
    workspace_id = "nnas-prod"
    error_message = error.get("message", "No message") if isinstance(error, dict) else str(error)
    error_stack = error.get("stack", "No stack trace") if isinstance(error, dict) else ""
    step_id = error.get("step_id", "unknown") if isinstance(error, dict) else "unknown"

    flow_name = SCRIPT_TO_FLOW.get(step_id, "Unknown Flow")
    script_name = SCRIPT_NAMES.get(step_id, step_id)
    failed_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    payload = {
        "type": "message",
        "attachments": [
            {
                "contentType": "application/vnd.microsoft.card.adaptive",
                "content": {
                    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                    "type": "AdaptiveCard",
                    "version": "1.2",
                    "body": [
                        {
                            "type": "TextBlock",
                            "text": "❌ Windmill Job Failed",
                            "weight": "Bolder",
                            "size": "Medium",
                            "color": "Attention",
                        },
                        {
                            "type": "FactSet",
                            "facts": [
                                {"title": "Flow", "value": flow_name},
                                {"title": "Script", "value": script_name},
                                {"title": "Workspace", "value": workspace_id},
                                {"title": "Time", "value": failed_at},
                                {"title": "Schedule", "value": schedule_path or "Manual trigger"},
                            ],
                        },
                        {
                            "type": "TextBlock",
                            "text": f"**Error:** {error_message[:300]}",
                            "wrap": True,
                            "color": "Attention",
                        },
                        {
                            "type": "TextBlock",
                            "text": f"**Stack Trace:**\n{error_stack[:500]}",
                            "wrap": True,
                            "isSubtle": True,
                            "size": "Small",
                        },
                    ],
                },
            }
        ],
    }

    response = requests.post(
        ms_teams_webhook["webhook_url"],
        json=payload,
        headers={"Content-Type": "application/json"},
    )

    return {"status": response.status_code}
