"""
send_email_draft.py -- create a draft email in a mailbox via Graph API

Run with:  python send_email_draft.py

This is an example of using graphkit's get_token() + GraphClient on
top of the raw Graph API -- there's no dedicated "send email" helper
in graphkit itself, because email is just one of many things you can
do with a Graph token. Copy this file and adapt it for your own use
case.

Requires an Azure app registration with the Mail.ReadWrite
Application permission (admin consent required) -- see README.md.

Before running, set these environment variables (e.g. via a .env
file loaded with python-dotenv, or your shell):
    GRAPH_TENANT_ID
    GRAPH_CLIENT_ID
    GRAPH_CLIENT_SECRET
    MAILBOX            -- the mailbox to create the draft in,
                           e.g. someone@yourdomain.com
"""

import os
import sys

# Uncomment these two lines if you're using a .env file:
# from dotenv import load_dotenv
# load_dotenv()

from graphkit import get_token, GraphClient


def create_draft(graph: GraphClient, mailbox: str, to: str, subject: str, body_html: str, cc: str = None) -> bool:
    """Create a draft email in `mailbox`'s Drafts folder. Returns True on success."""
    message = {
        "subject": subject,
        "body": {"contentType": "HTML", "content": body_html},
        "toRecipients": [{"emailAddress": {"address": to}}],
        "from": {"emailAddress": {"address": mailbox}},
    }
    if cc:
        message["ccRecipients"] = [{"emailAddress": {"address": cc}}]

    response = graph.post(f"/users/{mailbox}/messages", message)

    if response.status_code == 201:
        print(f"  Draft created -> {to} | {subject}")
        return True

    print(f"  FAILED ({response.status_code}): {GraphClient.error_message(response)}")
    return False


def main():
    mailbox = os.environ.get("MAILBOX")
    if not mailbox:
        print("Set the MAILBOX environment variable to the mailbox to draft from.")
        sys.exit(1)

    token = get_token()
    graph = GraphClient(token)

    # --- Replace this with your own recipients and content ---
    recipients = [
        {"to": "customer@example.com", "name": "Example Ltd", "cc": None},
    ]
    for recipient in recipients:
        body = f"<p>Dear {recipient['name']},</p><p>Your message here.</p>"
        create_draft(graph, mailbox, recipient["to"], "Your subject here", body, cc=recipient["cc"])


if __name__ == "__main__":
    main()
