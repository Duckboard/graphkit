"""
graphkit.auth
=============
Authenticates to the Microsoft Graph API using the OAuth2 "client
credentials" flow -- also called app-only auth. This is the flow you
want when a script needs to act on a mailbox, Planner plan, SharePoint
site, etc. *without* a human signing in interactively (e.g. a scheduled
job, a background automation, a shared-mailbox integration).

No secrets live in this file or anywhere else in this package. All
three required values are read from environment variables at runtime:

    GRAPH_TENANT_ID
    GRAPH_CLIENT_ID
    GRAPH_CLIENT_SECRET

See the README for how to create these by registering an app in
Azure AD (Entra ID), and .env.example for a template you can copy to
a real .env file (which you should never commit to version control).

Usage
-----
    from graphkit import get_token

    token = get_token()
"""

import os
from typing import Optional

import msal

GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]


class GraphAuthError(RuntimeError):
    """Raised when a Graph API access token could not be acquired."""


def _require_env(name: str) -> str:
    value = os.environ.get(name)
    if not value:
        raise GraphAuthError(
            f"Missing required environment variable: {name}\n"
            "Set it in your shell, or add it to a .env file in your "
            "project folder (see .env.example) and load it with "
            "python-dotenv before calling get_token()."
        )
    return value


def get_token(
    tenant_id: Optional[str] = None,
    client_id: Optional[str] = None,
    client_secret: Optional[str] = None,
) -> str:
    """
    Acquire an app-only Microsoft Graph access token.

    If tenant_id / client_id / client_secret are not passed in
    directly, they are read from the GRAPH_TENANT_ID, GRAPH_CLIENT_ID,
    and GRAPH_CLIENT_SECRET environment variables.

    MSAL keeps its own in-memory token cache for the lifetime of the
    process and will reuse a still-valid token instead of calling
    Azure again, so it's fine to call this once per script, or before
    every API call, without worrying about hammering the token
    endpoint.

    Raises GraphAuthError if credentials are missing or invalid.
    """
    tenant_id = tenant_id or _require_env("GRAPH_TENANT_ID")
    client_id = client_id or _require_env("GRAPH_CLIENT_ID")
    client_secret = client_secret or _require_env("GRAPH_CLIENT_SECRET")

    app = msal.ConfidentialClientApplication(
        client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        client_credential=client_secret,
    )

    result = app.acquire_token_for_client(scopes=GRAPH_SCOPE)

    if "access_token" not in result:
        raise GraphAuthError(
            result.get("error_description", str(result))
        )

    return result["access_token"]
