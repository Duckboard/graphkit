"""
graphkit.client
================
A deliberately thin wrapper around the Microsoft Graph REST API
(v1.0). It knows nothing about Outlook, Planner, SharePoint, or any
other specific Graph resource -- it just handles the repetitive parts
(base URL, auth header, JSON) so you can call any Graph endpoint with
one line and build your own helper functions on top.

Usage
-----
    from graphkit import get_token, GraphClient

    token = get_token()
    graph = GraphClient(token)

    me = graph.get("/users/someone@example.com")

    graph.post("/users/someone@example.com/messages", {
        "subject": "Hello",
        "body": {"contentType": "Text", "content": "Hi there"},
        "toRecipients": [{"emailAddress": {"address": "other@example.com"}}],
    })
"""

from typing import Optional

import requests

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"


class GraphClient:
    """
    Thin HTTP client for the Microsoft Graph API.

    Pass in a bearer token from graphkit.auth.get_token(). One
    GraphClient per token -- if a token expires mid-run, create a new
    client with a freshly acquired token.
    """

    def __init__(self, token: str, base_url: str = GRAPH_BASE_URL):
        self.token = token
        self.base_url = base_url

    def _headers(self, extra: Optional[dict] = None) -> dict:
        headers = {"Authorization": f"Bearer {self.token}"}
        if extra:
            headers.update(extra)
        return headers

    def get(self, path: str) -> dict:
        """GET a Graph path (e.g. '/me' or '/users/x@y.com/messages') and
        return the parsed JSON body. Raises on non-2xx responses."""
        r = requests.get(f"{self.base_url}{path}", headers=self._headers())
        r.raise_for_status()
        return r.json()

    def post(self, path: str, body: dict) -> requests.Response:
        """POST a JSON body to a Graph path. Returns the raw response --
        check response.status_code (Graph typically returns 201 for a
        successful create) rather than assuming success."""
        return requests.post(
            f"{self.base_url}{path}",
            headers=self._headers({"Content-Type": "application/json"}),
            json=body,
        )

    def patch(self, path: str, body: dict) -> requests.Response:
        """PATCH (partial update) a JSON body to a Graph path."""
        return requests.patch(
            f"{self.base_url}{path}",
            headers=self._headers({"Content-Type": "application/json"}),
            json=body,
        )

    def delete(self, path: str) -> requests.Response:
        """DELETE a Graph resource."""
        return requests.delete(f"{self.base_url}{path}", headers=self._headers())

    @staticmethod
    def error_message(response: requests.Response) -> str:
        """Best-effort extraction of a human-readable error message from
        a failed Graph response, for logging/printing."""
        try:
            return response.json().get("error", {}).get("message", response.text[:300])
        except ValueError:
            return response.text[:300]
