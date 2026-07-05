"""
graphkit -- a small, dependency-light toolkit for authenticating to the
Microsoft Graph API (app-only / client credentials flow) and making
simple REST calls against it.

    from graphkit import get_token, GraphClient

    token = get_token()
    graph = GraphClient(token)
    me = graph.get("/users/someone@example.com")
"""

from .auth import get_token, GraphAuthError
from .client import GraphClient

__all__ = ["get_token", "GraphAuthError", "GraphClient"]
__version__ = "0.1.0"
