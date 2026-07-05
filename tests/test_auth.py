"""
Tests for graphkit.auth.

Run with:  pytest

These tests don't call the real Azure endpoint -- they mock msal so
they run offline and don't need real credentials.
"""

import os
from unittest.mock import MagicMock, patch

import pytest

from graphkit.auth import get_token, GraphAuthError


def test_get_token_raises_without_env_vars(monkeypatch):
    monkeypatch.delenv("GRAPH_TENANT_ID", raising=False)
    monkeypatch.delenv("GRAPH_CLIENT_ID", raising=False)
    monkeypatch.delenv("GRAPH_CLIENT_SECRET", raising=False)

    with pytest.raises(GraphAuthError, match="GRAPH_TENANT_ID"):
        get_token()


@patch("graphkit.auth.msal.ConfidentialClientApplication")
def test_get_token_returns_access_token(mock_app_cls, monkeypatch):
    monkeypatch.setenv("GRAPH_TENANT_ID", "fake-tenant")
    monkeypatch.setenv("GRAPH_CLIENT_ID", "fake-client")
    monkeypatch.setenv("GRAPH_CLIENT_SECRET", "fake-secret")

    mock_app = MagicMock()
    mock_app.acquire_token_for_client.return_value = {"access_token": "fake-token-value"}
    mock_app_cls.return_value = mock_app

    token = get_token()

    assert token == "fake-token-value"
    mock_app_cls.assert_called_once_with(
        "fake-client",
        authority="https://login.microsoftonline.com/fake-tenant",
        client_credential="fake-secret",
    )


@patch("graphkit.auth.msal.ConfidentialClientApplication")
def test_get_token_raises_on_auth_failure(mock_app_cls, monkeypatch):
    monkeypatch.setenv("GRAPH_TENANT_ID", "fake-tenant")
    monkeypatch.setenv("GRAPH_CLIENT_ID", "fake-client")
    monkeypatch.setenv("GRAPH_CLIENT_SECRET", "wrong-secret")

    mock_app = MagicMock()
    mock_app.acquire_token_for_client.return_value = {
        "error": "invalid_client",
        "error_description": "AADSTS7000215: Invalid client secret provided.",
    }
    mock_app_cls.return_value = mock_app

    with pytest.raises(GraphAuthError, match="Invalid client secret"):
        get_token()


def test_get_token_accepts_explicit_args(monkeypatch):
    # Explicit args should be used even if env vars are unset.
    monkeypatch.delenv("GRAPH_TENANT_ID", raising=False)
    monkeypatch.delenv("GRAPH_CLIENT_ID", raising=False)
    monkeypatch.delenv("GRAPH_CLIENT_SECRET", raising=False)

    with patch("graphkit.auth.msal.ConfidentialClientApplication") as mock_app_cls:
        mock_app = MagicMock()
        mock_app.acquire_token_for_client.return_value = {"access_token": "explicit-token"}
        mock_app_cls.return_value = mock_app

        token = get_token(tenant_id="t", client_id="c", client_secret="s")

        assert token == "explicit-token"
