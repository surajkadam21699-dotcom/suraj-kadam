"""Offline stand-ins for HTTP so the tests never touch the network.

The fixture pages and API payloads under tests/fixtures hold synthetic figures;
only their layout mirrors the real sites and the UN Comtrade API.
"""

from pathlib import Path

import pytest
import requests

FIXTURES = Path(__file__).parent / "fixtures"


class FakeResponse:
    def __init__(self, payload=None, *, content: bytes = b"", status: int = 200):
        self._payload = payload
        self.content = content
        self.text = content.decode("utf-8", "replace")
        self.status_code = status

    def json(self):
        if self._payload is None:
            raise ValueError("response is not JSON")
        return self._payload

    def raise_for_status(self):
        if self.status_code >= 400:
            raise requests.HTTPError(f"{self.status_code} Client Error")


class FakeSession:
    """Answers every GET with ``respond(url, params)`` and records the calls."""

    def __init__(self, respond):
        self.respond = respond
        self.calls = []

    def get(self, url, params=None, timeout=None):
        self.calls.append((url, dict(params or {})))
        return self.respond(url, dict(params or {}))


@pytest.fixture
def fake_session():
    return FakeSession


@pytest.fixture
def fake_response():
    return FakeResponse


@pytest.fixture
def fixtures_dir():
    return FIXTURES
