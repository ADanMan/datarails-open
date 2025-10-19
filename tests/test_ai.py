import json

import httpx
import pytest

from app.ai import AIConfig, generate_insights


def test_generate_insights_sends_structured_payload():
    captured: dict[str, object] = {}

    def handler(request: httpx.Request) -> httpx.Response:
        captured["method"] = request.method
        captured["url"] = str(request.url)
        captured["headers"] = dict(request.headers)
        captured["json"] = json.loads(request.content.decode("utf-8"))
        return httpx.Response(
            200,
            json={
                "choices": [
                    {
                        "message": {
                            "content": "Key variances include revenue softness in Sales."
                        }
                    }
                ]
            },
        )

    transport = httpx.MockTransport(handler)
    client = httpx.Client(base_url="https://mock.api/v1", transport=transport)

    records = [
        {
            "period": "2024-01",
            "department": "Sales",
            "account": "Revenue",
            "actual": 1000,
            "budget": 1200,
            "variance": -200,
        },
        {
            "period": "2024-01",
            "department": "Marketing",
            "account": "Expense",
            "actual": 500,
            "budget": 450,
            "variance": 50,
        },
    ]

    config = AIConfig(api_key="test-key", api_base="https://mock.api/v1", model="demo-model")
    insights = generate_insights(records, config, client=client)

    assert insights == "Key variances include revenue softness in Sales."
    assert captured["method"] == "POST"
    assert captured["url"].endswith("/chat/completions")
    assert captured["headers"].get("authorization") == "Bearer test-key"

    payload = captured["json"]
    assert payload["model"] == "demo-model"
    assert payload["messages"][1]["content"].startswith("You are an FP&A analyst")
    assert "period,department,account,actual,budget,variance" in payload["messages"][1]["content"]


def test_generate_insights_requires_api_key():
    config = AIConfig(api_key="", api_base="https://mock.api/v1", model="demo-model")
    with pytest.raises(ValueError):
        generate_insights([], config)


def test_generate_insights_raises_for_http_errors():
    def handler(request: httpx.Request) -> httpx.Response:
        return httpx.Response(500, json={"error": "server error"}, request=request)

    transport = httpx.MockTransport(handler)
    client = httpx.Client(base_url="https://mock.api/v1", transport=transport)
    config = AIConfig(api_key="test-key", api_base="https://mock.api/v1", model="demo-model")

    with pytest.raises(httpx.HTTPStatusError):
        generate_insights([], config, client=client)


def test_generate_insights_supports_responses_mode():
    captured: dict[str, object] = {}

    def handler(request: httpx.Request) -> httpx.Response:
        captured["method"] = request.method
        captured["url"] = str(request.url)
        captured["json"] = json.loads(request.content.decode("utf-8"))
        return httpx.Response(
            200,
            json={
                "choices": [
                    {
                        "message": {
                            "content": [
                                {"type": "text", "text": "Narrative from responses endpoint."}
                            ]
                        }
                    }
                ]
            },
        )

    transport = httpx.MockTransport(handler)
    client = httpx.Client(base_url="https://mock.api/v1", transport=transport)

    config = AIConfig(
        api_key="test-key",
        api_base="https://mock.api/v1",
        model="demo-model",
        mode="responses",
    )

    insights = generate_insights([], config, client=client)

    assert insights == "Narrative from responses endpoint."
    assert captured["url"].endswith("/responses")
    payload = captured["json"]
    assert payload["model"] == "demo-model"
    assert payload["input"][0]["role"] == "system"
