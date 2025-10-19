"""Utilities for generating narrative insights via OpenAI-compatible APIs."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Literal, Mapping, Sequence

import httpx

AIRequestMode = Literal["chat_completions", "responses"]


@dataclass(frozen=True)
class AIConfig:
    """Configuration for the AI helper."""

    api_key: str
    model: str
    api_base: str = "https://api.openai.com/v1"
    timeout: float = 30.0
    mode: AIRequestMode = "chat_completions"

    def headers(self) -> dict[str, str]:
        headers: dict[str, str] = {"Content-Type": "application/json"}
        if self.api_key:
            headers["Authorization"] = f"Bearer {self.api_key}"
        return headers


def _format_records(records: Sequence[Mapping[str, object]]) -> str:
    if not records:
        return "No financial data was provided."

    keys = list(records[0].keys())
    header = ",".join(keys)
    lines = [header]
    for item in records:
        row = []
        for key in keys:
            value = item.get(key, "")
            row.append(str(value))
        lines.append(",".join(row))
    return "\n".join(lines)


def generate_insights(
    records: Sequence[Mapping[str, object]],
    config: AIConfig,
    *,
    prompt: str | None = None,
    client: httpx.Client | None = None,
) -> str:
    """Send scenario data to an OpenAI-compatible endpoint and return insights."""
    if not config.api_key:
        raise ValueError("An API key is required to request AI insights.")

    if prompt is None:
        prompt = (
            "You are an FP&A analyst. Review the variance report data, highlight major "
            "drivers and noteworthy patterns, and suggest follow-up questions for the "
            "finance team."
        )

    dataset = _format_records(records)
    system_text = (
        "You provide concise but detailed narrative insights about financial "
        "performance based on tabular data."
    )
    user_text = f"{prompt}\n\nVariance data (CSV):\n{dataset}"

    if config.mode == "responses":
        path = "/responses"
        payload = {
            "model": config.model,
            "input": [
                {
                    "role": "system",
                    "content": [{"type": "text", "text": system_text}],
                },
                {
                    "role": "user",
                    "content": [{"type": "text", "text": user_text}],
                },
            ],
        }
    else:
        path = "/chat/completions"
        payload = {
            "model": config.model,
            "messages": [
                {"role": "system", "content": system_text},
                {"role": "user", "content": user_text},
            ],
        }

    if client is None:
        base_url = config.api_base.rstrip("/")
        with httpx.Client(base_url=base_url, timeout=config.timeout) as http_client:
            response = http_client.post(path, json=payload, headers=config.headers())
    else:
        response = client.post(path, json=payload, headers=config.headers())

    response.raise_for_status()
    data = response.json()

    if config.mode == "responses":
        content: str | None = None
        try:
            choice = data["choices"][0]
            message = choice.get("message", {})
            message_content = message.get("content")
        except (KeyError, IndexError, TypeError) as exc:  # pragma: no cover - safety net
            raise RuntimeError("Unexpected response from AI service") from exc

        if isinstance(message_content, str):
            content = message_content
        elif isinstance(message_content, list):
            text_parts: list[str] = []
            for item in message_content:
                if isinstance(item, dict) and item.get("type") == "text":
                    text_value = item.get("text")
                    if isinstance(text_value, str):
                        text_parts.append(text_value)
            if text_parts:
                content = "".join(text_parts)
        if content is None and isinstance(data.get("output_text"), list):
            text_value = data["output_text"][0]
            if isinstance(text_value, str):
                content = text_value
        if content is None:
            raise RuntimeError("AI response did not contain textual content")
        return content.strip()

    try:
        content = data["choices"][0]["message"]["content"]
    except (KeyError, IndexError, TypeError) as exc:  # pragma: no cover - safety net
        raise RuntimeError("Unexpected response from AI service") from exc

    if not isinstance(content, str):
        raise RuntimeError("AI response did not contain textual content")

    return content.strip()
