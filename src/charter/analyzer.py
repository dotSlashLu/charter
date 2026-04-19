"""Use OpenAI to recommend analysis dimensions for an Excel dataset."""

from __future__ import annotations

import json

from openai import OpenAI

from . import log
from .config import LLMConfig, load_config
from .dsl import AnalysisResponse
from .reader import SheetSchema, schema_to_prompt

_log = log.get("analyzer")


def analyze(schema: SheetSchema, *, cfg: LLMConfig | None = None) -> AnalysisResponse:
    """Call OpenAI-compatible API and return validated analysis dimensions."""
    if cfg is None:
        cfg = load_config()

    client_kwargs: dict = {}
    if cfg.api_key:
        client_kwargs["api_key"] = cfg.api_key
    if cfg.base_url:
        client_kwargs["base_url"] = cfg.base_url

    _log.debug(
        "Creating OpenAI client  model=%s  base_url=%s  api_key=%s",
        cfg.model,
        cfg.base_url or "(default)",
        "***" + cfg.api_key[-4:] if cfg.api_key else "(env)",
    )
    client = OpenAI(**client_kwargs)

    user_content = schema_to_prompt(schema)
    _log.debug("System prompt length: %d chars", len(cfg.system_prompt))
    _log.debug("User prompt length: %d chars", len(user_content))
    _log.debug("===== SYSTEM PROMPT BEGIN =====\n%s\n===== SYSTEM PROMPT END =====", cfg.system_prompt)
    _log.debug("===== USER PROMPT BEGIN =====\n%s\n===== USER PROMPT END =====", user_content)

    _log.debug("Sending chat completion request (temperature=%.2f)...", cfg.temperature)
    resp = client.chat.completions.create(
        model=cfg.model,
        messages=[
            {"role": "system", "content": cfg.system_prompt},
            {"role": "user", "content": user_content},
        ],
        response_format={"type": "json_object"},
        temperature=cfg.temperature,
    )

    usage = resp.usage
    if usage:
        _log.debug(
            "Token usage:  prompt=%d  completion=%d  total=%d",
            usage.prompt_tokens,
            usage.completion_tokens,
            usage.total_tokens,
        )

    raw = resp.choices[0].message.content or "{}"
    _log.debug("Raw LLM response length: %d chars", len(raw))
    _log.debug("===== LLM RAW RESPONSE BEGIN =====\n%s\n===== LLM RAW RESPONSE END =====", raw)

    data = json.loads(raw)
    result = AnalysisResponse.model_validate(data)
    _log.debug("Parsed %d analysis dimensions", len(result.dimensions))
    for dim in result.dimensions:
        actions = " -> ".join(s.action for s in dim.dsl.pipeline)
        _log.debug(
            "  [%s] %s  (pipeline=%s)",
            dim.title,
            dim.description,
            actions,
        )

    return result
