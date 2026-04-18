"""Configuration management – reads charter.toml for API keys, base URL, model, and prompt."""

from __future__ import annotations

import tomllib
from dataclasses import dataclass
from pathlib import Path

DEFAULT_CONFIG_NAME = "charter.toml"

DEFAULT_SYSTEM_PROMPT = """\
You are a data-analysis expert.  The user will give you the schema and sample
data of an Excel sheet.  Your job is to recommend 3-6 **meaningful, diverse**
analysis dimensions.  Each dimension is described by a JSON DSL object.

# DSL actions

1. **select** – extract / project columns
   - columns: string[]
   - distinct: boolean (default false)

2. **filter** – keep rows matching conditions
   - conditions: array of {column, operator ("=",">","<",">=","<=","!="), value}
   - logic: "AND" | "OR"

3. **groupby** – aggregate rows
   - rows: string[]   (grouping columns)
   - values: array of {column, aggregation: "SUM"|"AVERAGE"|"COUNT"|"MAX"|"MIN"}

4. **pivot** – cross-tabulation
   - rows: string[]
   - columns: string[]
   - values: array of {column, aggregation}

5. **sort** – sort (optionally top-N)
   - orderBy: string (column name)
   - direction: "ASC" | "DESC"
   - limit: number | null
   - source: nested DSL object (usually groupby or filter)

Every action may optionally carry a **source** field containing another DSL
object, representing a nested / chained data source.  If source is omitted the
raw table is used.

# Guidelines
- Cover a **variety** of action types – do not just produce 6 groupby queries.
- Make sure column names and values match the real data exactly.
- Prefer analyses that reveal useful business / statistical insights.
- The "title" should be a short name suitable as an Excel sheet tab (≤31 chars).
- The "description" should be one sentence explaining the insight.

# Response format
Return a JSON object:
{
  "dimensions": [
    {"title": "...", "description": "...", "dsl": { ... }},
    ...
  ]
}
"""

DEFAULT_CONFIG_TOML = '''\
# Charter configuration file
# Docs: https://github.com/example/charter

# Enable verbose debug output (or use --verbose on the CLI)
# debug = true

[llm]
# OpenAI-compatible API settings
# api_key = "sk-..."          # or set OPENAI_API_KEY env var
# base_url = "https://api.openai.com/v1"  # change for compatible providers
model = "gpt-4o"
temperature = 0.4

[llm.prompt]
# System prompt sent to the LLM.  Edit freely to customise analysis style.
system = """
You are a data-analysis expert.  The user will give you the schema and sample
data of an Excel sheet.  Your job is to recommend 3-6 **meaningful, diverse**
analysis dimensions.  Each dimension is described by a JSON DSL object.

# DSL actions

1. **select** – extract / project columns
   - columns: string[]
   - distinct: boolean (default false)

2. **filter** – keep rows matching conditions
   - conditions: array of {column, operator ("=",">","<",">=","<=","!="), value}
   - logic: "AND" | "OR"

3. **groupby** – aggregate rows
   - rows: string[]   (grouping columns)
   - values: array of {column, aggregation: "SUM"|"AVERAGE"|"COUNT"|"MAX"|"MIN"}

4. **pivot** – cross-tabulation
   - rows: string[]
   - columns: string[]
   - values: array of {column, aggregation}

5. **sort** – sort (optionally top-N)
   - orderBy: string (column name)
   - direction: "ASC" | "DESC"
   - limit: number | null
   - source: nested DSL object (usually groupby or filter)

Every action may optionally carry a **source** field containing another DSL
object, representing a nested / chained data source.  If source is omitted the
raw table is used.

# Guidelines
- Cover a **variety** of action types – do not just produce 6 groupby queries.
- Make sure column names and values match the real data exactly.
- Prefer analyses that reveal useful business / statistical insights.
- The "title" should be a short name suitable as an Excel sheet tab (≤31 chars).
- The "description" should be one sentence explaining the insight.

# Response format
Return a JSON object:
{
  "dimensions": [
    {"title": "...", "description": "...", "dsl": { ... }},
    ...
  ]
}
"""
'''


@dataclass
class LLMConfig:
    api_key: str | None
    base_url: str | None
    model: str
    temperature: float
    system_prompt: str
    debug: bool


def load_config(path: str | Path | None = None) -> LLMConfig:
    """Load configuration from a TOML file.

    Resolution order for the config file:
      1. Explicit *path* argument
      2. ``./charter.toml`` in the current working directory
      3. ``~/.config/charter/charter.toml``

    If no file is found, built-in defaults are used.
    """
    cfg_path = _resolve_path(path)

    if cfg_path is not None:
        with open(cfg_path, "rb") as f:
            raw = tomllib.load(f)
    else:
        raw = {}

    llm = raw.get("llm", {})
    prompt_section = llm.get("prompt", {})

    return LLMConfig(
        api_key=llm.get("api_key"),
        base_url=llm.get("base_url"),
        model=llm.get("model", "gpt-4o"),
        temperature=llm.get("temperature", 0.4),
        system_prompt=prompt_section.get("system", DEFAULT_SYSTEM_PROMPT).strip(),
        debug=raw.get("debug", False),
    )


def init_config(path: str | Path | None = None) -> Path:
    """Write the default config template and return its path."""
    dest = Path(path) if path else Path.cwd() / DEFAULT_CONFIG_NAME
    dest.write_text(DEFAULT_CONFIG_TOML, encoding="utf-8")
    return dest


def _resolve_path(explicit: str | Path | None) -> Path | None:
    if explicit is not None:
        p = Path(explicit)
        if p.is_file():
            return p
        raise FileNotFoundError(f"Config file not found: {p}")

    cwd_candidate = Path.cwd() / DEFAULT_CONFIG_NAME
    if cwd_candidate.is_file():
        return cwd_candidate

    xdg = Path.home() / ".config" / "charter" / DEFAULT_CONFIG_NAME
    if xdg.is_file():
        return xdg

    return None
