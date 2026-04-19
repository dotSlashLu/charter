"""Pydantic models for the analysis DSL."""

from __future__ import annotations

from typing import Annotated, Any, Literal, Union

from pydantic import BaseModel, Field, model_validator


# ---------------------------------------------------------------------------
# Leaf / helper types
# ---------------------------------------------------------------------------

class FilterCondition(BaseModel):
    column: str
    operator: Literal["=", ">", "<", ">=", "<=", "!="]
    value: str | int | float


class AggregationValue(BaseModel):
    column: str
    aggregation: Literal["SUM", "AVERAGE", "COUNT", "MAX", "MIN"]


# ---------------------------------------------------------------------------
# Action models — no `source` field; sequencing is handled by the pipeline
# ---------------------------------------------------------------------------

class SelectAction(BaseModel):
    action: Literal["select"]
    columns: list[str]
    distinct: bool = False


class FilterAction(BaseModel):
    action: Literal["filter"]
    conditions: list[FilterCondition]
    logic: Literal["AND", "OR"] = "AND"


class GroupByAction(BaseModel):
    action: Literal["groupby"]
    rows: list[str]
    values: list[AggregationValue]


class PivotAction(BaseModel):
    action: Literal["pivot"]
    rows: list[str]
    columns: list[str]
    values: list[AggregationValue]


class SortAction(BaseModel):
    action: Literal["sort"]
    order_by: str = Field(alias="orderBy")
    direction: Literal["ASC", "DESC"] = "ASC"
    limit: int | None = None

    model_config = {"populate_by_name": True}


# ---------------------------------------------------------------------------
# Discriminated union
# ---------------------------------------------------------------------------

AnyAction = Annotated[
    Union[SelectAction, FilterAction, GroupByAction, PivotAction, SortAction],
    Field(discriminator="action"),
]


# ---------------------------------------------------------------------------
# Pipeline wrapper
# ---------------------------------------------------------------------------

class DSLSpec(BaseModel):
    """A sequential pipeline of actions.

    Accepts either ``{"pipeline": [...]}`` or a bare action dict like
    ``{"action": "groupby", ...}`` (auto-wrapped into a length-1 pipeline).
    """

    pipeline: list[AnyAction] = Field(min_length=1)

    @model_validator(mode="before")
    @classmethod
    def _wrap_bare_action(cls, data: Any) -> Any:
        if isinstance(data, dict) and "action" in data and "pipeline" not in data:
            return {"pipeline": [data]}
        return data


# ---------------------------------------------------------------------------
# Top-level response wrapper
# ---------------------------------------------------------------------------

class AnalysisDimension(BaseModel):
    title: str
    description: str
    dsl: DSLSpec


class AnalysisResponse(BaseModel):
    dimensions: list[AnalysisDimension]
