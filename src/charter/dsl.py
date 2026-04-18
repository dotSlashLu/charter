"""Pydantic models for the analysis DSL."""

from __future__ import annotations

from typing import Annotated, Literal, Union

from pydantic import BaseModel, Field


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
# Action models – each carries an optional `source` for nesting
# ---------------------------------------------------------------------------

class SelectAction(BaseModel):
    action: Literal["select"]
    columns: list[str]
    distinct: bool = False
    source: AnyAction | None = None


class FilterAction(BaseModel):
    action: Literal["filter"]
    conditions: list[FilterCondition]
    logic: Literal["AND", "OR"] = "AND"
    source: AnyAction | None = None


class GroupByAction(BaseModel):
    action: Literal["groupby"]
    rows: list[str]
    values: list[AggregationValue]
    source: AnyAction | None = None


class PivotAction(BaseModel):
    action: Literal["pivot"]
    rows: list[str]
    columns: list[str]
    values: list[AggregationValue]
    source: AnyAction | None = None


class SortAction(BaseModel):
    action: Literal["sort"]
    order_by: str = Field(alias="orderBy")
    direction: Literal["ASC", "DESC"] = "ASC"
    limit: int | None = None
    source: AnyAction | None = None

    model_config = {"populate_by_name": True}


# ---------------------------------------------------------------------------
# Discriminated union
# ---------------------------------------------------------------------------

AnyAction = Annotated[
    Union[SelectAction, FilterAction, GroupByAction, PivotAction, SortAction],
    Field(discriminator="action"),
]

# Rebuild forward-ref models now that AnyAction is defined.
SelectAction.model_rebuild()
FilterAction.model_rebuild()
GroupByAction.model_rebuild()
PivotAction.model_rebuild()
SortAction.model_rebuild()


# ---------------------------------------------------------------------------
# Top-level response wrapper
# ---------------------------------------------------------------------------

class AnalysisDimension(BaseModel):
    title: str
    description: str
    dsl: AnyAction


class AnalysisResponse(BaseModel):
    dimensions: list[AnalysisDimension]
