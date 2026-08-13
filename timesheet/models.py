"""Modelos pydantic del estado del timesheet."""

import uuid
from datetime import date, time
from typing import Annotated, Any

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator


class _Base(BaseModel):
    model_config = ConfigDict(
        extra="forbid",
        validate_assignment=True,
        populate_by_name=True,
    )


class Professional(_Base):
    name: str = Field(..., min_length=1)
    specialty: str = Field(..., min_length=1)
    hourly_rate: float = Field(..., ge=0)

    @field_validator("name", "specialty", mode="before")
    @classmethod
    def _strip(cls, v: Any) -> Any:
        if isinstance(v, str):
            v = v.strip()
        return v

    @field_validator("hourly_rate", mode="after")
    @classmethod
    def _round_rate(cls, v: float) -> float:
        return round(v, 2)


class Supervisor(_Base):
    name: str = Field(..., min_length=1)
    sup_date: date | None = Field(default=None, alias="date")

    @field_validator("name", mode="before")
    @classmethod
    def _strip(cls, v: Any) -> Any:
        if isinstance(v, str):
            v = v.strip()
        return v


class DayItem(_Base):
    id: str = Field(default_factory=lambda: str(uuid.uuid4()))
    description: str = Field(..., min_length=1)
    hours: float = Field(..., ge=0)

    @field_validator("description", mode="before")
    @classmethod
    def _strip_desc(cls, v: Any) -> Any:
        if not isinstance(v, str):
            raise ValueError("description debe ser string")
        v = v.strip()
        if not v:
            raise ValueError("description no puede quedar vacio tras strip")
        return v

    @field_validator("hours", mode="after")
    @classmethod
    def _round_hours(cls, v: float) -> float:
        return round(v, 2)


class Day(_Base):
    entry: time | None = None
    exit: time | None = None
    items: list[DayItem] = Field(default_factory=list)

    @field_validator("items", mode="before")
    @classmethod
    def _accept_null_items(cls, v: Any) -> Any:
        if v is None:
            return []
        return v


class State(_Base):
    schema_version: int = 1
    year: int = Field(..., ge=2000, le=2100)
    month: int = Field(..., ge=1, le=12)
    professional: Professional
    supervisor: Supervisor
    days: dict[str, Day] = Field(default_factory=dict)

    @field_validator("days", mode="before")
    @classmethod
    def _accept_null_days(cls, v: Any) -> Any:
        if v is None:
            return {}
        return v

    @model_validator(mode="after")
    def _check_keys(self) -> "State":
        for k in self.days:
            try:
                date.fromisoformat(k)
            except ValueError as exc:
                raise ValueError(f"clave de dia invalida: {k!r}") from exc
        return self


Professional.model_rebuild()
Supervisor.model_rebuild()
DayItem.model_rebuild()
Day.model_rebuild()
State.model_rebuild()


__all__ = [
    "Professional",
    "Supervisor",
    "DayItem",
    "Day",
    "State",
]
