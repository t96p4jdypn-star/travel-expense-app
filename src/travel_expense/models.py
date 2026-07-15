from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, time
from pathlib import Path
from typing import Literal

ReviewState = Literal["確認済み", "修正済み", "未確認", "保留"]


@dataclass(frozen=True)
class TransitPart:
    mode: str
    origin: str
    destination: str
    amount: int

    def label(self) -> str:
        return f"（{self.mode}）{self.origin}→{self.destination} {self.amount}円"


@dataclass(frozen=True)
class ExpenseLine:
    id: str
    travel_date: date
    destination: str
    paid_section: str
    amount: int
    reason: str
    review_state: ReviewState = "確認済み"
    excluded: bool = False
    submitted: bool = False
    duplicate_unresolved: bool = False
    amount_confirmed: bool = True
    private_or_online: bool = False
    start_time: time | None = None
    route_order: int = 0
    created_order: int = 0
    transit_parts: tuple[TransitPart, ...] = ()

    def is_exportable(self, year: int, month: int) -> bool:
        return (
            self.travel_date.year == year
            and self.travel_date.month == month
            and self.review_state in {"確認済み", "修正済み"}
            and self.amount >= 1
            and not self.excluded
            and not self.submitted
            and not self.duplicate_unresolved
            and self.amount_confirmed
            and not self.private_or_online
        )

    def section_text(self) -> str:
        if self.transit_parts:
            return "\n".join(part.label() for part in self.transit_parts)
        return self.paid_section


@dataclass(frozen=True)
class PassChange:
    effective_date: date
    new_section: str


@dataclass(frozen=True)
class ExportRequest:
    year: int
    month: int
    department: str
    employee_name: str
    commuter_pass: str | None
    submission_date: date
    lines: tuple[ExpenseLine, ...]
    pass_changes: tuple[PassChange, ...] = ()
    include_sample_sheet: bool = False
    template_sheet: str = "【原本】出張旅費精算"
    output_dir: Path = field(default_factory=Path.cwd)
    revision: int | None = None
