from __future__ import annotations

import argparse
import json
from datetime import date, time
from pathlib import Path

from .exporter import TravelExpenseExporter
from .models import ExportRequest, ExpenseLine, PassChange, TransitPart


def _line(value: dict) -> ExpenseLine:
    parts = tuple(TransitPart(**part) for part in value.get("transit_parts", []))
    return ExpenseLine(
        **{k: v for k, v in value.items() if k not in {"travel_date", "start_time", "transit_parts"}},
        travel_date=date.fromisoformat(value["travel_date"]),
        start_time=time.fromisoformat(value["start_time"]) if value.get("start_time") else None,
        transit_parts=parts,
    )


def main() -> int:
    parser = argparse.ArgumentParser(description="出張旅費代精算書をXLSXへ出力します")
    parser.add_argument("template", type=Path, help="2026年度版の原本（.ods/.xlsx）")
    parser.add_argument("data", type=Path, help="出力データJSON")
    parser.add_argument("--output", type=Path)
    parser.add_argument("--soffice", default="soffice")
    args = parser.parse_args()
    raw = json.loads(args.data.read_text(encoding="utf-8"))
    request = ExportRequest(
        year=raw["year"], month=raw["month"], department=raw["department"],
        employee_name=raw["employee_name"], commuter_pass=raw.get("commuter_pass"),
        submission_date=date.fromisoformat(raw["submission_date"]),
        lines=tuple(_line(item) for item in raw["lines"]),
        pass_changes=tuple(PassChange(date.fromisoformat(c["effective_date"]), c["new_section"]) for c in raw.get("pass_changes", [])),
        include_sample_sheet=raw.get("include_sample_sheet", False),
        output_dir=Path(raw.get("output_dir", ".")), revision=raw.get("revision"),
    )
    result = TravelExpenseExporter(args.soffice).export(args.template, request, args.output)
    print(json.dumps({"path": str(result.path), "details": result.detail_count, "pages": result.page_count,
                      "total": result.total, "page_subtotals": result.page_subtotals}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
