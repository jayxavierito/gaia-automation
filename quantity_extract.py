from __future__ import annotations

import argparse
import csv
import json
from dataclasses import asdict, fields
from pathlib import Path

from quantity_extractors import QuantityRecord, extract_quantity_source


def write_normalized_csv(path: Path, records: list[QuantityRecord]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = [field.name for field in fields(QuantityRecord)]
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for record in records:
            row = asdict(record)
            row["extraction_warnings"] = "; ".join(record.extraction_warnings)
            writer.writerow(row)


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "数量計算書(.xlsx/.xlsm/.pdf)をGAIA変換前の共通明細へ抽出します。"
        )
    )
    parser.add_argument("--source", required=True, type=Path)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--json", type=Path)
    parser.add_argument(
        "--work-dir",
        type=Path,
        help="PDFのページ・OCRセルを診断用に保存するディレクトリ。",
    )
    args = parser.parse_args()

    output_path = args.output or Path("outputs") / f"{args.source.stem}_normalized.csv"
    json_path = args.json or output_path.with_suffix(".json")
    result = extract_quantity_source(args.source, work_dir=args.work_dir)
    write_normalized_csv(output_path, result.records)
    json_path.parent.mkdir(parents=True, exist_ok=True)
    json_path.write_text(
        json.dumps(result.to_dict(), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    status_counts: dict[str, int] = {}
    for record in result.records:
        status_counts[record.extraction_status] = (
            status_counts.get(record.extraction_status, 0) + 1
        )
    summary = {
        "project_name": result.project_name,
        "source_format": result.source_format,
        "extractor": result.extractor,
        "records": len(result.records),
        "status_counts": status_counts,
        "document_warnings": result.document_warnings,
        "normalized_csv": str(output_path.resolve()),
        "extraction_json": str(json_path.resolve()),
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
