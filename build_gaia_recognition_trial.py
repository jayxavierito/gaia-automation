from __future__ import annotations

import argparse
from datetime import date
from pathlib import Path

from gaia_catalog import GaiaCatalogEntry, clean_text, load_gaia_catalog
from gaia_converter import (
    BoqItem,
    apply_gaia_catalog,
    build_gaia_candidate,
    classify_schema_item,
)


TRIAL_CASES = (
    (
        "CB・浅い階層",
        {
            "level1": "橋梁災害4057号",
            "level2": "道路土工",
            "level3": "作業土工",
            "item_name": "床掘り",
            "specification": "土砂 掘削深さ5m超20m以下 切梁腹起式 障害無し",
            "unit": "m3",
            "code": "CB210030",
            "item_column": 4,
        },
    ),
    (
        "WB・浅い階層",
        {
            "level1": "橋梁災害4057号",
            "level2": "構造物撤去工",
            "level3": "旧橋撤去工",
            "item_name": "構造物とりこわし",
            "specification": "鉄筋構造物 機械施工 制約無 夜間無 対策不要",
            "unit": "m3",
            "code": "WB824010",
            "item_column": 4,
        },
    ),
    (
        "CB・細別行＋規格行",
        {
            "level1": "6災4009号 港橋",
            "level2": "舗装",
            "level3": "舗装工",
            "level4": "アスファルト舗装工",
            "level5": "表層(車道・路肩部)",
            "item_name": "表層(車道・路肩部)",
            "specification": "3.0m超 60mm 密粒度アスコン(20F) タックコートPK-4",
            "unit": "m2",
            "code": "CB410260",
            "item_column": 6,
        },
    ),
    (
        "明細・名称のみ",
        {
            "level1": "道路災害3911号",
            "level2": "舗装工",
            "level3": "舗装打換え工",
            "item_name": "舗装版破砕",
            "specification": "",
            "unit": "m2",
            "code": "",
            "item_column": 4,
            "reference_type": "明細",
        },
    ),
)


def find_trial_entry(
    catalog: list[GaiaCatalogEntry], selector: dict[str, object]
) -> GaiaCatalogEntry:
    def entry_matches(entry: GaiaCatalogEntry) -> bool:
        for name, expected in selector.items():
            actual = getattr(entry, name)
            if isinstance(expected, str):
                if clean_text(actual) != clean_text(expected):
                    return False
            elif actual != expected:
                return False
        return True

    matches = [
        entry
        for entry in catalog
        if entry_matches(entry)
    ]
    if len(matches) != 1:
        raise ValueError(
            f"試験項目を一意に取得できません: {selector} ({len(matches)}件)"
        )
    return matches[0]


def make_trial_item(entry: GaiaCatalogEntry, label: str, order: int) -> BoqItem:
    item = BoqItem(
        source_sheet="GAIA認識トリガー比較",
        source_row=order,
        section=entry.level1,
        work_type=entry.level2,
        category=entry.level3,
        item_name=entry.item_name,
        specification=entry.specification,
        unit=entry.unit,
        quantity=1,
        remarks=label,
        standard="",
        daily_output="",
        setting="",
        notes="",
        source_reference="",
    )
    apply_gaia_catalog([item], [entry])
    classify_schema_item(item)
    return item


def main() -> int:
    root = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser(
        description="GAIAが明細・施工表を再生成する条件を確認する金抜試験Excelを作成します。"
    )
    parser.add_argument(
        "--template",
        type=Path,
        default=root / "assets" / "gaia_suzu_7sheet_template.xlsx",
    )
    parser.add_argument(
        "--catalog",
        type=Path,
        default=root / "assets" / "gaia_import_catalog.json",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=root / "outputs" / "GAIA認識トリガー比較_金抜.xlsx",
    )
    args = parser.parse_args()

    catalog = load_gaia_catalog(args.catalog)
    items = [
        make_trial_item(find_trial_entry(catalog, selector), label, i)
        for i, (label, selector) in enumerate(TRIAL_CASES, start=1)
    ]
    block_count = build_gaia_candidate(
        template_path=args.template,
        output_path=args.output,
        project_name="GAIA認識トリガー比較テスト",
        items=items,
        location="珠洲市内（テスト）",
        work_category="舗装工事",
        district="珠洲",
        price_date=date.today().replace(day=1),
    )
    print(f"{block_count} blocks -> {args.output.resolve()}")
    for label, item in zip((case[0] for case in TRIAL_CASES), items):
        path = " > ".join(
            value
            for value in (
                item.output_level1,
                item.output_level2,
                item.output_level3,
                item.output_subcategory,
                item.output_specification_group,
                item.gaia_item_name,
            )
            if value
        )
        print(f"- {label}: {path} [{item.gaia_code or '名称のみ'}]")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
