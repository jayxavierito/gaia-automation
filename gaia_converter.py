from __future__ import annotations

import argparse
import csv
import json
import re
import unicodedata
from collections import defaultdict
from dataclasses import asdict, dataclass, field
from datetime import date, datetime
from difflib import SequenceMatcher
from pathlib import Path
from typing import Iterable

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.worksheet.worksheet import Worksheet
from pypdf import PdfReader

from gaia_catalog import (
    GaiaCatalogEntry,
    load_gaia_catalog,
    match_gaia_catalog,
)
from quantity_extractors import ExtractionResult, extract_quantity_source


BLANK_VALUES = {"", "-", "―"}
SOURCE_YEAR_RE = re.compile(r"(?:令和|R)\s*(\d+)", re.IGNORECASE)
GAIA_CODE_RE = re.compile(r"^\[([A-Z]{2}\d{6})\]$")
ESTIMATE_DATE_LABELS = (
    "積算年月",
    "単価適用年月",
    "歩掛単価適用年月",
    "歩掛り単価適用年月",
    "単価年月",
)
ERA_YEAR_MONTH_RE = re.compile(
    r"(?:令和|R)\s*(\d{1,2})\s*(?:年|[./-])?\s*(\d{1,2})\s*月?",
    re.IGNORECASE,
)
WESTERN_YEAR_MONTH_RE = re.compile(
    r"\b(20\d{2})\s*(?:年|[./-])\s*(\d{1,2})\s*月?"
)
HIERARCHY_NUMBER_PREFIX_RES = (
    re.compile(r"^\s*[（(]\s*\d+(?:[.-]\d+)*\s*[）)]\s*"),
    re.compile(r"^\s*(?:§\s*)?\d+(?:[.-]\d+)*[.．、:：)]\s*"),
    re.compile(r"^\s*\d+\s+(?=\S)"),
)
BRIDGE_CONTEXT_TERMS = (
    "橋",
    "橋梁",
    "橋台",
    "橋脚",
    "胸壁",
    "地覆",
    "伸縮装置",
    "支承",
    "床版",
    "主桁",
)


@dataclass
class MappingRule:
    priority: int
    item_pattern: str
    spec_pattern: str = ""
    gaia_item_name: str = ""
    gaia_code: str = ""
    condition_append: str = ""
    notes: str = ""


@dataclass(frozen=True)
class GaiaCodeHistoryEntry:
    code: str
    name: str
    condition: str
    unit: str
    source_file: str
    source_sheet: str
    source_row: int


@dataclass
class BoqItem:
    source_sheet: str
    source_row: int
    section: str
    work_type: str
    category: str
    item_name: str
    specification: str
    unit: str
    quantity: float | int | None
    remarks: str
    standard: str
    daily_output: str
    setting: str
    notes: str
    source_reference: str
    source_format: str = "excel"
    extractor: str = ""
    source_page: int | None = None
    extraction_status: str = "READY"
    extraction_confidence: float = 1.0
    extraction_warnings: list[str] = field(default_factory=list)
    gaia_item_name: str = ""
    gaia_code: str = ""
    gaia_condition: str = ""
    match_status: str = ""
    confidence: float = 0.0
    matched_rule: str = ""
    tree_category: str = ""
    tree_status: str = ""
    tree_pages: str = ""
    tree_level1: str = ""
    tree_level2: str = ""
    tree_level3: str = ""
    tree_level4: str = ""
    tree_match_score: float = 0.0
    guideline_status: str = ""
    guideline_pages: str = ""
    package_match_type: str = ""
    package_name: str = ""
    package_no: str = ""
    package_unit: str = ""
    package_source: str = ""
    package_source_page: str = ""
    package_fiscal_year: int | None = None
    package_effective_from: str = ""
    package_scope: str = ""
    package_condition_fields: str = ""
    package_condition_status: str = ""
    package_unit_status: str = ""
    reference_date_status: str = ""
    source_standard_year: int | None = None
    source_standard_status: str = ""
    package_suggestions: str = ""
    output_level1: str = ""
    output_level2: str = ""
    output_level3: str = ""
    output_level4: str = ""
    catalog_status: str = ""
    catalog_score: float = 0.0
    catalog_name_score: float = 0.0
    catalog_specification_score: float = 0.0
    catalog_candidate_count: int = 0
    catalog_path_safe: bool = False
    catalog_code_safe: bool = False
    catalog_source_row: int | None = None
    warnings: list[str] = field(default_factory=list)


def clean_text(value: object) -> str:
    if value is None:
        return ""
    text = unicodedata.normalize("NFKC", str(value))
    return re.sub(r"\s+", " ", text).strip()


def normalize_hierarchy_label(value: object) -> str:
    text = clean_text(value)
    for pattern in HIERARCHY_NUMBER_PREFIX_RES:
        cleaned = pattern.sub("", text)
        if cleaned != text:
            return cleaned.strip()
    return text


def normalize_match_text(value: object) -> str:
    text = clean_text(value).replace("コンクリ-ト", "コンクリート")
    text = text.replace("積込み", "積込")
    normalized = re.sub(
        r"[\s・･()（）「」【】\[\]‐‑‒–—―-]", "", text
    ).lower()
    return normalized[:-1] if normalized.endswith("工") else normalized


def normalize_unit(value: object) -> str:
    text = clean_text(value).lower()
    replacements = {
        "m²": "m2",
        "m³": "m3",
        "㎡": "m2",
        "㎥": "m3",
        "ton": "t",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    return text.replace(" ", "")


def compare_units(source_unit: object, package_unit: object) -> str:
    source = normalize_unit(source_unit)
    package = normalize_unit(package_unit)
    if not package:
        return "REFERENCE_UNIT_MISSING"
    if source == package:
        return "MATCH"
    equivalent_pairs = {
        ("孔", "箇所"),
        ("箇所", "孔"),
        ("掛m2", "m2"),
        ("m2", "掛m2"),
    }
    return "EQUIVALENT" if (source, package) in equivalent_pairs else "MISMATCH"


def source_fiscal_year(value: object) -> int | None:
    match = SOURCE_YEAR_RE.search(clean_text(value))
    return 2018 + int(match.group(1)) if match else None


def fiscal_year_for_date(value: date) -> int:
    return value.year if value.month >= 4 else value.year - 1


def _year_month_from_value(value: object) -> date | None:
    if isinstance(value, datetime):
        return date(value.year, value.month, 1)
    if isinstance(value, date):
        return date(value.year, value.month, 1)
    text = clean_text(value)
    if not text:
        return None
    match = ERA_YEAR_MONTH_RE.search(text)
    if match:
        return date(2018 + int(match.group(1)), int(match.group(2)), 1)
    match = WESTERN_YEAR_MONTH_RE.search(text)
    if match:
        return date(int(match.group(1)), int(match.group(2)), 1)
    return None


def detect_estimate_date(source_path: Path) -> tuple[date, str]:
    suffix = source_path.suffix.lower()
    if suffix in {".xlsx", ".xlsm"}:
        workbook = load_workbook(
            source_path,
            data_only=True,
            read_only=True,
        )
        try:
            for sheet in workbook.worksheets:
                max_row = min(sheet.max_row, 150)
                max_column = min(sheet.max_column, 40)
                for row in sheet.iter_rows(
                    min_row=1,
                    max_row=max_row,
                    min_col=1,
                    max_col=max_column,
                    values_only=True,
                ):
                    for column, value in enumerate(row):
                        text = clean_text(value)
                        if not any(label in text for label in ESTIMATE_DATE_LABELS):
                            continue
                        for candidate in row[column : column + 5]:
                            detected = _year_month_from_value(candidate)
                            if detected:
                                return detected, "SOURCE"
        finally:
            workbook.close()
    elif suffix == ".pdf":
        reader = PdfReader(source_path)
        text = "\n".join(
            page.extract_text() or "" for page in reader.pages[:15]
        )
        for line in text.splitlines():
            if not any(label in line for label in ESTIMATE_DATE_LABELS):
                continue
            detected = _year_month_from_value(line)
            if detected:
                return detected, "SOURCE"
    return date.today(), "PC_DATE"


def append_warning(item: BoqItem, message: str) -> None:
    cleaned = clean_text(message)
    if cleaned and cleaned not in item.warnings:
        item.warnings.append(cleaned)


def combine_unique(*values: str) -> str:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        cleaned = clean_text(value)
        if not cleaned or cleaned in {"-", "―"} or cleaned in seen:
            continue
        seen.add(cleaned)
        result.append(cleaned)
    return " / ".join(result)


def boq_items_from_extraction(extraction: ExtractionResult) -> list[BoqItem]:
    items: list[BoqItem] = []
    for record in extraction.records:
        item = BoqItem(
            source_sheet=record.source_sheet,
            source_row=record.source_row,
            section=normalize_hierarchy_label(record.section),
            work_type=normalize_hierarchy_label(record.work_type),
            category=normalize_hierarchy_label(record.category),
            item_name=record.item_name,
            specification=record.specification,
            unit=record.unit,
            quantity=record.quantity,
            remarks=record.remarks,
            standard=record.standard,
            daily_output=record.daily_output,
            setting=record.setting,
            notes=record.notes,
            source_reference=record.source_reference,
            source_format=record.source_format,
            extractor=record.extractor,
            source_page=record.source_page,
            extraction_status=record.extraction_status,
            extraction_confidence=record.extraction_confidence,
            extraction_warnings=list(record.extraction_warnings),
        )
        item.gaia_item_name = item.item_name
        item.gaia_condition = combine_unique(item.specification, item.setting)
        for warning in item.extraction_warnings:
            append_warning(item, warning)
        items.append(item)
    return items


def apply_gaia_catalog(
    items: Iterable[BoqItem], catalog: Iterable[GaiaCatalogEntry]
) -> None:
    item_list = list(items)
    catalog_entries = list(catalog)
    for item in item_list:
        item.output_level1 = normalize_hierarchy_label(item.section) or "本工事費"
        item.output_level2 = normalize_hierarchy_label(item.work_type) or "工種未確認"
        item.output_level3 = normalize_hierarchy_label(item.category) or "種別未確認"
        item.output_level4 = clean_text(item.item_name)
        item.gaia_item_name = item.output_level4
        item.gaia_condition = combine_unique(item.specification, item.setting)

        match = match_gaia_catalog(
            catalog_entries,
            item_name=item.item_name,
            specification=item.gaia_condition,
            unit=item.unit,
            work_type=item.work_type,
            category=item.category,
        )
        item.catalog_status = match.status
        item.catalog_score = match.score
        item.catalog_name_score = match.name_score
        item.catalog_specification_score = match.specification_score
        item.catalog_candidate_count = match.candidate_count
        item.catalog_path_safe = match.path_safe
        item.catalog_code_safe = match.code_safe
        if match.entry is None:
            continue

        item.catalog_source_row = match.entry.source_row
        if match.status in {"EXACT_CODE", "EXACT_PATH", "FUZZY_PATH"}:
            item.output_level4 = match.entry.item_name
            item.gaia_item_name = match.entry.item_name
        if match.path_safe and not clean_text(item.work_type):
            item.output_level2 = match.entry.level2 or item.output_level2
        if match.path_safe and not clean_text(item.category):
            item.output_level3 = match.entry.level3 or item.output_level3
        if (
            match.code_safe
            and item.extraction_status == "READY"
            and item.quantity is not None
            and item.quantity > 0
        ):
            item.gaia_code = match.entry.code
        item.matched_rule = (
            f"GAIA_CATALOG:{match.entry.source_file}:"
            f"本工事費内訳表!{match.entry.source_row}"
        )

    # A confirmed item path is useful evidence for unmatched siblings in the
    # same source 工種/種別 group, but never supplies a GAIA code.
    grouped: dict[tuple[str, str, str, str], list[BoqItem]] = defaultdict(list)
    for item in item_list:
        work_key = normalize_match_text(item.work_type)
        category_key = normalize_match_text(item.category)
        if not work_key and not category_key:
            continue
        grouped[
            (
                clean_text(item.source_sheet),
                normalize_match_text(item.section),
                work_key,
                category_key,
            )
        ].append(item)
    for group_items in grouped.values():
        confirmed_paths = {
            (item.output_level2, item.output_level3)
            for item in group_items
            if item.catalog_path_safe
        }
        if len(confirmed_paths) != 1:
            continue
        level2, level3 = next(iter(confirmed_paths))
        for item in group_items:
            if item.catalog_path_safe:
                continue
            if not clean_text(item.work_type):
                item.output_level2 = level2
            if not clean_text(item.category):
                item.output_level3 = level3
            if item.catalog_status in {"NOT_FOUND", "AMBIGUOUS", "EXACT_AMBIGUOUS"}:
                item.catalog_status = f"{item.catalog_status}_GROUP_PATH"


def classify_schema_item(item: BoqItem) -> None:
    if item.quantity is None or item.quantity <= 0:
        item.match_status = "INVALID_QUANTITY"
        item.confidence = 0.0
        append_warning(item, "数量を取得できない、または数量が0以下です")
    elif item.extraction_status != "READY":
        item.match_status = "SOURCE_REVIEW_REQUIRED"
        item.confidence = min(0.5, item.extraction_confidence)
        append_warning(item, "原本照合が完了するまでGAIAコードを確定しません")
    elif item.gaia_code and item.catalog_code_safe:
        item.match_status = "CATALOG_EXACT_CODE"
        item.confidence = 0.98
    elif item.catalog_status in {"EXACT_PATH", "EXACT_CODE"}:
        item.match_status = "GAIA_NAME_SEARCH"
        item.confidence = 0.9
    elif item.catalog_status == "FUZZY_PATH":
        item.match_status = "GAIA_NAME_SEARCH_REVIEW"
        item.confidence = 0.65
    elif item.catalog_status.endswith("_GROUP_PATH"):
        item.match_status = "GAIA_NAME_SEARCH_GROUP"
        item.confidence = 0.55
    else:
        item.match_status = "GAIA_NAME_SEARCH_REQUIRED"
        item.confidence = 0.35


def infer_work_category(project_name: str, items: Iterable[BoqItem]) -> str:
    scores = {
        "橋梁工事": 0,
        "下水道工事（２）": 0,
        "道路工事": 0,
    }
    terms = {
        "橋梁工事": ("橋", "橋梁", "支承", "伸縮装置", "地覆", "胸壁", "床版", "主桁"),
        "下水道工事（２）": ("下水", "管渠", "マンホール", "汚水", "雨水", "推進工"),
        "道路工事": ("道路", "舗装", "路盤", "区画線", "側溝"),
    }
    project_text = clean_text(project_name)
    for category, keywords in terms.items():
        if any(keyword in project_text for keyword in keywords):
            scores[category] += 3
    for item in items:
        context = combine_unique(
            item.work_type,
            item.category,
            item.item_name,
            item.specification,
        )
        for category, keywords in terms.items():
            if any(keyword in context for keyword in keywords):
                scores[category] += 1
    selected, score = max(scores.items(), key=lambda pair: pair[1])
    return selected if score >= 2 else ""


def extract_boq_items(
    source_path: Path, work_dir: Path | None = None
) -> tuple[str, list[BoqItem]]:
    extraction = extract_quantity_source(source_path, work_dir=work_dir)
    items = boq_items_from_extraction(extraction)
    return extraction.project_name, items


def load_mapping_rules(path: Path | None) -> list[MappingRule]:
    if path is None or not path.exists():
        return []
    rules: list[MappingRule] = []
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        for row in csv.DictReader(handle):
            if not clean_text(row.get("item_pattern")):
                continue
            rules.append(
                MappingRule(
                    priority=int(row.get("priority") or 0),
                    item_pattern=row["item_pattern"],
                    spec_pattern=row.get("spec_pattern", ""),
                    gaia_item_name=row.get("gaia_item_name", ""),
                    gaia_code=row.get("gaia_code", ""),
                    condition_append=row.get("condition_append", ""),
                    notes=row.get("notes", ""),
                )
            )
    return sorted(rules, key=lambda rule: rule.priority, reverse=True)


def apply_mapping_rules(items: Iterable[BoqItem], rules: list[MappingRule]) -> None:
    for item in items:
        normalized_name = re.sub(r"\s+", "", clean_text(item.item_name))
        normalized_spec = re.sub(r"\s+", "", clean_text(item.specification))
        for rule in rules:
            if not re.search(rule.item_pattern, normalized_name, flags=re.IGNORECASE):
                continue
            if rule.spec_pattern and not re.search(
                rule.spec_pattern, normalized_spec, flags=re.IGNORECASE
            ):
                continue
            if rule.gaia_item_name:
                item.gaia_item_name = clean_text(rule.gaia_item_name)
            if rule.gaia_code:
                item.gaia_code = clean_text(rule.gaia_code)
            if rule.condition_append:
                item.gaia_condition = combine_unique(
                    item.gaia_condition, rule.condition_append
                )
            item.matched_rule = f"{rule.item_pattern} / {rule.spec_pattern}"
            if rule.notes:
                append_warning(item, rule.notes)
            break


def extract_gaia_code_history(path: Path) -> list[GaiaCodeHistoryEntry]:
    """Read trusted GAIA codes from an already accepted 金抜 workbook."""
    workbook = load_workbook(path, data_only=False, read_only=True)
    if "本工事費内訳表" not in workbook.sheetnames:
        workbook.close()
        raise ValueError(f"コード履歴に本工事費内訳表がありません: {path}")

    sheet = workbook["本工事費内訳表"]
    rows = list(sheet.iter_rows(min_col=1, max_col=13, values_only=True))
    entries: list[GaiaCodeHistoryEntry] = []
    for index, values in enumerate(rows):
        code_match = GAIA_CODE_RE.fullmatch(clean_text(values[12]))
        if not code_match:
            continue

        name = next(
            (clean_text(value) for value in reversed(values[:5]) if clean_text(value)),
            "",
        )
        next_values = rows[index + 1] if index + 1 < len(rows) else ()
        condition = next(
            (
                clean_text(value)
                for value in reversed(next_values[:5])
                if clean_text(value)
            ),
            "",
        )
        if not name:
            continue
        entries.append(
            GaiaCodeHistoryEntry(
                code=code_match.group(1),
                name=name,
                condition=condition,
                unit=clean_text(values[9]),
                source_file=path.name,
                source_sheet=sheet.title,
                source_row=index + 1,
            )
        )
    workbook.close()
    return entries


HISTORY_CODE_ELIGIBLE_STATUSES = {
    "PACKAGE_EXACT",
    "PACKAGE_DATE_REVIEW",
    "PACKAGE_YEAR_REVIEW",
    "PACKAGE_CONDITION_REVIEW",
    "TREE_BRANCH_REVIEW",
    "TREE_REVIEW",
    "QUANTITY_RULE_REVIEW",
}


def apply_gaia_code_history(
    items: Iterable[BoqItem], entries: Iterable[GaiaCodeHistoryEntry]
) -> None:
    """Apply only one-code, same-package, unit-compatible history matches."""
    entries_by_name: dict[str, list[GaiaCodeHistoryEntry]] = {}
    for entry in entries:
        entries_by_name.setdefault(normalize_match_text(entry.name), []).append(entry)

    for item in items:
        if item.gaia_code or item.match_status not in HISTORY_CODE_ELIGIBLE_STATUSES:
            continue
        if item.package_match_type != "EXACT_NAME" or not item.package_name:
            continue

        candidates = entries_by_name.get(normalize_match_text(item.package_name), [])
        target_unit = item.package_unit or item.unit
        compatible = [
            entry
            for entry in candidates
            if entry.unit
            and compare_units(target_unit, entry.unit) in {"MATCH", "EQUIVALENT"}
        ]
        codes = {entry.code for entry in compatible}
        if len(codes) != 1:
            continue

        evidence = min(compatible, key=lambda entry: entry.source_row)
        item.gaia_code = next(iter(codes))
        item.gaia_item_name = item.package_name
        item.matched_rule = (
            f"GAIA_HISTORY:{evidence.source_file}:"
            f"{evidence.source_sheet}!{evidence.source_row}"
        )
        append_warning(
            item,
            f"既存GAIA設計書の同一施工パッケージ名・単位から"
            f"コード候補 {item.gaia_code} を補完しました",
        )


def load_reference_index(path: Path | None) -> dict:
    if path is None:
        return {}
    if not path.exists():
        raise FileNotFoundError(f"参照索引がありません: {path}")
    index = json.loads(path.read_text(encoding="utf-8"))
    if index.get("schema_version") != 1:
        raise ValueError("未対応の参照索引バージョンです")
    return index


def package_applies_to_district(package: dict, district: str) -> bool:
    if package.get("source_kind") != "ISHIKAWA_DISASTER_PACKAGE":
        return True
    return "珠洲" in clean_text(district)


def package_sort_key(package: dict, price_date: date | None) -> tuple[int, int, int]:
    package_year = int(package.get("fiscal_year") or 0)
    target_year = fiscal_year_for_date(price_date) if price_date else package_year
    year_match = int(package_year == target_year)
    return year_match, int(package.get("priority") or 0), package_year


def page_matches(
    pages: list[dict], item_keys: list[str], context_keys: list[str]
) -> tuple[str, str]:
    found_pages: list[int] = []
    context_pages: list[int] = []
    usable_item_keys = [key for key in item_keys if len(key) >= 2]
    usable_context_keys = [key for key in context_keys if len(key) >= 2]
    for page in pages:
        text = page.get("search_text", "")
        if not any(key in text for key in usable_item_keys):
            continue
        found_pages.append(int(page["page"]))
        if any(key in text for key in usable_context_keys):
            context_pages.append(int(page["page"]))
    if not found_pages:
        return "NOT_FOUND", ""
    if context_pages:
        return "EXACT_CONTEXT", ",".join(str(page) for page in context_pages[:10])
    return "ITEM_FOUND", ",".join(str(page) for page in found_pages[:10])


def _tree_similarity(left: object, right: object) -> float:
    left_key = normalize_match_text(left)
    right_key = normalize_match_text(right)
    if not left_key or not right_key:
        return 0.0
    if left_key == right_key:
        return 1.0
    if (
        min(len(left_key), len(right_key)) >= 3
        and (left_key in right_key or right_key in left_key)
    ):
        return 0.9
    return SequenceMatcher(None, left_key, right_key).ratio()


def _tree_item_names(item: BoqItem) -> list[str]:
    return list(
        dict.fromkeys(
            value
            for value in (
                item.item_name,
                item.gaia_item_name,
                item.package_name,
            )
            if clean_text(value)
        )
    )


def _tree_context_names(item: BoqItem) -> list[str]:
    return list(
        dict.fromkeys(
            value
            for value in (item.category, item.work_type, item.section)
            if clean_text(value)
        )
    )


def infer_tree_category(items: Iterable[BoqItem], paths: list[dict]) -> str:
    scores: dict[str, float] = defaultdict(float)
    for item in items:
        item_names = _tree_item_names(item)
        context_names = _tree_context_names(item)
        per_category: dict[str, float] = defaultdict(float)
        for path in paths:
            category = clean_text(path.get("business_category"))
            if not category:
                continue
            level4_score = max(
                (_tree_similarity(name, path.get("level4")) for name in item_names),
                default=0.0,
            )
            level3_score = max(
                (
                    _tree_similarity(context, path.get("level3"))
                    for context in context_names
                ),
                default=0.0,
            )
            score = 0.0
            if level4_score >= 0.9:
                score += 2.0
            if level3_score >= 0.9:
                score += 3.0
            if level4_score >= 0.9 and level3_score >= 0.75:
                score += 2.0
            per_category[category] = max(per_category[category], score)
        for category, score in per_category.items():
            scores[category] += score
    if not scores:
        return ""
    ranked = sorted(scores.items(), key=lambda pair: pair[1], reverse=True)
    if ranked[0][1] <= 0:
        return ""
    return ranked[0][0]


def infer_preferred_tree_level1(
    items: Iterable[BoqItem],
    paths: list[dict],
    tree_category: str,
    project_name: str,
) -> str:
    category_key = normalize_match_text(tree_category)
    level1_names = {
        clean_text(path.get("level1"))
        for path in paths
        if normalize_match_text(path.get("business_category")) == category_key
        and clean_text(path.get("level1"))
    }
    bridge_level1_names = {
        name for name in level1_names if "橋梁" in normalize_match_text(name)
    }
    if len(bridge_level1_names) != 1:
        return ""

    project_text = clean_text(project_name)
    if "橋" in project_text or "橋梁" in project_text:
        return next(iter(bridge_level1_names))

    bridge_evidence = 0
    for item in items:
        context = combine_unique(
            item.section,
            item.work_type,
            item.category,
            item.item_name,
        )
        if any(term in context for term in BRIDGE_CONTEXT_TERMS):
            bridge_evidence += 1
    return next(iter(bridge_level1_names)) if bridge_evidence >= 3 else ""


def resolve_level_tree_paths(
    items: list[BoqItem],
    paths: list[dict],
    tree_category: str,
    project_name: str = "",
) -> str:
    selected_category = clean_text(tree_category) or infer_tree_category(items, paths)
    branch_paths = [
        path
        for path in paths
        if normalize_match_text(path.get("business_category"))
        == normalize_match_text(selected_category)
    ]
    if not branch_paths:
        for item in items:
            item.tree_category = selected_category
            item.tree_status = "NOT_FOUND"
        return selected_category
    preferred_level1 = infer_preferred_tree_level1(
        items, paths, selected_category, project_name
    )
    resolution_paths = (
        [
            path
            for path in branch_paths
            if clean_text(path.get("level1")) == preferred_level1
        ]
        if preferred_level1
        else branch_paths
    )

    for item in items:
        item.tree_category = selected_category
        item.tree_level1 = preferred_level1
        item.tree_level4 = clean_text(item.gaia_item_name or item.item_name)
        item_names = _tree_item_names(item)
        context_names = _tree_context_names(item)

        scored: list[tuple[float, float, float, dict]] = []
        for path in resolution_paths:
            level4_score = max(
                (_tree_similarity(name, path.get("level4")) for name in item_names),
                default=0.0,
            )
            if level4_score < 0.9:
                continue
            level3_score = max(
                (
                    _tree_similarity(context, path.get("level3"))
                    for context in context_names
                ),
                default=0.0,
            )
            level2_score = max(
                (
                    _tree_similarity(context, path.get("level2"))
                    for context in context_names
                ),
                default=0.0,
            )
            scored.append(
                (
                    level4_score * 2
                    + level3_score * 3
                    + level2_score,
                    level3_score,
                    level2_score,
                    path,
                )
            )

        if scored:
            scored.sort(key=lambda pair: pair[0], reverse=True)
            all_hierarchies = {
                (
                    clean_text(path.get("level1")),
                    clean_text(path.get("level2")),
                    clean_text(path.get("level3")),
                    clean_text(path.get("level4")),
                )
                for _, _, _, path in scored
            }
            if len(all_hierarchies) == 1:
                eligible = [
                    candidate
                    for candidate in scored
                    if not context_names
                    or candidate[1] >= 0.75
                    or candidate[2] >= 0.9
                ]
            else:
                eligible = [
                    candidate
                    for candidate in scored
                    if candidate[1] >= 0.75
                    and (candidate[2] >= 0.9 or preferred_level1)
                ]
            best_score, _, _, best_path = eligible[0] if eligible else scored[0]
            tied_paths = [
                path
                for score, _, _, path in eligible
                if abs(score - best_score) < 0.05
            ]
            tied_hierarchies = {
                (
                    clean_text(path.get("level1")),
                    clean_text(path.get("level2")),
                    clean_text(path.get("level3")),
                    clean_text(path.get("level4")),
                )
                for path in tied_paths
            }
            if eligible and len(tied_hierarchies) == 1:
                item.tree_level1 = clean_text(best_path.get("level1"))
                item.tree_level2 = clean_text(best_path.get("level2"))
                item.tree_level3 = clean_text(best_path.get("level3"))
                item.tree_level4 = clean_text(best_path.get("level4"))
                item.tree_pages = ",".join(
                    str(page) for page in best_path.get("pages", [])[:10]
                )
                item.tree_match_score = round(best_score, 3)
                item.tree_status = "LEVEL4_VERIFIED"
                continue

        context_paths: list[tuple[float, dict]] = []
        for path in resolution_paths:
            level3_score = max(
                (
                    _tree_similarity(context, path.get("level3"))
                    for context in context_names
                ),
                default=0.0,
            )
            if level3_score < 0.9:
                continue
            context_paths.append((level3_score, path))
        if context_paths:
            context_paths.sort(key=lambda pair: pair[0], reverse=True)
            best_score, best_path = context_paths[0]
            best_hierarchy = (
                clean_text(best_path.get("level1")),
                clean_text(best_path.get("level2")),
                clean_text(best_path.get("level3")),
            )
            tied_hierarchies = {
                (
                    clean_text(path.get("level1")),
                    clean_text(path.get("level2")),
                    clean_text(path.get("level3")),
                )
                for score, path in context_paths
                if abs(score - best_score) < 0.05
            }
            if len(tied_hierarchies) == 1:
                item.tree_level1, item.tree_level2, item.tree_level3 = best_hierarchy
                item.tree_pages = ",".join(
                    str(page) for page in best_path.get("pages", [])[:10]
                )
                item.tree_match_score = round(best_score, 3)
                item.tree_status = "LEVEL3_VERIFIED_LEVEL4_UNVERIFIED"
                append_warning(
                    item,
                    "積算体系ツリーでレベル3までは確認できましたが、"
                    "レベル4細別は原文のまま要確認です",
                )
                continue

        if scored:
            item.tree_status = "LEVEL4_AMBIGUOUS_PATH"
            append_warning(
                item,
                "レベル4名称は積算体系ツリーにありますが、"
                "上位区分が複数あるため自動確定していません",
            )
        elif context_paths:
            item.tree_status = "LEVEL3_AMBIGUOUS_PATH"
            append_warning(
                item,
                "レベル3名称は積算体系ツリーにありますが、"
                "上位区分が複数あるため自動確定していません",
            )
        else:
            item.tree_status = "NOT_FOUND"
    return selected_category


def guideline_matches(
    pages: list[dict], item_keys: list[str], context_keys: list[str]
) -> tuple[str, str]:
    status, page_numbers = page_matches(pages, item_keys, context_keys)
    if status != "NOT_FOUND":
        return "ITEM_FOUND", page_numbers
    context_status, context_pages = page_matches(pages, context_keys, [])
    if context_status != "NOT_FOUND":
        return "CONTEXT_FOUND", context_pages
    return "NOT_FOUND", ""


def package_reference_label(item: BoqItem) -> str:
    source_labels = {
        "ISHIKAWA_DISASTER_PACKAGE": "石川県被災地SP",
        "NATIONAL_PACKAGE_PDF": "国SP",
        "NATIONAL_REFERENCE_WORKBOOK": "国SP参考",
    }
    if not item.package_no:
        return ""
    label = source_labels.get(item.package_source, item.package_source)
    page = f" p.{item.package_source_page}" if item.package_source_page else ""
    return f"{label} R{(item.package_fiscal_year or 2018) - 2018} No.{item.package_no}{page}"


def apply_reference_index(
    items: Iterable[BoqItem],
    index: dict,
    *,
    price_date: date | None,
    district: str,
    tree_category: str,
    project_name: str = "",
) -> None:
    item_list = list(items)
    packages = [
        package
        for package in index.get("package_catalog", [])
        if package_applies_to_district(package, district)
    ]
    packages_by_name: dict[str, list[dict]] = {}
    for package in packages:
        packages_by_name.setdefault(package.get("search_name", ""), []).append(package)

    all_tree_pages = index.get("level_tree", {}).get("pages", [])
    tree_paths = index.get("level_tree", {}).get("paths", [])
    selected_tree_category = (
        resolve_level_tree_paths(
            item_list, tree_paths, tree_category, project_name=project_name
        )
        if tree_paths
        else clean_text(tree_category)
    )
    normalized_tree_category = normalize_match_text(selected_tree_category)
    tree_pages = (
        [
            page
            for page in all_tree_pages
            if normalized_tree_category
            in normalize_match_text(page.get("heading", ""))
        ]
        if normalized_tree_category
        else all_tree_pages
    )
    guideline_pages = index.get("quantity_guideline", {}).get("pages", [])

    for item in item_list:
        item.tree_category = selected_tree_category
        item.source_standard_year = source_fiscal_year(item.standard)
        item_keys = list(
            dict.fromkeys(
                key
                for key in [
                    normalize_match_text(item.item_name),
                    normalize_match_text(item.gaia_item_name),
                    normalize_match_text(item.setting),
                ]
                if key
            )
        )
        context_keys = list(
            dict.fromkeys(
                key
                for key in [
                    normalize_match_text(item.category),
                    normalize_match_text(item.work_type),
                ]
                if key
            )
        )

        exact_candidates: list[dict] = []
        for key in item_keys:
            exact_candidates.extend(packages_by_name.get(key, []))
        unique_candidates = {
            (
                candidate.get("source_kind", ""),
                candidate.get("package_no", ""),
                candidate.get("fiscal_year"),
            ): candidate
            for candidate in exact_candidates
        }
        exact_candidates = list(unique_candidates.values())

        if exact_candidates:
            best_sort_key = max(
                package_sort_key(candidate, price_date)
                for candidate in exact_candidates
            )
            best_candidates = [
                candidate
                for candidate in exact_candidates
                if package_sort_key(candidate, price_date) == best_sort_key
            ]
            distinct_package_numbers = {
                candidate.get("package_no", "") for candidate in best_candidates
            }
            if len(distinct_package_numbers) > 1:
                item.package_match_type = "AMBIGUOUS_EXACT_NAME"
                item.package_suggestions = " | ".join(
                    f"{candidate.get('source_kind')} No.{candidate.get('package_no')} "
                    f"{candidate.get('name')}"
                    for candidate in best_candidates[:5]
                )
                if not item.tree_status:
                    item.tree_status, item.tree_pages = page_matches(
                        tree_pages, item_keys, context_keys
                    )
                item.guideline_status, item.guideline_pages = guideline_matches(
                    guideline_pages, item_keys, context_keys
                )
                append_warning(
                    item, "同名の施工パッケージが複数あるため自動確定しません"
                )
                continue
            package = best_candidates[0]
            item.package_match_type = "EXACT_NAME"
            item.package_name = clean_text(package.get("name"))
            item.package_no = clean_text(package.get("package_no"))
            item.package_unit = normalize_unit(package.get("unit"))
            item.package_source = clean_text(package.get("source_kind"))
            item.package_source_page = clean_text(package.get("source_page"))
            item.package_fiscal_year = package.get("fiscal_year")
            item.package_effective_from = clean_text(package.get("effective_from"))
            item.package_scope = clean_text(package.get("scope"))
            condition_fields = [
                clean_text(value) for value in package.get("condition_fields", []) if value
            ]
            item.package_condition_fields = " / ".join(condition_fields)
            item.package_condition_status = (
                "PROVIDED"
                if clean_text(item.setting) not in BLANK_VALUES
                else "MISSING_EXPLICIT_SETTINGS"
            )
            if not condition_fields:
                item.package_condition_status = "SCHEMA_UNAVAILABLE"

            if item.package_unit:
                item.package_unit_status = compare_units(item.unit, item.package_unit)
            else:
                item.package_unit_status = "REFERENCE_UNIT_MISSING"

            if price_date is None:
                item.reference_date_status = "MISSING_PRICE_DATE"
            elif (
                item.package_fiscal_year
                and fiscal_year_for_date(price_date) != item.package_fiscal_year
            ):
                item.reference_date_status = "FISCAL_YEAR_MISMATCH"
            elif (
                item.package_effective_from
                and price_date < date.fromisoformat(item.package_effective_from)
            ):
                item.reference_date_status = "BEFORE_EFFECTIVE_DATE"
            else:
                item.reference_date_status = "VALID"

            item.source_standard_status = (
                "YEAR_MISMATCH"
                if item.source_standard_year
                and item.package_fiscal_year
                and item.source_standard_year != item.package_fiscal_year
                else "CONSISTENT"
                if item.source_standard_year
                else "SOURCE_YEAR_UNKNOWN"
            )
            item_keys.append(normalize_match_text(item.package_name))
        else:
            scored: list[tuple[float, dict]] = []
            primary_key = normalize_match_text(item.item_name)
            for package in packages:
                package_key = package.get("search_name", "")
                if not primary_key or not package_key:
                    continue
                score = SequenceMatcher(None, primary_key, package_key).ratio()
                if score >= 0.78:
                    scored.append((score, package))
            scored.sort(
                key=lambda pair: (
                    pair[0],
                    package_sort_key(pair[1], price_date),
                ),
                reverse=True,
            )
            if scored:
                item.package_match_type = "FUZZY_SUGGESTION"
                suggestions = []
                for score, package in scored[:3]:
                    suggestions.append(
                        f"{package.get('source_kind')} No.{package.get('package_no')} "
                        f"{package.get('name')} ({score:.2f})"
                    )
                item.package_suggestions = " | ".join(suggestions)
            else:
                item.package_match_type = "NOT_FOUND"

        if not item.tree_status:
            item.tree_status, item.tree_pages = page_matches(
                tree_pages, item_keys, context_keys
            )
        item.guideline_status, item.guideline_pages = guideline_matches(
            guideline_pages, item_keys, context_keys
        )

        if item.tree_status == "NOT_FOUND":
            append_warning(item, "積算体系ツリーで細別・施工パッケージ名称を確認できません")
        if not item.tree_category:
            append_warning(item, "積算体系ツリーの事業区分が指定されていません")
        if item.guideline_status == "NOT_FOUND":
            append_warning(item, "数量算出要領で細別または上位区分を確認できません")
        if item.package_match_type == "FUZZY_SUGGESTION":
            append_warning(item, "施工パッケージは類似候補のみです。自動確定しません")
        elif item.package_match_type == "NOT_FOUND":
            append_warning(item, "施工パッケージ名称の完全一致がありません")
        if item.package_unit_status == "MISMATCH":
            append_warning(
                item,
                f"単位不一致: 数量計算書={item.unit}, 施工パッケージ={item.package_unit}",
            )
        if item.package_unit_status == "EQUIVALENT":
            append_warning(
                item,
                f"等価単位へ表記変換: {item.unit} -> {item.package_unit} (数量は変更しません)",
            )
        if item.package_condition_status == "MISSING_EXPLICIT_SETTINGS":
            append_warning(item, "施工パッケージ条件区分の設定内容が不足しています")
        if item.package_condition_status == "SCHEMA_UNAVAILABLE":
            append_warning(item, "現年度の条件区分スキーマを自動確認できません")
        if item.reference_date_status != "VALID" and item.package_match_type == "EXACT_NAME":
            append_warning(item, "積算年月と施工パッケージ適用年度の確認が必要です")
        if item.source_standard_status == "YEAR_MISMATCH":
            append_warning(
                item,
                f"数量計算書の基準年度({item.source_standard_year})と施工パッケージ年度"
                f"({item.package_fiscal_year})が一致しません",
            )


def classify_item(item: BoqItem, *, reference_index_active: bool = False) -> None:
    searchable = " ".join(
        [item.standard, item.remarks, item.setting, item.notes, item.specification]
    )

    if item.quantity is None or item.quantity <= 0:
        item.match_status = "INVALID_QUANTITY"
        item.confidence = 0.0
        append_warning(item, "数量を取得できない、または数量が0以下です")
    elif item.extraction_status != "READY":
        item.match_status = "SOURCE_REVIEW_REQUIRED"
        item.confidence = min(0.5, item.extraction_confidence)
        append_warning(item, "原本照合が完了するまで自動確定しません")
    elif "適用不可" in searchable:
        item.match_status = "BLOCKED_REVIEW"
        item.confidence = 0.0
        append_warning(item, "積算基準の適用不可と明記されています")
    elif "見積" in item.standard:
        item.match_status = "QUOTATION_REQUIRED"
        item.confidence = 0.2
        append_warning(item, "見積価格の入力が必要です")
    elif (
        "材料費" in item.remarks
        or "市場単価" in item.notes
        or "建設物価" in item.standard
        or "土木コスト情報" in item.standard
    ):
        item.match_status = "MARKET_PRICE_REVIEW"
        item.confidence = 0.5
        append_warning(item, "契約している市販単価データと適用月の確認が必要です")
    elif reference_index_active and item.package_match_type == "EXACT_NAME":
        if item.package_unit_status not in {"MATCH", "EQUIVALENT"}:
            item.match_status = "PACKAGE_UNIT_REVIEW"
            item.confidence = 0.2
        elif item.reference_date_status != "VALID":
            item.match_status = "PACKAGE_DATE_REVIEW"
            item.confidence = 0.5
        elif item.source_standard_status == "YEAR_MISMATCH":
            item.match_status = "PACKAGE_YEAR_REVIEW"
            item.confidence = 0.5
        elif item.package_condition_status != "PROVIDED":
            item.match_status = "PACKAGE_CONDITION_REVIEW"
            item.confidence = 0.6
        elif not item.tree_category:
            item.match_status = "TREE_BRANCH_REVIEW"
            item.confidence = 0.5
        elif item.tree_status in {
            "NOT_FOUND",
            "LEVEL2_PROJECT_CONTEXT",
            "LEVEL3_VERIFIED_LEVEL4_UNVERIFIED",
            "LEVEL3_AMBIGUOUS_PATH",
            "LEVEL4_AMBIGUOUS_PATH",
        }:
            item.match_status = "TREE_REVIEW"
            item.confidence = 0.55
        elif item.guideline_status == "NOT_FOUND":
            item.match_status = "QUANTITY_RULE_REVIEW"
            item.confidence = 0.55
        elif item.gaia_code:
            item.match_status = "EXACT_CODE"
            item.confidence = 0.98
        else:
            item.match_status = "PACKAGE_EXACT"
            item.confidence = (
                0.93
                if item.package_source == "ISHIKAWA_DISASTER_PACKAGE"
                else 0.9
            )
    elif reference_index_active and item.gaia_code:
        item.match_status = "CODE_REFERENCE_REVIEW"
        item.confidence = 0.6
    elif reference_index_active:
        item.match_status = "MANUAL_REVIEW"
        item.confidence = 0.35 if item.package_suggestions else 0.25
    elif item.gaia_code:
        item.match_status = "EXACT_CODE"
        item.confidence = 0.98
    elif "国土交通省" in item.standard and item.setting not in BLANK_VALUES:
        item.match_status = "AUTO_CANDIDATE"
        item.confidence = 0.85
    else:
        item.match_status = "MANUAL_REVIEW"
        item.confidence = 0.4


def clear_values(sheet: Worksheet) -> None:
    for row in sheet.iter_rows():
        for cell in row:
            if not isinstance(cell, MergedCell):
                cell.value = None


def set_cover_fields(
    workbook,
    project_name: str,
    location: str,
    work_category: str,
    district: str,
    price_date: date | None,
) -> None:
    sheet = workbook["鏡"]
    sheet["D6"] = project_name
    sheet["D7"] = location
    sheet["E8"] = None
    sheet["J20"] = None
    sheet["D21"] = work_category
    sheet["J21"] = price_date
    sheet["O20"] = district


@dataclass
class OutputBlock:
    level: str
    value: str
    marker: str = ""
    item: BoqItem | None = None


def build_output_blocks(items: Iterable[BoqItem]) -> list[OutputBlock]:
    blocks: list[OutputBlock] = []
    previous_section = ""
    previous_work_type = ""
    previous_category = ""

    for item in items:
        if item.match_status == "INVALID_QUANTITY":
            continue
        section = clean_text(item.output_level1 or item.section) or "本工事費"
        work_type = clean_text(item.output_level2 or item.work_type) or "工種未確認"
        category = clean_text(item.output_level3 or item.category) or "種別未確認"
        item_name = clean_text(
            item.output_level4 or item.gaia_item_name or item.item_name
        )
        if section != previous_section:
            blocks.append(OutputBlock("fee", section, "費目行"))
            previous_section = section
            previous_work_type = ""
            previous_category = ""
        if work_type != previous_work_type:
            blocks.append(OutputBlock("work", work_type, "工種行"))
            previous_work_type = work_type
            previous_category = ""
        if category != previous_category:
            blocks.append(OutputBlock("category", category, "種別行"))
            previous_category = category
        blocks.append(OutputBlock("item", item_name, item=item))
    return blocks


def write_page_header(sheet: Worksheet, page_index: int) -> None:
    start = 1 + page_index * 40
    sheet.cell(start, 1).value = "本工事費内訳表"
    sheet.cell(start + 1, 13).value = page_index + 2
    sheet.cell(start + 2, 1).value = "費目　工種　種別　細別・規格"
    sheet.cell(start + 2, 8).value = "数　　量"
    sheet.cell(start + 2, 10).value = "単　位"
    sheet.cell(start + 2, 11).value = "単　　価"
    sheet.cell(start + 2, 12).value = "金　　額"
    sheet.cell(start + 2, 13).value = "摘　　　　要"


def write_block(sheet: Worksheet, block_index: int, block: OutputBlock) -> None:
    page_index, slot = divmod(block_index, 9)
    row = 1 + page_index * 40 + 3 + slot * 4
    if slot == 0:
        write_page_header(sheet, page_index)

    if block.level != "item":
        column = {"fee": 1, "work": 2, "category": 3}[block.level]
        sheet.cell(row, column).value = block.value
        sheet.cell(row, 8).value = 1
        sheet.cell(row, 10).value = "式"
        sheet.cell(row + 1, 13).value = block.marker
        return

    item = block.item
    assert item is not None
    sheet.cell(row, 4).value = block.value
    sheet.cell(row, 8).value = item.quantity
    sheet.cell(row, 10).value = item.unit
    sheet.cell(row, 13).value = f"[{item.gaia_code}]" if item.gaia_code else ""
    sheet.cell(row + 1, 4).value = item.gaia_condition


def build_gaia_candidate(
    template_path: Path,
    output_path: Path,
    project_name: str,
    items: list[BoqItem],
    location: str,
    work_category: str,
    district: str,
    price_date: date | None,
) -> int:
    workbook = load_workbook(template_path, data_only=False, read_only=False)
    required = {
        "鏡",
        "本工事費内訳表",
        "内訳書",
        "明細書",
        "代価表",
        "単価表",
        "施工パッケージ",
    }
    missing = required.difference(workbook.sheetnames)
    if missing:
        raise ValueError(f"テンプレートに必要なシートがありません: {sorted(missing)}")

    set_cover_fields(
        workbook, project_name, location, work_category, district, price_date
    )
    main_sheet = workbook["本工事費内訳表"]
    clear_values(main_sheet)
    for sheet_name in required - {"鏡", "本工事費内訳表"}:
        clear_values(workbook[sheet_name])

    blocks = build_output_blocks(items)
    capacity = ((main_sheet.max_row - 1) // 40 + 1) * 9
    if len(blocks) > capacity:
        raise ValueError(
            f"テンプレート容量不足: {len(blocks)} blocks > {capacity} blocks"
        )
    for index, block in enumerate(blocks):
        write_block(main_sheet, index, block)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)
    workbook.close()
    return len(blocks)


def write_review_csv(path: Path, items: list[BoqItem]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = [
        "source_format",
        "extractor",
        "source_sheet",
        "source_page",
        "source_row",
        "extraction_status",
        "extraction_confidence",
        "extraction_warnings",
        "section",
        "work_type",
        "category",
        "item_name",
        "specification",
        "unit",
        "quantity",
        "standard",
        "daily_output",
        "setting",
        "remarks",
        "notes",
        "source_reference",
        "gaia_item_name",
        "gaia_code",
        "gaia_condition",
        "match_status",
        "confidence",
        "matched_rule",
        "tree_category",
        "tree_status",
        "tree_pages",
        "tree_level1",
        "tree_level2",
        "tree_level3",
        "tree_level4",
        "tree_match_score",
        "guideline_status",
        "guideline_pages",
        "package_match_type",
        "package_name",
        "package_no",
        "package_unit",
        "package_source",
        "package_source_page",
        "package_fiscal_year",
        "package_effective_from",
        "package_scope",
        "package_condition_fields",
        "package_condition_status",
        "package_unit_status",
        "reference_date_status",
        "source_standard_year",
        "source_standard_status",
        "package_suggestions",
        "output_level1",
        "output_level2",
        "output_level3",
        "output_level4",
        "catalog_status",
        "catalog_score",
        "catalog_name_score",
        "catalog_specification_score",
        "catalog_candidate_count",
        "catalog_path_safe",
        "catalog_code_safe",
        "catalog_source_row",
        "warnings",
    ]
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for item in items:
            row = asdict(item)
            row["extraction_warnings"] = "; ".join(item.extraction_warnings)
            row["warnings"] = "; ".join(item.warnings)
            writer.writerow({name: row.get(name, "") for name in fieldnames})


def parse_iso_date(value: str) -> date | None:
    return date.fromisoformat(value) if value else None


def main() -> int:
    parser = argparse.ArgumentParser(
        description="数量計算書の設計数量総括表をGaia取込候補へ変換します。"
    )
    parser.add_argument("--source", required=True, type=Path)
    parser.add_argument(
        "--extraction-work-dir",
        type=Path,
        help="PDFページとOCRセルを診断用に保存するディレクトリ。",
    )
    parser.add_argument("--template", required=True, type=Path)
    parser.add_argument(
        "--catalog",
        type=Path,
        default=Path(__file__).resolve().parent
        / "assets"
        / "gaia_import_catalog.json",
        help="GAIA取込実績から作成した名称・コードカタログ。",
    )
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--review", type=Path)
    parser.add_argument(
        "--rules",
        type=Path,
        help=argparse.SUPPRESS,
    )
    parser.add_argument(
        "--code-history",
        action="append",
        default=[],
        type=Path,
        help=argparse.SUPPRESS,
    )
    parser.add_argument(
        "--reference-index",
        type=Path,
        help=argparse.SUPPRESS,
    )
    parser.add_argument(
        "--tree-category",
        default="",
        help=argparse.SUPPRESS,
    )
    parser.add_argument("--project-name", default="")
    parser.add_argument("--location", default="")
    parser.add_argument(
        "--work-category",
        default="",
        help="鏡シートの工種区分。省略時は工事名と明細から推定します。",
    )
    parser.add_argument("--district", default="珠洲")
    parser.add_argument(
        "--price-date",
        default="",
        help="YYYY-MM-DD。省略時は原本から検出し、無ければPC日付を使用します。",
    )
    args = parser.parse_args()

    project_name, items = extract_boq_items(
        args.source, work_dir=args.extraction_work_dir
    )
    if args.project_name:
        project_name = clean_text(args.project_name)
    if args.price_date:
        price_date = parse_iso_date(args.price_date)
        price_date_source = "USER"
    else:
        price_date, price_date_source = detect_estimate_date(args.source)
    catalog = load_gaia_catalog(args.catalog)
    apply_gaia_catalog(items, catalog)
    for item in items:
        classify_schema_item(item)
    work_category = clean_text(args.work_category) or infer_work_category(
        project_name, items
    )

    review_path = args.review or args.output.with_name(
        f"{args.output.stem}_review.csv"
    )
    write_review_csv(review_path, items)
    block_count = build_gaia_candidate(
        template_path=args.template,
        output_path=args.output,
        project_name=project_name,
        items=items,
        location=clean_text(args.location),
        work_category=work_category,
        district=clean_text(args.district),
        price_date=price_date,
    )

    counts: dict[str, int] = {}
    for item in items:
        counts[item.match_status] = counts.get(item.match_status, 0) + 1
    result = {
        "project_name": project_name,
        "source_items": len(items),
        "source_format": items[0].source_format if items else "",
        "extraction_review_required": sum(
            item.extraction_status != "READY" for item in items
        ),
        "output_blocks": block_count,
        "status_counts": counts,
        "catalog": str(args.catalog.resolve()),
        "catalog_entries": len(catalog),
        "catalog_name_matched": sum(
            item.catalog_status.startswith("EXACT")
            or item.catalog_status == "FUZZY_PATH"
            for item in items
        ),
        "catalog_codes_applied": sum(bool(item.gaia_code) for item in items),
        "price_date": price_date.isoformat(),
        "price_date_source": price_date_source,
        "work_category": work_category,
        "output": str(args.output.resolve()),
        "review": str(review_path.resolve()),
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
