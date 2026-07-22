from __future__ import annotations

import argparse
import csv
import json
import re
import unicodedata
from dataclasses import asdict, dataclass, field
from datetime import date
from difflib import SequenceMatcher
from pathlib import Path
from typing import Iterable

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.worksheet.worksheet import Worksheet


SUMMARY_SHEET_RE = re.compile(r"^設計数量総括表[（(](.+?)[）)]$")
HEADER_NAMES = {"工種", "種別", "細別", "規格", "単位", "数量"}
BLANK_VALUES = {"", "-", "―"}
SOURCE_YEAR_RE = re.compile(r"(?:令和|R)\s*(\d+)", re.IGNORECASE)
GAIA_CODE_RE = re.compile(r"^\[([A-Z]{2}\d{6})\]$")


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
    gaia_item_name: str = ""
    gaia_code: str = ""
    gaia_condition: str = ""
    match_status: str = ""
    confidence: float = 0.0
    matched_rule: str = ""
    tree_category: str = ""
    tree_status: str = ""
    tree_pages: str = ""
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
    warnings: list[str] = field(default_factory=list)


def clean_text(value: object) -> str:
    if value is None:
        return ""
    text = unicodedata.normalize("NFKC", str(value))
    return re.sub(r"\s+", " ", text).strip()


def clean_label(value: object) -> str:
    text = clean_text(value)
    japanese = r"一-龯々ぁ-ゖァ-ヺー"
    return re.sub(rf"(?<=[{japanese}])\s+(?=[{japanese}])", "", text)


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


def append_warning(item: BoqItem, message: str) -> None:
    cleaned = clean_text(message)
    if cleaned and cleaned not in item.warnings:
        item.warnings.append(cleaned)


def parse_quantity(value: object) -> float | int | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        number = float(value)
    else:
        text = clean_text(value).replace(",", "")
        match = re.fullmatch(r"[-+]?\d+(?:\.\d+)?", text)
        if not match:
            return None
        number = float(text)
    if number.is_integer():
        return int(number)
    return number


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


def extract_project_name(workbook) -> str:
    for sheet in workbook.worksheets:
        if not SUMMARY_SHEET_RE.match(sheet.title):
            continue
        for row in range(1, min(sheet.max_row, 10) + 1):
            value = clean_text(sheet.cell(row, 1).value)
            if value.startswith("設計書名:"):
                return value.split(":", 1)[1].strip()
    return "数量計算書変換工事"


def is_header_or_title(value: str) -> bool:
    normalized = clean_text(value).replace(" ", "")
    return (
        not normalized
        or normalized in HEADER_NAMES
        or normalized.startswith("設計数量総括表")
        or normalized.startswith("設計書名:")
    )


def find_quantity_row(sheet: Worksheet, item_row: int) -> tuple[float | int | None, int]:
    direct = parse_quantity(sheet.cell(item_row, 6).value)
    if direct is not None:
        return direct, item_row

    for row in range(item_row + 1, min(item_row + 4, sheet.max_row) + 1):
        if clean_text(sheet.cell(row, 3).value) or clean_text(sheet.cell(row, 5).value):
            break
        quantity = parse_quantity(sheet.cell(row, 6).value)
        if quantity is not None:
            return quantity, row
    return None, item_row


def extract_boq_items(source_path: Path) -> tuple[str, list[BoqItem]]:
    workbook = load_workbook(source_path, data_only=True, read_only=False)
    project_name = extract_project_name(workbook)
    items: list[BoqItem] = []

    for sheet in workbook.worksheets:
        match = SUMMARY_SHEET_RE.match(sheet.title)
        if not match:
            continue

        section = clean_text(match.group(1))
        current_work_type = ""
        current_category = ""

        for row in range(1, sheet.max_row + 1):
            col_a = clean_label(sheet.cell(row, 1).value)
            col_b = clean_label(sheet.cell(row, 2).value)
            col_c = clean_label(sheet.cell(row, 3).value)
            unit = clean_text(sheet.cell(row, 5).value)

            if col_a and not is_header_or_title(col_a):
                current_work_type = col_a
                current_category = ""
            if col_b and not is_header_or_title(col_b):
                current_category = col_b

            if not col_c or is_header_or_title(col_c) or not unit:
                continue

            quantity, quantity_row = find_quantity_row(sheet, row)
            item = BoqItem(
                source_sheet=sheet.title,
                source_row=row,
                section=section,
                work_type=current_work_type,
                category=current_category,
                item_name=col_c,
                specification=clean_text(sheet.cell(row, 4).value),
                unit=unit,
                quantity=quantity,
                remarks=clean_text(sheet.cell(row, 7).value),
                standard=clean_text(sheet.cell(row, 8).value),
                daily_output=clean_text(sheet.cell(row, 9).value),
                setting=clean_text(sheet.cell(row, 10).value),
                notes=clean_text(sheet.cell(row, 11).value),
                source_reference=clean_text(sheet.cell(quantity_row, 8).value),
            )
            item.gaia_item_name = item.item_name
            item.gaia_condition = combine_unique(item.specification, item.setting, item.remarks)
            items.append(item)

    workbook.close()
    return project_name, items


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
) -> None:
    packages = [
        package
        for package in index.get("package_catalog", [])
        if package_applies_to_district(package, district)
    ]
    packages_by_name: dict[str, list[dict]] = {}
    for package in packages:
        packages_by_name.setdefault(package.get("search_name", ""), []).append(package)

    all_tree_pages = index.get("level_tree", {}).get("pages", [])
    normalized_tree_category = normalize_match_text(tree_category)
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

    for item in items:
        item.tree_category = clean_text(tree_category)
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
        elif item.tree_status == "NOT_FOUND":
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
        if item.section != previous_section:
            blocks.append(OutputBlock("fee", item.section, "費目行"))
            previous_section = item.section
            previous_work_type = ""
            previous_category = ""
        if item.work_type and item.work_type != previous_work_type:
            blocks.append(OutputBlock("work", item.work_type, "工種行"))
            previous_work_type = item.work_type
            previous_category = ""
        if item.category and item.category != previous_category:
            blocks.append(OutputBlock("category", item.category, "種別行"))
            previous_category = item.category
        blocks.append(OutputBlock("item", item.gaia_item_name, item=item))
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
    sheet.cell(row, 10).value = (
        item.package_unit
        if item.package_unit_status == "EQUIVALENT"
        else item.unit
    )
    sheet.cell(row, 13).value = f"[{item.gaia_code}]" if item.gaia_code else ""
    sheet.cell(row + 1, 4).value = item.gaia_condition
    sheet.cell(row + 1, 13).value = item.standard
    sheet.cell(row + 2, 13).value = combine_unique(
        item.source_reference,
        package_reference_label(item),
        f"TREE[{item.tree_category}] p.{item.tree_pages}"
        if item.tree_pages
        else "",
        f"数量要領 p.{item.guideline_pages}" if item.guideline_pages else "",
    )
    sheet.cell(row + 3, 13).value = combine_unique(
        item.match_status,
        item.package_condition_status,
        item.notes,
        "; ".join(item.warnings),
    )


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
        "source_sheet",
        "source_row",
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
        "warnings",
    ]
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for item in items:
            row = asdict(item)
            row["warnings"] = "; ".join(item.warnings)
            writer.writerow({name: row.get(name, "") for name in fieldnames})


def parse_iso_date(value: str) -> date | None:
    return date.fromisoformat(value) if value else None


def main() -> int:
    parser = argparse.ArgumentParser(
        description="数量計算書の設計数量総括表をGaia取込候補へ変換します。"
    )
    parser.add_argument("--source", required=True, type=Path)
    parser.add_argument("--template", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--review", type=Path)
    parser.add_argument("--rules", type=Path)
    parser.add_argument(
        "--code-history",
        action="append",
        default=[],
        type=Path,
        help=(
            "GAIAコードを抽出する既存の取込実績Excel。複数回指定できます。"
            "同一施工パッケージ名・単位が1コードに確定する場合だけ補完します。"
        ),
    )
    parser.add_argument("--reference-index", type=Path)
    parser.add_argument(
        "--tree-category",
        default="",
        help="積算体系ツリーの事業区分。例: 河川改修、道路新設・改築、道路維持・修繕",
    )
    parser.add_argument("--project-name", default="")
    parser.add_argument("--location", default="")
    parser.add_argument("--work-category", default="橋梁工事")
    parser.add_argument("--district", default="珠洲")
    parser.add_argument("--price-date", default="", help="YYYY-MM-DD")
    args = parser.parse_args()

    project_name, items = extract_boq_items(args.source)
    if args.project_name:
        project_name = clean_text(args.project_name)
    price_date = parse_iso_date(args.price_date)
    rules = load_mapping_rules(args.rules)
    apply_mapping_rules(items, rules)
    reference_index = load_reference_index(args.reference_index)
    if reference_index:
        apply_reference_index(
            items,
            reference_index,
            price_date=price_date,
            district=clean_text(args.district),
            tree_category=clean_text(args.tree_category),
        )
    for item in items:
        classify_item(item, reference_index_active=bool(reference_index))

    code_history: list[GaiaCodeHistoryEntry] = []
    for history_path in args.code_history:
        if not history_path.exists():
            raise FileNotFoundError(f"GAIAコード履歴がありません: {history_path}")
        code_history.extend(extract_gaia_code_history(history_path))
    if code_history:
        apply_gaia_code_history(items, code_history)
        for item in items:
            classify_item(item, reference_index_active=bool(reference_index))

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
        work_category=clean_text(args.work_category),
        district=clean_text(args.district),
        price_date=price_date,
    )

    counts: dict[str, int] = {}
    for item in items:
        counts[item.match_status] = counts.get(item.match_status, 0) + 1
    result = {
        "project_name": project_name,
        "source_items": len(items),
        "output_blocks": block_count,
        "status_counts": counts,
        "reference_index": str(args.reference_index.resolve())
        if args.reference_index
        else "",
        "code_history_entries": len(code_history),
        "history_codes_applied": sum(
            item.matched_rule.startswith("GAIA_HISTORY:") for item in items
        ),
        "price_date": price_date.isoformat() if price_date else "",
        "tree_category": clean_text(args.tree_category),
        "output": str(args.output.resolve()),
        "review": str(review_path.resolve()),
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
