import unittest
from datetime import date
from pathlib import Path

from openpyxl import Workbook

from gaia_converter import (
    BoqItem,
    GaiaCodeHistoryEntry,
    apply_gaia_code_history,
    apply_reference_index,
    classify_item,
    compare_units,
    extract_gaia_code_history,
    extract_boq_items,
    normalize_match_text,
)


class ConverterTests(unittest.TestCase):
    def make_package_item(self, standard="R8_国土交通省土木工事積算基準"):
        return BoqItem(
            source_sheet="設計数量総括表(下部工)",
            source_row=10,
            section="下部工",
            work_type="舗装工",
            category="アスファルト舗装工",
            item_name="表層(車道・路肩部)",
            specification="再生密粒度アスコン t=6cm",
            unit="m2",
            quantity=100,
            remarks="",
            standard=standard,
            daily_output="2300m2",
            setting="3.0m超,t=6cm,再生密粒度アスコン(20F),無し",
            notes="",
            source_reference="P.1442",
            gaia_item_name="表層(車道・路肩部)",
        )

    def make_reference_index(self):
        item_key = normalize_match_text("表層(車道・路肩部)")
        category_key = normalize_match_text("アスファルト舗装工")
        return {
            "schema_version": 1,
            "level_tree": {
                "pages": [
                    {
                        "page": 20,
                        "heading": "積算体系ツリー 河川改修",
                        "search_text": f"{category_key}{item_key}",
                    }
                ]
            },
            "quantity_guideline": {
                "pages": [{"page": 300, "search_text": item_key}]
            },
            "package_catalog": [
                {
                    "package_no": "018",
                    "name": "表層(車道・路肩部)",
                    "search_name": item_key,
                    "unit": "m2",
                    "condition_fields": ["施工幅員", "舗装厚", "材料規格"],
                    "source_kind": "ISHIKAWA_DISASTER_PACKAGE",
                    "source_page": "018-1～7",
                    "fiscal_year": 2026,
                    "effective_from": "2026-04-01",
                    "scope": "石川県中能登・奥能登地域(珠洲市を含む)",
                    "priority": 300,
                }
            ],
        }

    def test_extracts_two_row_summary_item(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "設計数量総括表(上部工)"
        sheet["A2"] = "設計書名：テスト橋"
        sheet.append(["工種", "種別", "細別", "規格", "単位", "数量"])
        sheet.append(["舗装工", None, None, None, None, None])
        sheet.append([None, "舗装工", None, None, None, None])
        sheet.append(
            [
                None,
                None,
                "表層(車道・路肩部)",
                "再生密粒度アスコン t=5cm",
                "m2",
                None,
                None,
                "R7_国土交通省土木工事積算基準",
                "2300m2",
                "3.0m超",
                None,
            ]
        )
        sheet.append([None, None, None, None, None, 12.5, None, "P.1442"])

        path = Path.cwd() / "test_source.xlsx"
        try:
            workbook.save(path)
            project_name, items = extract_boq_items(path)
        finally:
            path.unlink(missing_ok=True)

        self.assertEqual(project_name, "テスト橋")
        self.assertEqual(len(items), 1)
        self.assertEqual(items[0].quantity, 12.5)
        self.assertEqual(items[0].work_type, "舗装工")
        self.assertEqual(items[0].category, "舗装工")
        self.assertEqual(items[0].source_reference, "P.1442")

    def test_explicit_not_applicable_is_blocked(self):
        item = BoqItem(
            source_sheet="設計数量総括表(上部工)",
            source_row=10,
            section="上部工",
            work_type="RC橋工",
            category="支承工",
            item_name="ゴム支承設置工",
            specification="帯状ゴム支承",
            unit="式",
            quantity=1,
            remarks="",
            standard="見積り",
            daily_output="(該当なし)",
            setting="―",
            notes="積算基準に記載が無いため、適用不可",
            source_reference="",
        )
        classify_item(item)
        self.assertEqual(item.match_status, "BLOCKED_REVIEW")
        self.assertEqual(item.confidence, 0.0)

    def test_current_ishikawa_package_is_exact_but_not_a_gaia_code(self):
        item = self.make_package_item()
        apply_reference_index(
            [item],
            self.make_reference_index(),
            price_date=date(2026, 7, 1),
            district="珠洲",
            tree_category="河川改修",
        )
        classify_item(item, reference_index_active=True)

        self.assertEqual(item.match_status, "PACKAGE_EXACT")
        self.assertEqual(item.package_source, "ISHIKAWA_DISASTER_PACKAGE")
        self.assertEqual(item.package_no, "018")
        self.assertEqual(item.package_unit_status, "MATCH")
        self.assertEqual(item.gaia_code, "")

    def test_missing_price_date_blocks_package_confirmation(self):
        item = self.make_package_item()
        apply_reference_index(
            [item],
            self.make_reference_index(),
            price_date=None,
            district="珠洲",
            tree_category="河川改修",
        )
        classify_item(item, reference_index_active=True)

        self.assertEqual(item.reference_date_status, "MISSING_PRICE_DATE")
        self.assertEqual(item.match_status, "PACKAGE_DATE_REVIEW")

    def test_source_standard_year_mismatch_is_reviewed(self):
        item = self.make_package_item(standard="R7_国土交通省土木工事積算基準")
        apply_reference_index(
            [item],
            self.make_reference_index(),
            price_date=date(2026, 7, 1),
            district="珠洲",
            tree_category="河川改修",
        )
        classify_item(item, reference_index_active=True)

        self.assertEqual(item.source_standard_status, "YEAR_MISMATCH")
        self.assertEqual(item.match_status, "PACKAGE_YEAR_REVIEW")

    def test_trailing_work_suffix_and_tsumikomi_are_normalized(self):
        self.assertEqual(
            normalize_match_text("排水管設置工"),
            normalize_match_text("排水管設置"),
        )
        self.assertEqual(
            normalize_match_text("積込み(コンクリート殻)"),
            normalize_match_text("積込(コンクリート殻)"),
        )
        self.assertEqual(compare_units("孔", "箇所"), "EQUIVALENT")
        self.assertEqual(compare_units("掛m2", "m2"), "EQUIVALENT")

    def test_duplicate_current_package_names_are_not_auto_selected(self):
        item = self.make_package_item()
        item.item_name = "コンクリート工"
        item.gaia_item_name = "コンクリート工"
        item.setting = "無筋・鉄筋構造物、ポンプ車打設"
        index = self.make_reference_index()
        key = normalize_match_text("コンクリート")
        index["package_catalog"] = [
            {
                **index["package_catalog"][0],
                "package_no": package_no,
                "name": "コンクリート",
                "search_name": key,
                "source_kind": "NATIONAL_PACKAGE_PDF",
                "priority": 200,
            }
            for package_no in ("157", "333")
        ]
        apply_reference_index(
            [item],
            index,
            price_date=date(2026, 7, 1),
            district="珠洲",
            tree_category="河川改修",
        )
        classify_item(item, reference_index_active=True)

        self.assertEqual(item.package_match_type, "AMBIGUOUS_EXACT_NAME")
        self.assertEqual(item.match_status, "MANUAL_REVIEW")
        self.assertEqual(item.package_no, "")

    def test_extracts_and_applies_unambiguous_gaia_code_history(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "本工事費内訳表"
        sheet["E4"] = "表層(車道・路肩部)"
        sheet["J4"] = "m2"
        sheet["M4"] = "[CB410260]"
        sheet["D5"] = "3.0m超 60mm"
        path = Path.cwd() / "test_history.xlsx"
        try:
            workbook.save(path)
            entries = extract_gaia_code_history(path)
        finally:
            path.unlink(missing_ok=True)

        item = self.make_package_item(standard="R8_国土交通省土木工事積算基準")
        apply_reference_index(
            [item],
            self.make_reference_index(),
            price_date=date(2026, 7, 1),
            district="珠洲",
            tree_category="河川改修",
        )
        classify_item(item, reference_index_active=True)
        apply_gaia_code_history([item], entries)
        classify_item(item, reference_index_active=True)

        self.assertEqual(len(entries), 1)
        self.assertEqual(item.gaia_code, "CB410260")
        self.assertEqual(item.match_status, "EXACT_CODE")
        self.assertTrue(item.matched_rule.startswith("GAIA_HISTORY:"))

    def test_history_code_is_not_applied_when_codes_are_ambiguous(self):
        item = self.make_package_item(standard="R8_国土交通省土木工事積算基準")
        apply_reference_index(
            [item],
            self.make_reference_index(),
            price_date=date(2026, 7, 1),
            district="珠洲",
            tree_category="河川改修",
        )
        classify_item(item, reference_index_active=True)
        entries = [
            GaiaCodeHistoryEntry(
                code=code,
                name="表層(車道・路肩部)",
                condition="",
                unit="m2",
                source_file="history.xlsx",
                source_sheet="本工事費内訳表",
                source_row=row,
            )
            for code, row in (("CB410260", 4), ("CB410240", 8))
        ]

        apply_gaia_code_history([item], entries)

        self.assertEqual(item.gaia_code, "")
        self.assertEqual(item.match_status, "PACKAGE_EXACT")

    def test_history_code_does_not_override_quotation_review(self):
        item = self.make_package_item(standard="見積り")
        apply_reference_index(
            [item],
            self.make_reference_index(),
            price_date=date(2026, 7, 1),
            district="珠洲",
            tree_category="河川改修",
        )
        classify_item(item, reference_index_active=True)
        entry = GaiaCodeHistoryEntry(
            code="CB410260",
            name="表層(車道・路肩部)",
            condition="",
            unit="m2",
            source_file="history.xlsx",
            source_sheet="本工事費内訳表",
            source_row=4,
        )

        apply_gaia_code_history([item], [entry])

        self.assertEqual(item.gaia_code, "")
        self.assertEqual(item.match_status, "QUOTATION_REQUIRED")

    def test_history_code_requires_unit_evidence(self):
        item = self.make_package_item(standard="R8_国土交通省土木工事積算基準")
        apply_reference_index(
            [item],
            self.make_reference_index(),
            price_date=date(2026, 7, 1),
            district="珠洲",
            tree_category="河川改修",
        )
        classify_item(item, reference_index_active=True)
        entry = GaiaCodeHistoryEntry(
            code="CB410260",
            name="表層(車道・路肩部)",
            condition="",
            unit="",
            source_file="history.xlsx",
            source_sheet="本工事費内訳表",
            source_row=4,
        )

        apply_gaia_code_history([item], [entry])

        self.assertEqual(item.gaia_code, "")


if __name__ == "__main__":
    unittest.main()
