# GAIA quantity workbook converter for 珠洲市

This prototype converts `設計数量総括表(...)` sheets in a quantity-calculation workbook into a GAIA-style `金抜設計書` candidate. For 珠洲市 work, it can validate the transfer against the supplied 積算体系ツリー, 土木工事数量算出要領, and dated 施工パッケージ references.

## Quick start on another Windows PC

Requirements: Python 3.11 or newer, Microsoft Excel for reviewing the output, and access to the six governing reference files. GAIA does not need to be installed on the machine that performs the conversion.

```powershell
git clone YOUR_GITHUB_REPOSITORY_URL
cd gaia-automation
Set-ExecutionPolicy -Scope Process Bypass
.\setup.ps1
```

Build the local reference index once. The index is intentionally excluded from Git because it contains extracted reference text and local file metadata.

```powershell
.\build_references.ps1 `
  -LevelTree "D:\references\00_level tree.pdf" `
  -QuantityGuideline "D:\references\ss-a-0804(0521kaitei).pdf" `
  -NationalReferenceXlsx "D:\references\20250319_sekopsankou0804.xlsx" `
  -NationalPackagePdf "D:\references\20260319_sekoptanka0804.pdf" `
  -IshikawaPackageXlsx "D:\references\20260319_sekoptankaishikawa0804.xlsx" `
  -IshikawaPackagePdf "D:\references\20260319_sekoptankaishikawa0804.pdf"
```

Convert a quantity workbook. `PriceDate` must be the official 積算年月 or applicable date, not the workbook filename date unless those are confirmed to be the same.

```powershell
.\run_converter.ps1 `
  -Source "D:\projects\数量計算書.xlsx" `
  -Template "D:\templates\accepted_GAIA金抜.xlsx" `
  -PriceDate "2025-10-24" `
  -TreeCategory "河川改修" `
  -Location "珠洲市〇〇町地内" `
  -Output ".\outputs\project_GAIA取込候補.xlsx"
```

Open the generated workbook and its adjacent `_review.csv`. Only `EXACT_CODE` and `PACKAGE_EXACT` rows are candidates for automated GAIA selection. All other statuses require review.

## Safe GAIA trial

1. Create or copy a disposable test project in GAIA. Do not begin with a live estimate.
2. Use the same Excel import operation that accepts your existing 金抜 workbook and select the generated candidate.
3. Confirm that 費目・工種・種別・細別, quantities, and units were imported in the expected hierarchy.
4. For `EXACT_CODE` and `PACKAGE_EXACT`, confirm GAIA selects the expected 施工パッケージ and displays the correct condition screen.
5. Confirm GAIA supplies prices for the intended region and month. The converter deliberately leaves prices blank.
6. Compare several representative lines manually before attempting UI automation.

The exact GAIA import menu varies by installation/version, so the first trial should record the menu name, importer messages, and any rejected rows. Those results will define the next automation step.

## Conversion model

| Quantity workbook | Normalized field | GAIA candidate |
| --- | --- | --- |
| Sheet suffix, e.g. `上部工` | Section | `費目` |
| A: 工種 | Work type | `工種` |
| B: 種別 | Category | `種別` |
| C: 細別 | Item name | `細別` |
| D: 規格 | Specification | Condition line |
| E: 単位 | Unit | Unit |
| F: 数量 | Quantity | Quantity |
| H-K | Standard, settings, notes | Match evidence and review status |

The source workbook remains unchanged. The generated workbook is a candidate for a controlled GAIA import test, not a guaranteed official import format.

## Governing-reference checks

- `積算体系ツリー`: validates that the preserved 工種・種別・細別 hierarchy can be found in the municipal tree context.
- `土木工事数量算出要領`: adds page evidence for the quantity classification. It does not recalculate geometric quantities from detailed calculation sheets.
- `施工パッケージ`: matches the exact package name, unit, condition fields, scope, fiscal year, and effective date. The R8 Ishikawa disaster-area table has priority for 珠洲市 where it contains the package; otherwise the current national table is used.

施工パッケージ numbers such as `018` are reference-table numbers, not GAIA `CB...` codes. A GAIA code is written only when it comes from an independently verified rule in `mapping_rules.csv`. Reference prices are never copied into the 金抜 workbook; GAIA remains responsible for the applicable regional/monthly unit prices.

## Safety classifications

- `EXACT_CODE`: all reference gates pass and an approved rule supplies a GAIA code.
- `PACKAGE_EXACT`: the current package, unit, hierarchy, guideline evidence, and conditions pass, but no verified GAIA code is available.
- `PACKAGE_DATE_REVIEW`: 積算年月 is missing or outside the reference period.
- `PACKAGE_YEAR_REVIEW`: the workbook's stated standard year differs from the package year.
- `PACKAGE_UNIT_REVIEW`: the workbook and package units differ.
- `PACKAGE_CONDITION_REVIEW`: explicit package condition settings are missing.
- `TREE_REVIEW` / `QUANTITY_RULE_REVIEW`: reference-page evidence is insufficient.
- `TREE_BRANCH_REVIEW`: the project branch such as 河川改修 or 道路維持・修繕 was not supplied.
- `MARKET_PRICE_REVIEW`: licensed market-price data and applicable month must be checked.
- `QUOTATION_REQUIRED`: quotation pricing is required.
- `BLOCKED_REVIEW`: the source explicitly says the standard is not applicable.
- `MANUAL_REVIEW`: no exact package can be safely confirmed.
- `INVALID_QUANTITY`: quantity is missing or non-positive and is excluded.

## Build the reference index

Run this only when the governing reference files change:

```powershell
python build_reference_index.py `
  --level-tree "path\to\00_level tree.pdf" `
  --quantity-guideline "path\to\ss-a-0804(0521kaitei).pdf" `
  --national-reference-xlsx "path\to\20250319_sekopsankou0804.xlsx" `
  --national-package-pdf "path\to\20260319_sekoptanka0804.pdf" `
  --ishikawa-package-xlsx "path\to\20260319_sekoptankaishikawa0804.xlsx" `
  --ishikawa-package-pdf "path\to\20260319_sekoptankaishikawa0804.pdf" `
  --output "references\suzu_reference_index.json"
```

## Convert a quantity workbook

```powershell
python gaia_converter.py `
  --source "path\to\数量計算書.xlsx" `
  --template "path\to\accepted_金抜設計書.xlsx" `
  --rules mapping_rules.csv `
  --reference-index "references\suzu_reference_index.json" `
  --tree-category "河川改修" `
  --output "outputs\project_GAIA取込候補.xlsx" `
  --location "珠洲市〇〇町地内" `
  --price-date 2026-07-01
```

Review the generated CSV before importing. Confirm the location, work category, district, price date, transport distances, correction rules, market-price contracts, package condition fields, and every item that is not `EXACT_CODE` or `PACKAGE_EXACT`.
