# 珠洲市向け GAIA 数量計算書コンバーター

数量計算書（Excel または PDF）から数量総括表を読み取り、共通形式へ整理したうえで、GAIA 用の金抜設計書候補を作成するためのツールです。

変換処理は GAIA 本体から独立しており、GAIA を起動・操作しません。`gaia_ui_automation.py` は別の試験用ツールであり、通常の変換処理では使用しません。

## 現在の推奨運用

通常は次の3段階で使用します。

1. 数量計算書を選び、GAIA 取込用 Excel と確認用 CSV を同時に作成する。
2. 確認用 CSV の名称・規格・単位・数量・レベル判定を原本と照合する。
3. GAIA で `_GAIA取込用.xlsx` を手動選択して取り込む。

特に PDF は埋め込み文字または OCR から抽出するため、確認前のデータを自動的に GAIA コードや施工パッケージへ確定させません。
GAIA の画面操作は自動化せず、人が取込ファイルを選択します。

## 対応ファイル

- Excel `.xlsx` / `.xlsm`
  - `工種・種別・細別・規格（規格等）・単位・数量` の見出しから列位置を自動判定します。
  - 列の位置や開始行が変わっても対応できます。
  - 結合セル、省略された上位区分、数量が次の行に記載されている形式にも対応します。
- PDF `.pdf`
  - 文字情報を持つ PDF は、複数ページの `数量総括表`・`数量集計表` を直接読み取ります。
  - 文字情報のないスキャン PDF は、罫線表を検出して Windows 日本語 OCR で読み取ります。
- 未登録の形式
  - 無理に推測せず、診断メッセージを表示して停止します。
  - 新しい形式は、実例をテストデータとして追加して対応範囲を広げます。

「どのような書式でも100%自動変換」を安全に保証することはできません。その代わり、未知の書式で誤った GAIA データを作らない設計にしています。

## 他の PC で使う最も簡単な方法

### 推奨: Portable EXE を ZIP で配布する

Python をインストールしてもらうのではなく、`GaiaQuantityConverter-win64.zip` を配布する方法を推奨します。

受け取った人の操作は次のとおりです。

1. ZIP を右クリックして「すべて展開」します。
2. 展開したフォルダー内の `GaiaQuantityConverter.exe` をダブルクリックします。
3. 「ファイルを選ぶ」で数量計算書を選択します。
4. 「GAIA取込Excelを作成」を押します。
5. 完了後、「確認表を開く」で原本との照合を行います。
6. GAIA では、結果フォルダーの `_GAIA取込用.xlsx` を選択します。CSV は取込ファイルではありません。

数量計算書を EXE のアイコンへドラッグして起動することもできます。

重要事項:

- EXE だけをフォルダー外へ移動しないでください。展開したフォルダー全体で使用します。
- 元の Excel・PDF は変更しません。
- Python のインストールは不要です。
- Windows 10 / 11（64ビット）を対象とします。
- 画像 PDF を使う PC では、Windows の日本語 OCR 言語が利用可能である必要があります。
- EXE は GAIA 取込用 Excel、確認用 CSV、抽出結果 JSON を作成します。
- GAIA 本体への入力・画面操作は行いません。
- 署名していない社内試用版 EXE では、Windows SmartScreen や会社のセキュリティ製品による確認が表示される場合があります。本運用前にはコード署名を推奨します。

単一ファイルの `onefile` 形式ではなく、フォルダー形式の `onedir` を採用しています。起動が速く、PDF 用 DLL を確実に同梱でき、ウイルス対策ソフトの誤検知も比較的少ないためです。

### Portable EXE の作成方法

開発用 PC で一度だけ実行します。

```powershell
cd "C:\Users\chinb\Documents\GAIA automation"
Set-ExecutionPolicy -Scope Process Bypass
.\setup.ps1
.\.venv\Scripts\python.exe -m pip install -r requirements-build.txt
.\build_portable_exe.ps1
```

作成物:

- 実行ファイル: `dist\GaiaQuantityConverter\GaiaQuantityConverter.exe`
- 配布用 ZIP: `dist\GaiaQuantityConverter-win64.zip`

配布するときは ZIP をそのまま渡してください。

## 開発用 PC のセットアップ

Python から直接実行・修正する場合の手順です。

必要環境:

- Windows 10 / 11
- Python 3.11 以上
- 出力確認用の Microsoft Excel
- 画像 PDF を使う場合は Windows 日本語 OCR
- GAIA 取込用 Excel の確認に使用する Microsoft Excel

```powershell
git clone YOUR_GITHUB_REPOSITORY_URL
cd gaia-automation
Set-ExecutionPolicy -Scope Process Bypass
.\setup.ps1
```

## まず確認用データを作る

### Excel

```powershell
.\run_quantity_extract.ps1 `
  -Source "D:\projects\数量計算書.xlsx" `
  -Output ".\outputs\project_normalized.csv"
```

### PDF

```powershell
.\run_quantity_extract.ps1 `
  -Source "D:\projects\数量計算書.pdf" `
  -Output ".\outputs\project_normalized.csv"
```

文字情報を持つ PDF は、埋め込み文字と罫線から複数ページの数量総括表・数量集計表を抽出します。画像 PDF は自動的に Windows OCR に切り替わり、ページ画像・OCR セル画像が診断用として `tmp\quantity-extraction` に保存されます。

確認用 CSV では、最低限次の列を原本と照合してください。

- `work_type`: 工種
- `category`: 種別
- `item_name`: 細別
- `specification`: 規格等
- `unit`: 単位
- `quantity`: 数量
- `extraction_status`: 抽出時の確認状態

PDF の行は、文字 PDF では `PDF_TEXT_REVIEW_REQUIRED`、画像 PDF では `OCR_REVIEW_REQUIRED` になります。GAIA 用レビュー CSV ではどちらも `SOURCE_REVIEW_REQUIRED` となり、原本確認が終わるまで自動確定されません。A1・A2 など複数の数量列に `合計` があれば合計を優先し、合計列が無い場合は列別の明細として残します。

## 珠洲市の基準資料を登録する

次の資料が変更されたときだけ、参照インデックスを再作成します。

```powershell
.\build_references.ps1 `
  -LevelTree "D:\references\00_level tree.pdf" `
  -QuantityGuideline "D:\references\ss-a-0804(0521kaitei).pdf" `
  -NationalReferenceXlsx "D:\references\20250319_sekopsankou0804.xlsx" `
  -NationalPackagePdf "D:\references\20260319_sekoptanka0804.pdf" `
  -IshikawaPackageXlsx "D:\references\20260319_sekoptankaishikawa0804.xlsx" `
  -IshikawaPackagePdf "D:\references\20260319_sekoptankaishikawa0804.pdf"
```

生成される `references\suzu_reference_index.json` には、施工パッケージ・数量算出要領の索引とローカルファイル情報が含まれるため Git には登録しません。

Portable EXE と GitHub 版には、個別ファイルパスを含まない `assets\suzu_level_tree_index.json` を同梱します。これは積算体系ツリーのレベル1～4確認に使用します。施工パッケージまで照合する社内版を作る場合だけ、完全版索引を安全な社内経路で配布してください。

## GAIA 取込候補 Excel を作る

GUI / Portable EXE では、原本に `積算年月` または `単価適用年月` があればその年月を使用し、見つからなければ PC の当日を使用します。GAIA 取込後に人が修正できます。

コマンドから明示する場合だけ `PriceDate` を指定します。省略時は GUI と同じ自動判定です。

```powershell
.\run_converter.ps1 `
  -Source "D:\projects\数量計算書.xlsx" `
  -Location "珠洲市〇〇町地内" `
  -Output ".\outputs\project_GAIA取込候補.xlsx"
```

積算年月や事業区分を明示する場合は、`-PriceDate "2026-07-01"`、`-TreeCategory "道路新設・改築"` を追加します。

生成物:

- `project_GAIA取込候補.xlsx`: GAIA 取込候補
- `project_GAIA取込候補_review.csv`: 判定根拠・警告・確認状態

`run_converter.ps1` は、承認済みテンプレート内の GAIA コード履歴も参照します。ただし、次の条件をすべて満たす場合だけコードを補完します。

- 現年度の施工パッケージ名称が完全一致する。
- 単位が一致または安全な等価単位である。
- 同じ名称・単位の過去データが1つの GAIA コードに確定する。

見積、市場単価、単位不一致、複数候補は自動補完しません。

## 変換項目の対応

| 数量計算書 | 共通データ | GAIA 取込候補 |
| --- | --- | --- |
| 数量総括シート・表題 | Section | レベル1 / 費目 |
| `工種` | Work type | レベル2 / 工種 |
| `種別` | Category | レベル3 / 種別 |
| `細別` | Item name | レベル4 / 細別 |
| `規格` / `規格等` | Specification | 条件・規格行 |
| `単位` | Unit | 単位 |
| `数量` | Quantity | 数量 |
| 積算基準・設定・備考 | Match evidence | 判定根拠・警告 |

元の数量計算書は変更しません。生成した Excel は、正式な設計書ではなく確認・取込試験用の候補です。

## 珠洲市基準の確認内容

- `積算体系ツリー`
  - レベル1 工事区分、レベル2 工種、レベル3 種別、レベル4 細別の経路を確認します。
  - 上位経路が複数ある名称は自動確定せず、確認用 CSV に `AMBIGUOUS` と記録します。
- `土木工事数量算出要領`
  - 数量区分に関係するページを根拠として記録します。
  - 詳細な幾何計算を再計算する機能ではありません。
- `施工パッケージ`
  - 名称、単位、条件区分、対象地域、年度、適用開始日を確認します。
  - 珠洲市を含む令和8年度石川県災害地区資料に同じパッケージがある場合は、そちらを優先します。

施工パッケージ資料の番号（例: `018`）は GAIA の `CB...` コードではありません。GAIA コードは、独立して確認済みのルールまたは承認済み GAIA ファイル内の一意な履歴からのみ設定します。

参考単価は金抜設計書へコピーしません。地域・月・補正条件に応じた価格の確定は GAIA 側で行います。

## 判定ステータス

- `SOURCE_REVIEW_REQUIRED`: PDF/OCR 行。原本照合が必要です。
- `EXACT_CODE`: 参照条件を満たし、確認済み GAIA コードがあります。
- `PACKAGE_EXACT`: 施工パッケージ、単位、階層、条件が一致しますが、確認済み GAIA コードはありません。
- `PACKAGE_DATE_REVIEW`: 積算年月が未指定、または適用期間外です。
- `PACKAGE_YEAR_REVIEW`: 数量計算書と施工パッケージの年度が一致しません。
- `PACKAGE_UNIT_REVIEW`: 数量計算書と施工パッケージの単位が一致しません。
- `PACKAGE_CONDITION_REVIEW`: 条件区分の設定内容が不足しています。
- `TREE_REVIEW`: 積算体系ツリーで十分な根拠を確認できません。
- `LEVEL4_VERIFIED`: レベル1～4の経路を一意に確認できました。
- `LEVEL4_AMBIGUOUS_PATH`: レベル4名称はありますが、上位経路が複数あります。
- `LEVEL3_VERIFIED_LEVEL4_UNVERIFIED`: レベル3までは確認でき、レベル4は原文のまま要確認です。
- `LEVEL3_AMBIGUOUS_PATH`: レベル3の上位経路が複数あります。
- `QUANTITY_RULE_REVIEW`: 数量算出要領で十分な根拠を確認できません。
- `TREE_BRANCH_REVIEW`: 河川改修、道路維持・修繕などの事業区分が未指定です。
- `MARKET_PRICE_REVIEW`: 市販単価データと適用月の確認が必要です。
- `QUOTATION_REQUIRED`: 見積価格が必要です。
- `BLOCKED_REVIEW`: 原資料に積算基準の適用不可と記載されています。
- `MANUAL_REVIEW`: 完全一致する施工パッケージを安全に確定できません。
- `INVALID_QUANTITY`: 数量が未取得または0以下のため、出力対象外です。

## GAIA での手動確認

承認済みの7シート形式（`鏡`、`本工事費内訳表`、`内訳書`、`明細書`、`代価表`、`単価表`、`施工パッケージ`）を内蔵テンプレートの基準にしています。

新しい原本形式を追加したときは、次を手動確認します。

1. 結果フォルダーの `_GAIA取込用.xlsx` を GAIA で選択する。
2. `珠洲市Excel`、`珠洲市`、`一般土木`、`金抜き` として認識されることを確認する。
3. 費目・工種・種別・細別、規格、単位、数量を原本と比較する。
4. 単価、歩掛り、補正条件、経費区分は GAIA 側で確認・設定する。

磐若橋 PDF の試験出力は構造・表示・セルエラーを確認済みですが、GAIA での最終受入は手動確認待ちです。

## 安全な GAIA 取込試験

1. 本番工事ではなく、GAIA 内に使い捨てのテスト工事を作成します。
2. 生成した GAIA 取込候補 Excel を読み込みます。
3. 費目・工種・種別・細別、数量、単位の階層を確認します。
4. `EXACT_CODE` と `PACKAGE_EXACT` でも、GAIA の条件画面を確認します。
5. 地区・積算月に合う単価が設定されることを確認します。
6. 代表行を手作業の積算結果と比較してから対象範囲を広げます。

## GAIA UI 自動操作について

通常運用では使用しません。GAIA のファイル選択、取込、歩掛り・単価設定は人が行います。Portable EXE は GAIA を起動・操作しません。

## テスト

```powershell
.\.venv\Scripts\python.exe -m unittest -v
```

現在の自動テストには、次が含まれます。

- 見出し位置が変わる Excel 数量総括表
- 省略階層・次行数量・合計行の処理
- PDF 埋め込み文字表・罫線 OCR 行の確認必須化
- 施工パッケージ年度・単位・条件区分の安全判定
- GAIA コード履歴の一意性確認

## 今後の UI

Portable EXE は、GAIA 取込用 Excel と確認用データを作る簡易 UI です。実例を増やして抽出精度を確認した後、次を追加する予定です。

- 確認が必要なセルを色付きで表示するレビュー画面
- 原本 PDF と抽出行の並列表示
- 承認済み修正を次回へ反映する学習辞書
- 正式な積算年月・事業区分を必要に応じて上書きする画面

GAIA 本体の UI 自動操作は、コンバーターとは分離したままにします。
