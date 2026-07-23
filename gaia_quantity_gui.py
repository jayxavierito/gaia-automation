from __future__ import annotations

import json
import os
import sys
import threading
import traceback
from datetime import date, datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

from gaia_converter import (
    apply_mapping_rules,
    apply_reference_index,
    boq_items_from_extraction,
    build_gaia_candidate,
    classify_item,
    detect_estimate_date,
    load_mapping_rules,
    load_reference_index,
    write_review_csv,
)
from quantity_extractors import extract_quantity_source


APP_TITLE = "GAIA 数量計算書コンバーター"
SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm", ".pdf"}


def resource_path(*parts: str) -> Path:
    root = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return root.joinpath(*parts)


def first_existing_path(*paths: Path) -> Path | None:
    return next((path for path in paths if path.is_file()), None)


def available_output_path(directory: Path, stem: str, suffix: str) -> Path:
    candidate = directory / f"{stem}{suffix}"
    if not candidate.exists():
        return candidate
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return directory / f"{stem}_{timestamp}{suffix}"


def convert_to_gaia_files(source_path: Path, output_directory: Path):
    output_directory.mkdir(parents=True, exist_ok=True)
    work_directory = (
        output_directory / "_OCR確認資料"
        if source_path.suffix.lower() == ".pdf"
        else None
    )
    result = extract_quantity_source(source_path, work_dir=work_directory)
    items = boq_items_from_extraction(result)
    price_date, price_date_source = detect_estimate_date(source_path)

    rules_path = first_existing_path(
        resource_path("mapping_rules.csv"),
        Path(__file__).resolve().parent / "mapping_rules.csv",
    )
    apply_mapping_rules(items, load_mapping_rules(rules_path))

    reference_path = first_existing_path(
        resource_path("references", "suzu_reference_index.json"),
        resource_path("assets", "suzu_level_tree_index.json"),
        Path(__file__).resolve().parent
        / "references"
        / "suzu_reference_index.json",
        Path(__file__).resolve().parent
        / "assets"
        / "suzu_level_tree_index.json",
    )
    reference_index = load_reference_index(reference_path)
    if reference_index:
        apply_reference_index(
            items,
            reference_index,
            price_date=price_date,
            district="珠洲",
            tree_category="",
            project_name=result.project_name,
        )
    for item in items:
        classify_item(item, reference_index_active=bool(reference_index))

    template_path = first_existing_path(
        resource_path("assets", "gaia_suzu_7sheet_template.xlsx"),
        Path(__file__).resolve().parent
        / "assets"
        / "gaia_suzu_7sheet_template.xlsx",
    )
    if template_path is None:
        raise FileNotFoundError(
            "GAIA取込テンプレートがありません。"
            "assets\\gaia_suzu_7sheet_template.xlsx を確認してください。"
        )

    xlsx_path = available_output_path(
        output_directory, f"{source_path.stem}_GAIA取込用", ".xlsx"
    )
    review_path = xlsx_path.with_name(f"{xlsx_path.stem}_確認.csv")
    json_path = xlsx_path.with_name(f"{xlsx_path.stem}_抽出結果.json")
    write_review_csv(review_path, items)
    block_count = build_gaia_candidate(
        template_path=template_path,
        output_path=xlsx_path,
        project_name=result.project_name,
        items=items,
        location="",
        work_category="橋梁工事",
        district="珠洲",
        price_date=price_date,
    )
    json_path.write_text(
        json.dumps(
            {
                "extraction": result.to_dict(),
                "gaia": {
                    "output": str(xlsx_path.resolve()),
                    "review": str(review_path.resolve()),
                    "template": str(template_path.resolve()),
                    "reference_index": (
                        str(reference_path.resolve()) if reference_path else ""
                    ),
                    "price_date": price_date.isoformat(),
                    "price_date_source": price_date_source,
                    "tree_category": items[0].tree_category if items else "",
                    "level4_verified": sum(
                        item.tree_status == "LEVEL4_VERIFIED" for item in items
                    ),
                    "output_blocks": block_count,
                },
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    return result, items, xlsx_path, review_path, price_date, price_date_source


class ConverterApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("860x680")
        self.root.minsize(760, 620)
        self.root.configure(bg="#F3F0E8")

        self.source_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.status_var = tk.StringVar(value="数量計算書を選択してください。")
        self.result_var = tk.StringVar()
        self.last_xlsx: Path | None = None
        self.last_csv: Path | None = None
        self.last_output_directory: Path | None = None

        self._build_ui()

    def _build_ui(self) -> None:
        header = tk.Frame(self.root, bg="#123B4A", padx=34, pady=24)
        header.pack(fill="x")
        tk.Label(
            header,
            text=APP_TITLE,
            bg="#123B4A",
            fg="white",
            font=("Yu Gothic UI", 22, "bold"),
        ).pack(anchor="w")
        tk.Label(
            header,
            text="Excel・PDFの数量計算書から、GAIA取込用Excelを作成します",
            bg="#123B4A",
            fg="#D9E6E8",
            font=("Yu Gothic UI", 11),
        ).pack(anchor="w", pady=(6, 0))

        content = tk.Frame(self.root, bg="#F3F0E8", padx=34, pady=26)
        content.pack(fill="both", expand=True)

        self._section_label(content, "1  数量計算書を選ぶ")
        source_row = tk.Frame(content, bg="#F3F0E8")
        source_row.pack(fill="x", pady=(8, 22))
        self.source_entry = tk.Entry(
            source_row,
            textvariable=self.source_var,
            font=("Yu Gothic UI", 11),
            relief="solid",
            bd=1,
            bg="white",
        )
        self.source_entry.pack(side="left", fill="x", expand=True, ipady=9)
        self.source_button = tk.Button(
            source_row,
            text="ファイルを選ぶ",
            command=self.select_source,
            bg="#E3A626",
            fg="#1E2528",
            activebackground="#F0BC4D",
            font=("Yu Gothic UI", 11, "bold"),
            relief="flat",
            padx=18,
            pady=10,
            cursor="hand2",
        )
        self.source_button.pack(side="left", padx=(12, 0))

        self._section_label(content, "2  保存先を確認する")
        output_row = tk.Frame(content, bg="#F3F0E8")
        output_row.pack(fill="x", pady=(8, 22))
        self.output_entry = tk.Entry(
            output_row,
            textvariable=self.output_var,
            font=("Yu Gothic UI", 11),
            relief="solid",
            bd=1,
            bg="white",
        )
        self.output_entry.pack(side="left", fill="x", expand=True, ipady=9)
        self.output_button = tk.Button(
            output_row,
            text="保存先を変更",
            command=self.select_output_directory,
            bg="#DCE5E3",
            fg="#123B4A",
            activebackground="#C9D8D5",
            font=("Yu Gothic UI", 10, "bold"),
            relief="flat",
            padx=16,
            pady=10,
            cursor="hand2",
        )
        self.output_button.pack(side="left", padx=(12, 0))

        self.convert_button = tk.Button(
            content,
            text="3  GAIA取込Excelを作成",
            command=self.start_conversion,
            bg="#C84A2F",
            fg="white",
            activebackground="#A93D28",
            activeforeground="white",
            disabledforeground="#E5E5E5",
            font=("Yu Gothic UI", 16, "bold"),
            relief="flat",
            pady=15,
            cursor="hand2",
        )
        self.convert_button.pack(fill="x", pady=(0, 16))

        status_frame = tk.Frame(content, bg="#E8E3D8", padx=18, pady=14)
        status_frame.pack(fill="x")
        tk.Label(
            status_frame,
            textvariable=self.status_var,
            bg="#E8E3D8",
            fg="#283438",
            font=("Yu Gothic UI", 11, "bold"),
            justify="left",
            anchor="w",
            wraplength=730,
        ).pack(fill="x")

        self.result_frame = tk.Frame(content, bg="#DDEBDD", padx=18, pady=16)
        tk.Label(
            self.result_frame,
            textvariable=self.result_var,
            bg="#DDEBDD",
            fg="#173E2C",
            font=("Yu Gothic UI", 11, "bold"),
            justify="left",
            anchor="w",
            wraplength=720,
        ).pack(fill="x")
        result_buttons = tk.Frame(self.result_frame, bg="#DDEBDD")
        result_buttons.pack(fill="x", pady=(12, 0))
        tk.Button(
            result_buttons,
            text="GAIA取込Excelを開く",
            command=self.open_xlsx,
            bg="#1E6B51",
            fg="white",
            activebackground="#185740",
            activeforeground="white",
            font=("Yu Gothic UI", 10, "bold"),
            relief="flat",
            padx=16,
            pady=9,
            cursor="hand2",
        ).pack(side="left")
        tk.Button(
            result_buttons,
            text="確認表を開く",
            command=self.open_csv,
            bg="#FFFFFF",
            fg="#1E6B51",
            activebackground="#F4F7F5",
            font=("Yu Gothic UI", 10, "bold"),
            relief="flat",
            padx=16,
            pady=9,
            cursor="hand2",
        ).pack(side="left", padx=(10, 0))
        tk.Button(
            result_buttons,
            text="結果フォルダーを開く",
            command=self.open_output_directory,
            bg="#FFFFFF",
            fg="#1E6B51",
            activebackground="#F4F7F5",
            font=("Yu Gothic UI", 10, "bold"),
            relief="flat",
            padx=16,
            pady=9,
            cursor="hand2",
        ).pack(side="left", padx=(10, 0))

        tk.Label(
            content,
            text=(
                "GAIAには「_GAIA取込用.xlsx」を選択してください。CSVは確認用です。"
                "PDFの名称・規格・単位・数量は、必ず原本と照合してください。"
            ),
            bg="#F3F0E8",
            fg="#6B4B21",
            font=("Yu Gothic UI", 9),
            justify="left",
            anchor="w",
            wraplength=750,
        ).pack(fill="x", pady=(18, 0))

    @staticmethod
    def _section_label(parent: tk.Widget, text: str) -> None:
        tk.Label(
            parent,
            text=text,
            bg="#F3F0E8",
            fg="#123B4A",
            font=("Yu Gothic UI", 12, "bold"),
            anchor="w",
        ).pack(fill="x")

    def select_source(self) -> None:
        filename = filedialog.askopenfilename(
            title="数量計算書を選択",
            filetypes=[
                ("対応ファイル", "*.xlsx *.xlsm *.pdf"),
                ("Excel", "*.xlsx *.xlsm"),
                ("PDF", "*.pdf"),
            ],
        )
        if filename:
            self.set_source(Path(filename))

    def set_source(self, source_path: Path) -> None:
        self.source_var.set(str(source_path))
        if not self.output_var.get().strip():
            self.output_var.set(str(source_path.parent / "GAIA変換結果"))
        self.status_var.set("ファイルを確認しました。「変換する」を押してください。")
        self.result_frame.pack_forget()

    def select_output_directory(self) -> None:
        initial = self.output_var.get().strip() or str(Path.home() / "Documents")
        directory = filedialog.askdirectory(title="保存先を選択", initialdir=initial)
        if directory:
            self.output_var.set(directory)

    def start_conversion(self) -> None:
        source_path = Path(self.source_var.get().strip())
        output_directory = Path(self.output_var.get().strip())
        if not source_path.is_file():
            messagebox.showerror(APP_TITLE, "数量計算書を選択してください。")
            return
        if source_path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            messagebox.showerror(
                APP_TITLE,
                "対応形式は .xlsx、.xlsm、.pdf です。",
            )
            return
        if not str(output_directory):
            messagebox.showerror(APP_TITLE, "保存先を選択してください。")
            return

        self._set_busy(True)
        self.result_frame.pack_forget()
        if source_path.suffix.lower() == ".pdf":
            self.status_var.set("PDFの数量表を読取り中です。しばらくお待ちください…")
        else:
            self.status_var.set("Excelの数量総括表を読取り中です…")
        threading.Thread(
            target=self._convert_worker,
            args=(source_path, output_directory),
            daemon=True,
        ).start()

    def _convert_worker(self, source_path: Path, output_directory: Path) -> None:
        try:
            (
                result,
                items,
                xlsx_path,
                review_path,
                price_date,
                price_date_source,
            ) = convert_to_gaia_files(source_path, output_directory)
            review_count = sum(
                item.extraction_status != "READY"
                or item.tree_status != "LEVEL4_VERIFIED"
                for item in items
            )
            verified_count = sum(
                item.tree_status == "LEVEL4_VERIFIED" for item in items
            )
            self.root.after(
                0,
                self._conversion_succeeded,
                result.project_name,
                len(result.records),
                review_count,
                verified_count,
                xlsx_path,
                review_path,
                output_directory,
                price_date,
                price_date_source,
            )
        except Exception as error:
            output_directory.mkdir(parents=True, exist_ok=True)
            error_log = output_directory / "変換エラー.txt"
            error_log.write_text(traceback.format_exc(), encoding="utf-8")
            self.root.after(
                0,
                self._conversion_failed,
                str(error),
                error_log,
            )

    def _conversion_succeeded(
        self,
        project_name: str,
        record_count: int,
        review_count: int,
        verified_count: int,
        xlsx_path: Path,
        review_path: Path,
        output_directory: Path,
        price_date: date,
        price_date_source: str,
    ) -> None:
        self._set_busy(False)
        self.last_xlsx = xlsx_path
        self.last_csv = review_path
        self.last_output_directory = output_directory
        if review_count:
            review_text = f"\n要確認: {review_count}行（確認表を原本と照合）"
        else:
            review_text = "\n要確認: なし"
        date_note = "原本" if price_date_source == "SOURCE" else "PC日付"
        self.result_var.set(
            f"GAIA取込Excelを作成しました。\n工事名: {project_name}"
            f"\n明細数: {record_count}行 / レベル4確認済: {verified_count}行"
            f"\n積算年月: {price_date:%Y年%m月}（{date_note}）"
            f"{review_text}\nGAIAで選ぶファイル: {xlsx_path}"
        )
        self.status_var.set(
            "完了しました。GAIAでは「_GAIA取込用.xlsx」を選択してください。"
        )
        self.result_frame.pack(fill="x", pady=(14, 0))

    def _conversion_failed(self, message: str, error_log: Path) -> None:
        self._set_busy(False)
        self.status_var.set("変換できませんでした。エラー内容を確認してください。")
        messagebox.showerror(
            APP_TITLE,
            f"変換できませんでした。\n\n{message}\n\n詳細: {error_log}",
        )

    def _set_busy(self, busy: bool) -> None:
        state = "disabled" if busy else "normal"
        for widget in (
            self.source_button,
            self.output_button,
            self.convert_button,
            self.source_entry,
            self.output_entry,
        ):
            widget.configure(state=state)
        self.root.configure(cursor="wait" if busy else "")

    def open_csv(self) -> None:
        if self.last_csv and self.last_csv.exists():
            os.startfile(self.last_csv)

    def open_xlsx(self) -> None:
        if self.last_xlsx and self.last_xlsx.exists():
            os.startfile(self.last_xlsx)

    def open_output_directory(self) -> None:
        if self.last_output_directory and self.last_output_directory.exists():
            os.startfile(self.last_output_directory)


def main() -> int:
    if len(sys.argv) >= 3 and sys.argv[1] == "--gui-smoke-test":
        report_path = Path(sys.argv[2])
        root = tk.Tk()
        ConverterApp(root)
        root.update_idletasks()
        root.update()
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(
            json.dumps(
                {
                    "title": root.title(),
                    "geometry": root.geometry(),
                    "state": root.state(),
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        root.destroy()
        return 0

    if len(sys.argv) >= 3 and sys.argv[1] == "--batch":
        source_path = Path(sys.argv[2])
        output_directory = (
            Path(sys.argv[4])
            if len(sys.argv) >= 5 and sys.argv[3] == "--output"
            else source_path.parent / "GAIA変換結果"
        )
        try:
            (
                result,
                items,
                xlsx_path,
                review_path,
                price_date,
                price_date_source,
            ) = convert_to_gaia_files(source_path, output_directory)
            summary_path = output_directory / "batch_summary.json"
            summary_path.write_text(
                json.dumps(
                    {
                        "project_name": result.project_name,
                        "records": len(result.records),
                        "source_format": result.source_format,
                        "xlsx": str(xlsx_path.resolve()),
                        "review": str(review_path.resolve()),
                        "price_date": price_date.isoformat(),
                        "price_date_source": price_date_source,
                        "level4_verified": sum(
                            item.tree_status == "LEVEL4_VERIFIED"
                            for item in items
                        ),
                    },
                    ensure_ascii=False,
                    indent=2,
                ),
                encoding="utf-8",
            )
            return 0
        except Exception:
            output_directory.mkdir(parents=True, exist_ok=True)
            (output_directory / "変換エラー.txt").write_text(
                traceback.format_exc(), encoding="utf-8"
            )
            return 1

    root = tk.Tk()
    app = ConverterApp(root)
    if len(sys.argv) > 1:
        dropped_path = Path(sys.argv[1].strip('"'))
        if dropped_path.is_file():
            app.set_source(dropped_path)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
