from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import time
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable

from gaia_reconcile import (
    load_source_review,
    parse_gaia_review_workbook,
    reconcile_rows,
    write_reconciliation,
)


TEST_PROJECT_PREFIX = "GAIA自動化テスト_"
DEFAULT_GAIA_EXE = Path(r"D:\Program Files (x86)\GaiaCloud\GaiaCloud.exe")
DEFAULT_CACHE_ROOT = Path(r"D:\ProgramData\GaiaCloud\DB\User")


@dataclass(frozen=True)
class TrialConfig:
    candidate: Path
    source_review: Path
    project_name: str
    folder_name: str
    output_dir: Path
    gaia_exe: Path
    cache_root: Path
    timeout_seconds: int
    confirm_start: bool
    export_review: bool


def validate_trial_config(config: TrialConfig) -> None:
    if not config.candidate.exists() or config.candidate.suffix.lower() != ".xlsx":
        raise ValueError(f"GAIA取込候補Excelがありません: {config.candidate}")
    if not config.source_review.exists():
        raise ValueError(f"変換レビューCSVがありません: {config.source_review}")
    if not config.gaia_exe.exists():
        raise ValueError(f"GAIA Cloud実行ファイルがありません: {config.gaia_exe}")
    if not config.project_name.startswith(TEST_PROJECT_PREFIX):
        raise ValueError(
            f"技術試験の工事名は {TEST_PROJECT_PREFIX} で開始してください"
        )
    if not config.confirm_start:
        raise ValueError("実行には --confirm-start が必要です")
    if config.timeout_seconds < 60:
        raise ValueError("--timeout-seconds は60秒以上にしてください")


def import_pywinauto():
    try:
        from pywinauto import Application, Desktop
    except ImportError as error:
        raise RuntimeError(
            "pywinautoがありません。setup.ps1を実行してください。"
        ) from error
    return Application, Desktop


def control_texts(window) -> list[str]:
    texts: list[str] = []
    for control in [window, *window.descendants()]:
        try:
            text = control.window_text().strip()
        except Exception:
            continue
        if text and text not in texts:
            texts.append(text)
    return texts


def wait_for_window(Desktop, required_text: str, timeout: int):
    deadline = time.monotonic() + timeout
    last_titles: list[str] = []
    while time.monotonic() < deadline:
        for window in Desktop(backend="uia").windows():
            try:
                if not window.is_visible():
                    continue
                texts = control_texts(window)
            except Exception:
                continue
            title = window.window_text().strip()
            if title:
                last_titles.append(title)
            if any(required_text in text for text in texts):
                return window
        time.sleep(1)
    raise TimeoutError(
        f"'{required_text}' を含むGAIA画面を待機しましたが見つかりません。"
        f"表示中タイトル: {sorted(set(last_titles))[-10:]}"
    )


def find_control(window, labels: Iterable[str], control_types: set[str] | None = None):
    labels = tuple(labels)
    controls = []
    for control in window.descendants():
        try:
            controls.append(
                (control, control.window_text().strip(), control.element_info.control_type)
            )
        except Exception:
            continue
    for exact in (True, False):
        for control, text, control_type in controls:
            if control_types and control_type not in control_types:
                continue
            if any(
                text == label if exact else label in text
                for label in labels
            ):
                return control
    raise LookupError(f"操作対象が見つかりません: {labels}")


def click_control(window, *labels: str):
    control = find_control(
        window,
        labels,
        {"Button", "RadioButton", "CheckBox", "Text", "TreeItem"},
    )
    try:
        control.click_input()
    except Exception:
        control.invoke()
    return control


def capture(window, output_dir: Path, name: str) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    try:
        window.capture_as_image().save(output_dir / name)
    except Exception:
        pass


def wait_for_enabled_control(window, labels: Iterable[str], timeout: int):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        try:
            control = find_control(window, labels, {"Button"})
            if control.is_enabled() and control.is_visible():
                return control
        except (LookupError, RuntimeError):
            pass
        time.sleep(1)
    raise TimeoutError(f"操作可能になるまで待機しました: {tuple(labels)}")


def attach_or_launch(Application, Desktop, executable: Path, timeout: int):
    try:
        app = Application(backend="uia").connect(path=str(executable))
        if app.windows():
            return app
    except Exception:
        pass

    subprocess.Popen([str(executable)])
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        windows = Desktop(backend="uia").windows()
        if any(window.is_visible() for window in windows):
            try:
                return Application(backend="uia").connect(path=str(executable))
            except Exception:
                pass
        time.sleep(1)
    raise TimeoutError(
        "GAIA Cloudの画面を開けません。通知領域に残っているGAIAを終了して再試行してください。"
    )


def fill_native_gaia_file_dialog(candidate: Path) -> bool:
    """Fill GAIA's classic Win32 file dialog without UIA control discovery."""
    import ctypes
    from ctypes import wintypes

    user32 = ctypes.windll.user32
    windows: list[int] = []
    enum_windows_proc = ctypes.WINFUNCTYPE(
        ctypes.c_bool, wintypes.HWND, wintypes.LPARAM
    )

    @enum_windows_proc
    def collect_window(hwnd, _):
        length = user32.GetWindowTextLengthW(hwnd)
        title_buffer = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, title_buffer, length + 1)
        class_buffer = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(hwnd, class_buffer, len(class_buffer))
        if (
            class_buffer.value == "#32770"
            and "設計書ファイル" in title_buffer.value
            and user32.IsWindowVisible(hwnd)
        ):
            windows.append(hwnd)
        return True

    user32.EnumWindows(collect_window, 0)
    if not windows:
        return False

    dialog = windows[0]
    filename_edits: list[int] = []
    open_buttons: list[int] = []
    enum_children_proc = ctypes.WINFUNCTYPE(
        ctypes.c_bool, wintypes.HWND, wintypes.LPARAM
    )

    @enum_children_proc
    def collect_child(hwnd, _):
        class_buffer = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(hwnd, class_buffer, len(class_buffer))
        control_id = user32.GetDlgCtrlID(hwnd)
        if control_id == 1148 and class_buffer.value == "Edit":
            filename_edits.append(hwnd)
        elif control_id == 1 and class_buffer.value == "Button":
            open_buttons.append(hwnd)
        return True

    user32.EnumChildWindows(dialog, collect_child, 0)
    if not filename_edits or not open_buttons:
        return False

    user32.SendMessageW(filename_edits[0], 0x000C, 0, str(candidate.resolve()))
    user32.SendMessageW(open_buttons[0], 0x00F5, 0, 0)
    return True


def handle_open_dialog(Desktop, candidate: Path, timeout: int) -> None:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if fill_native_gaia_file_dialog(candidate):
            return
        for backend in ("uia", "win32"):
            for window in Desktop(backend=backend).windows():
                try:
                    title = window.window_text()
                    if "開く" not in title and "設計書ファイル" not in title:
                        continue
                    edits = window.descendants(control_type="Edit")
                    if not edits:
                        continue
                    edits[-1].set_edit_text(str(candidate.resolve()))
                    click_control(window, "開く")
                    return
                except Exception:
                    continue
        time.sleep(0.5)
    raise TimeoutError("Excelファイル選択画面を操作できませんでした")


def set_project_name(window, project_name: str) -> None:
    labels = [
        control
        for control in window.descendants()
        if "工事名" in control.window_text()
    ]
    edits = window.descendants(control_type="Edit")
    if not edits:
        raise LookupError("工事名入力欄が見つかりません")
    if not labels:
        edits[0].set_edit_text(project_name)
        return

    label_rectangle = labels[0].rectangle()
    edit = min(
        edits,
        key=lambda candidate: abs(
            candidate.rectangle().mid_point().y - label_rectangle.mid_point().y
        ),
    )
    edit.set_edit_text(project_name)


def cache_snapshot(cache_root: Path) -> dict[Path, float]:
    if not cache_root.exists():
        return {}
    return {path: path.stat().st_mtime for path in cache_root.rglob("*.xlsx")}


def wait_for_cache_export(
    cache_root: Path,
    before: dict[Path, float],
    project_name: str,
    suffix: str,
    timeout: int,
) -> Path:
    deadline = time.monotonic() + timeout
    expected = f"{project_name}_{suffix}.xlsx"
    while time.monotonic() < deadline:
        candidates = [
            path
            for path in cache_root.rglob("*.xlsx")
            if (
                path.name == expected
                and path.stat().st_size > 0
                and path.stat().st_mtime > before.get(path, 0)
            )
        ]
        if candidates:
            newest = max(candidates, key=lambda path: path.stat().st_mtime)
            size = newest.stat().st_size
            time.sleep(1)
            if newest.stat().st_size == size:
                return newest
        time.sleep(1)
    raise TimeoutError(f"GAIAのExcel出力が見つかりません: {expected}")


def export_review_sheet(
    Desktop,
    overview,
    tab_name: str,
    suffix: str,
    destination: Path,
    config: TrialConfig,
) -> None:
    tabs = overview.descendants(control_type="Tab")
    if not tabs:
        raise LookupError("一覧表タブが見つかりません")
    tabs[0].select(tab_name)
    overview.set_focus()
    before = cache_snapshot(config.cache_root)
    click_control(overview, "Excel出力")
    print_dialog = wait_for_window(Desktop, "一覧表印刷設定", config.timeout_seconds)
    click_control(print_dialog, "Excel出力")
    generated = wait_for_cache_export(
        config.cache_root,
        before,
        config.project_name,
        suffix,
        config.timeout_seconds,
    )
    shutil.copy2(generated, destination)


def run_trial(config: TrialConfig) -> dict:
    validate_trial_config(config)
    Application, Desktop = import_pywinauto()
    attach_or_launch(Application, Desktop, config.gaia_exe, config.timeout_seconds)

    try:
        login = wait_for_window(Desktop, "ログイン", min(60, config.timeout_seconds))
        click_control(login, "ログイン")
    except (TimeoutError, LookupError):
        pass

    project_window = wait_for_window(
        Desktop, "設計書取込", config.timeout_seconds
    )
    capture(project_window, config.output_dir, "01_project_window.png")
    try:
        click_control(project_window, config.folder_name)
    except LookupError:
        raise LookupError(f"試験用フォルダーが見つかりません: {config.folder_name}")
    click_control(project_window, "設計書取込")

    try:
        # GAIA 11.1 normally opens the Win32 file picker immediately.
        handle_open_dialog(Desktop, config.candidate, min(15, config.timeout_seconds))
    except TimeoutError:
        # Some installations show an intermediate source-selection form.
        source_window = wait_for_window(
            Desktop, "参照", config.timeout_seconds
        )
        click_control(source_window, "参照")
        handle_open_dialog(Desktop, config.candidate, config.timeout_seconds)

    import_window = wait_for_window(Desktop, "金抜き", config.timeout_seconds)
    time.sleep(2)
    capture(import_window, config.output_dir, "02_import_recognition.png")
    texts = control_texts(import_window)
    for expected in ("珠洲市Excel", "珠洲市", "一般土木"):
        if not any(expected in text for text in texts):
            raise RuntimeError(f"GAIA取込認識の確認に失敗しました: {expected}")
    click_control(import_window, "金抜き")
    click_control(import_window, "取込")

    try:
        completed = wait_for_window(
            Desktop, "取込が完了", min(30, config.timeout_seconds)
        )
        click_control(completed, "OK", "閉じる")
    except (TimeoutError, LookupError):
        pass

    wizard = wait_for_window(Desktop, "次へ", config.timeout_seconds)
    page = 1
    while page <= 6:
        capture(wizard, config.output_dir, f"03_wizard_{page}.png")
        try:
            click_control(wizard, "次へ")
            page += 1
            time.sleep(1)
        except LookupError:
            break

    set_project_name(wizard, config.project_name)
    checkbox = find_control(
        wizard,
        ["設計書取込後に全自動積算する"],
        {"CheckBox"},
    )
    if checkbox.get_toggle_state() == 0:
        checkbox.click_input()
    capture(wizard, config.output_dir, "04_ready_to_start.png")
    click_control(wizard, "開始", "完了")

    result_window = wait_for_window(
        Desktop, "全自動積算", config.timeout_seconds
    )
    close_button = wait_for_enabled_control(
        result_window, ["閉じる", "OK"], config.timeout_seconds
    )
    capture(result_window, config.output_dir, "05_auto_result.png")
    close_button.click_input()

    result = {
        "project_name": config.project_name,
        "candidate": str(config.candidate.resolve()),
        "started_at": datetime.now().isoformat(timespec="seconds"),
        "technical_trial": True,
    }
    if not config.export_review:
        return result

    estimate_window = wait_for_window(Desktop, "一覧表", config.timeout_seconds)
    click_control(estimate_window, "一覧表")
    overview = wait_for_window(Desktop, "Excel出力", config.timeout_seconds)
    unit_path = config.output_dir / "gaia_confirm_unit_prices.xlsx"
    work_path = config.output_dir / "gaia_confirm_work_types.xlsx"
    export_review_sheet(
        Desktop, overview, "確認単価", "確認単価", unit_path, config
    )
    export_review_sheet(
        Desktop, overview, "確認工種", "確認工種", work_path, config
    )
    review_rows = parse_gaia_review_workbook(unit_path, "unit")
    review_rows.extend(parse_gaia_review_workbook(work_path, "work"))
    reconciliation_path = config.output_dir / "gaia_reconciliation.csv"
    reconciled = reconcile_rows(load_source_review(config.source_review), review_rows)
    result["reconciliation_counts"] = write_reconciliation(
        reconciliation_path, reconciled
    )
    result["reconciliation"] = str(reconciliation_path.resolve())
    return result


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "GAIA Cloud 11.1で金抜Excelの技術試験取込、全自動積算、"
            "確認一覧出力を実行します。実案件への使用はまだ禁止しています。"
        )
    )
    parser.add_argument("--candidate", required=True, type=Path)
    parser.add_argument("--source-review", required=True, type=Path)
    parser.add_argument("--project-name", required=True)
    parser.add_argument("--folder-name", default="テスト")
    parser.add_argument("--output-dir", type=Path, default=Path("outputs/ui_trial"))
    parser.add_argument("--gaia-exe", type=Path, default=DEFAULT_GAIA_EXE)
    parser.add_argument("--cache-root", type=Path, default=DEFAULT_CACHE_ROOT)
    parser.add_argument("--timeout-seconds", type=int, default=1200)
    parser.add_argument("--confirm-start", action="store_true")
    parser.add_argument("--export-review", action="store_true")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    config = TrialConfig(
        candidate=args.candidate.resolve(),
        source_review=args.source_review.resolve(),
        project_name=args.project_name,
        folder_name=args.folder_name,
        output_dir=args.output_dir.resolve(),
        gaia_exe=args.gaia_exe.resolve(),
        cache_root=args.cache_root.resolve(),
        timeout_seconds=args.timeout_seconds,
        confirm_start=args.confirm_start or args.dry_run,
        export_review=args.export_review,
    )
    validate_trial_config(config)
    if args.dry_run:
        print(json.dumps(asdict(config), ensure_ascii=False, indent=2, default=str))
        return 0
    result = run_trial(config)
    config.output_dir.mkdir(parents=True, exist_ok=True)
    (config.output_dir / "trial_result.json").write_text(
        json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
