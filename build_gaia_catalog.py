from __future__ import annotations

import argparse
from pathlib import Path

from gaia_catalog import write_gaia_catalog


def main() -> int:
    parser = argparse.ArgumentParser(
        description="GAIAで読込済みの設計書から取込名称・階層・コードを抽出します。"
    )
    parser.add_argument("sources", nargs="+", type=Path)
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("assets") / "gaia_import_catalog.json",
    )
    args = parser.parse_args()
    entries = write_gaia_catalog(args.sources, args.output)
    print(f"{len(entries)} entries -> {args.output.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
