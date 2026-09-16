#!/usr/bin/env python3
"""せいぎデザイン dist ZIP を作成する。"""

from __future__ import annotations

import shutil
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DIST = ROOT / "dist"
PPT = ROOT / "ppt" / "templates"
XLS = ROOT / "excel" / "templates"


def zip_dir(src_files: list[Path], dest: Path, arc_prefix: str):
    dest.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(dest, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for path in src_files:
            zf.write(path, arcname=f"{arc_prefix}/{path.name}" if path.parent.name != "layouts" else f"{arc_prefix}/layouts/{path.name}")


def main():
    if DIST.exists():
        shutil.rmtree(DIST)
    DIST.mkdir(parents=True)

    ppt_files = sorted(PPT.glob("*.pptx")) + sorted((PPT / "layouts").glob("*.pptx"))
    xls_files = sorted(XLS.glob("*.xlsx"))
    starter = [
        PPT / "seigi-starter.pptx",
        PPT / "seigi-basic.pptx",
        PPT / "seigi-jig-dr.pptx",
        XLS / "seigi-minutes.xlsx",
        XLS / "seigi-checklist.xlsx",
        XLS / "seigi-jig-dr.xlsx",
        XLS / "seigi-department-text.xlsx",
    ]
    starter = [p for p in starter if p.exists()]

    zip_dir(ppt_files, DIST / "seigi-powerpoint-all.zip", "seigi-powerpoint")
    zip_dir(xls_files, DIST / "seigi-excel-all.zip", "seigi-excel")

    with zipfile.ZipFile(DIST / "seigi-starter-kit.zip", "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for path in starter:
            folder = "powerpoint" if path.suffix == ".pptx" else "excel"
            zf.write(path, arcname=f"seigi-starter/{folder}/{path.name}")
        readme = DIST / "_starter_readme.txt"
        readme.write_text(
            "せいぎデザイン スターターキット\n"
            "1. powerpoint/ の seigi-starter.pptx か seigi-basic.pptx を開く\n"
            "2. ジグDRは seigi-jig-dr.pptx / seigi-jig-dr.xlsx\n"
            "3. 色は変えなくてOK。文言だけ差し替える\n"
            "4. 最新版は Box のせいぎデザインフォルダを参照\n",
            encoding="utf-8",
        )
        zf.write(readme, arcname="seigi-starter/README.txt")
        readme.unlink(missing_ok=True)

    print("dist:")
    for p in sorted(DIST.glob("*.zip")):
        print(f"  {p.name} ({p.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
