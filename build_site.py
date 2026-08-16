# -*- coding: utf-8 -*-
"""Build static site/ from Excel for Netlify deploy."""
import json
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SITE = ROOT / "site"

# Prefer pruned practical list; backup is the full unpruned source.
CANDIDATES = [
    Path(r"e:\Users\i\Desktop\单词\面试常用单词_Java_Python_FastAPI_LangChain.xlsx"),
    ROOT / "面试常用单词_Java_Python_FastAPI_LangChain.xlsx",
    ROOT / "vocab.xlsx",
    ROOT / "面试常用单词_Java_Python_FastAPI_LangChain.backup.xlsx",
    Path(r"e:\Users\i\Desktop\专高六\面试常用单词_Java_Python_FastAPI_LangChain.xlsx"),
]


def main() -> None:
    import openpyxl

    subprocess.check_call([sys.executable, str(ROOT / "write_index.py")])

    src = next((p for p in CANDIDATES if p.exists()), None)
    if src is None:
        raise SystemExit("No Excel source found")
    print("using", src)

    wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
    ws = wb["全部汇总"]
    rows_raw = list(ws.iter_rows(values_only=True))
    header = ["" if c is None else str(c) for c in rows_raw[0]][:10]
    rows = []
    for r in rows_raw[1:]:
        if not r or r[1] is None or str(r[1]).strip() == "":
            continue
        row = []
        for i in range(10):
            v = r[i] if i < len(r) else None
            row.append("" if v is None else str(v).strip())
        rows.append(row)
    wb.close()
    print("rows", len(rows))

    SITE.mkdir(exist_ok=True)
    shutil.copy2(src, SITE / "interview-vocab.xlsx")
    (SITE / "vocab-data.json").write_text(
        json.dumps({"header": header, "rows": rows}, ensure_ascii=False, separators=(",", ":")),
        encoding="utf-8",
    )

    html_path = SITE / "index.html"
    html = html_path.read_text(encoding="utf-8").replace("__COUNT__", str(len(rows)))
    html_path.write_text(html, encoding="utf-8")

    # Mirror to docs/ for GitHub Pages
    docs = ROOT / "docs"
    docs.mkdir(exist_ok=True)
    for name in ("index.html", "vocab-data.json", "interview-vocab.xlsx"):
        shutil.copy2(SITE / name, docs / name)
    (docs / ".nojekyll").write_text("", encoding="utf-8")
    print("site ready:", SITE)
    print("docs ready:", docs)


if __name__ == "__main__":
    main()
