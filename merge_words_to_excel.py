# -*- coding: utf-8 -*-
"""合并社区投稿到 Excel，并重建 docs/vocab-data.json。"""
from __future__ import annotations

import json
import re
import shutil
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent
DESKTOP_XLSX = Path(r"e:\Users\i\Desktop\单词\面试常用单词_Java_Python_FastAPI_LangChain.xlsx")
REPO_XLSX = ROOT / "面试常用单词_Java_Python_FastAPI_LangChain.xlsx"
COMMUNITY_JSON = ROOT / "docs" / "community-words.json"
PENDING_JSON = ROOT / "docs" / "pending-words.json"
DOCS = ROOT / "docs"
ISSUE_REPO = "Tewelve-creator/twelve"

HEADER = [
    "类别", "单词", "用法/解释", "中文谐音", "音标",
    "发音", "技术栈", "重要程度", "对比词", "一句话回答",
]


def pick(body: str, key: str) -> str:
    m = re.search(rf"[-*]\s*\*\*{re.escape(key)}\*\*:\s*(.+)", body or "")
    return (m.group(1).strip().replace("—", "") if m else "")


def load_community_words() -> list[dict]:
    words: list[dict] = []
    for path in (COMMUNITY_JSON, PENDING_JSON):
        if not path.exists():
            continue
        raw = json.loads(path.read_text(encoding="utf-8"))
        if isinstance(raw, list):
            words.extend(raw)
        elif isinstance(raw, dict):
            words.extend(raw.get("words") or [])

    try:
        url = f"https://api.github.com/repos/{ISSUE_REPO}/issues?labels=word-submission&state=open&per_page=100"
        req = urllib.request.Request(
            url,
            headers={"Accept": "application/vnd.github+json", "User-Agent": "twelve-merge"},
        )
        with urllib.request.urlopen(req, timeout=20) as resp:
            issues = json.loads(resp.read().decode("utf-8"))
        for issue in issues or []:
            if issue.get("pull_request"):
                continue
            body = issue.get("body") or ""
            word = pick(body, "单词") or str(issue.get("title") or "").replace("[词条投稿]", "").strip()
            if not word:
                continue
            trouble = [t for t in pick(body, "卡在哪").replace("，", "、").split("、") if t and t != "（未选）"]
            words.append(
                {
                    "word": word,
                    "explain": pick(body, "用法/解释"),
                    "homo": pick(body, "中文谐音"),
                    "ipa": pick(body, "音标"),
                    "pron": pick(body, "发音"),
                    "tech": pick(body, "技术栈") or "用户投稿",
                    "level": pick(body, "重要程度") or "了解",
                    "compare": pick(body, "对比词") or "—",
                    "oneliner": pick(body, "一句话回答"),
                    "trouble": trouble,
                    "cat": pick(body, "分类") or "",
                }
            )
    except Exception as e:
        print("issues fetch skipped:", e)

    dedup = {}
    for w in words:
        if not isinstance(w, dict):
            continue
        key = str(w.get("word") or "").strip().lower()
        if key:
            dedup[key] = w
    return list(dedup.values())


def word_to_row(w: dict) -> list[str]:
    trouble = w.get("trouble") or []
    cat = str(w.get("cat") or "").strip()
    if not cat:
        if isinstance(trouble, list) and trouble:
            cat = "社区投稿·" + "/".join(str(t) for t in trouble if t)
        else:
            cat = "社区投稿"
    return [
        cat,
        str(w.get("word") or "").strip(),
        str(w.get("explain") or "").strip(),
        str(w.get("homo") or "").strip(),
        str(w.get("ipa") or "").strip(),
        str(w.get("pron") or "").strip(),
        str(w.get("tech") or "用户投稿").strip() or "用户投稿",
        str(w.get("level") or "了解").strip() or "了解",
        str(w.get("compare") or w.get("note") or "—").strip() or "—",
        str(w.get("oneliner") or "").strip(),
    ]


def merge_into_workbook(xlsx_path: Path, community: list[dict]) -> tuple[int, int]:
    import openpyxl

    wb = openpyxl.load_workbook(xlsx_path)
    if "全部汇总" not in wb.sheetnames:
        raise SystemExit(f"missing sheet 全部汇总 in {xlsx_path}")
    ws = wb["全部汇总"]

    existing = set()
    for row in ws.iter_rows(min_row=2, max_col=2, values_only=True):
        if row and len(row) > 1 and row[1] is not None and str(row[1]).strip():
            existing.add(str(row[1]).strip().lower())

    added = 0
    for w in community:
        key = str(w.get("word") or "").strip().lower()
        if not key or key in existing:
            continue
        ws.append(word_to_row(w))
        existing.add(key)
        added += 1

    total = sum(
        1
        for row in ws.iter_rows(min_row=2, max_col=2, values_only=True)
        if row and len(row) > 1 and row[1] is not None and str(row[1]).strip()
    )
    wb.save(xlsx_path)
    wb.close()
    return added, total


def export_vocab_json(xlsx_path: Path) -> int:
    import openpyxl

    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb["全部汇总"]
    rows_raw = list(ws.iter_rows(values_only=True))
    wb.close()
    header = ["" if c is None else str(c) for c in (rows_raw[0] if rows_raw else HEADER)][:10]
    while len(header) < 10:
        header.append(HEADER[len(header)])
    rows = []
    for r in rows_raw[1:]:
        if not r or r[1] is None or str(r[1]).strip() == "":
            continue
        rows.append([("" if (i >= len(r) or r[i] is None) else str(r[i]).strip()) for i in range(10)])

    DOCS.mkdir(exist_ok=True)
    (DOCS / "vocab-data.json").write_text(
        json.dumps({"header": header, "rows": rows}, ensure_ascii=False, separators=(",", ":")),
        encoding="utf-8",
    )
    shutil.copy2(xlsx_path, DOCS / "interview-vocab.xlsx")
    return len(rows)


def main() -> None:
    community = load_community_words()
    print("community words:", len(community))
    candidates = [p for p in (DESKTOP_XLSX, REPO_XLSX) if p.exists()]
    if not candidates:
        raise SystemExit("No Excel file found")

    primary = None
    last_err = None
    added = total = 0
    for path in candidates:
        try:
            added, total = merge_into_workbook(path, community)
            primary = path
            print(f"merged into {path}: +{added}, total {total}")
            break
        except PermissionError as e:
            last_err = e
            print("locked, skip", path)
        except Exception as e:
            last_err = e
            print("failed", path, e)
    if primary is None:
        raise SystemExit(f"cannot write Excel: {last_err}")

    for other in candidates:
        if other.resolve() == primary.resolve():
            continue
        try:
            shutil.copy2(primary, other)
            print("copied to", other)
        except PermissionError:
            print("copy skipped (locked):", other)

    if primary.resolve() != REPO_XLSX.resolve():
        try:
            shutil.copy2(primary, REPO_XLSX)
        except PermissionError:
            pass

    n = export_vocab_json(REPO_XLSX if REPO_XLSX.exists() else primary)
    print("docs/vocab-data.json rows:", n)


if __name__ == "__main__":
    main()
