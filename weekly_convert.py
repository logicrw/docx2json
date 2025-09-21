#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
weekly_convert.py

批次將同一週的繁體與英文 DOCX 轉換為多語 JSON，並共用單一資產目錄。

用法示例：
    python weekly_convert.py \
        --traditional-doc "250915 - 單向上行.docx" \
        --english-doc "250915 - Up-Only.docx"

選項：
    --output-dir            產出 JSON 的目錄（預設當前目錄）
    --assets-dir            圖像資產根目錄（預設 assets）
    --keep-duplicate-assets 保留英文 DOCX 轉換時生成的獨立資產目錄（預設刪除以節省空間）
"""

import argparse
import json
import os
import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, Iterable, Tuple

import to_ncj


def slug_from_doc_path(doc_path: str) -> str:
    """仿照 to_ncj 內部邏輯，從 DOCX 路徑推導資產子目錄 slug。"""
    doc_base = os.path.splitext(os.path.basename(doc_path))[0]
    tmp = doc_base.lower()
    tmp = re.sub(r"[^0-9a-z]+", "_", tmp)
    tmp = re.sub(r"_+", "_", tmp).strip("_")
    return tmp or "doc"


def derive_date_token(doc_meta: Dict[str, Any], fallback: str) -> str:
    """優先使用 doc.date (YYYY-MM-DD)，否則回退到檔名內的 6 位數字。"""
    date_str = doc_meta.get("date")
    if date_str:
        try:
            dt = datetime.strptime(date_str, "%Y-%m-%d")
            return dt.strftime("%y%m%d")
        except ValueError:
            pass
    match = re.search(r"(\d{6})", fallback)
    if match:
        return match.group(1)
    return datetime.now().strftime("%y%m%d")


def build_output_filename(date_token: str, title: str, locale: str) -> str:
    """根據日期、標題與 locale 生成穩定的輸出檔名。"""
    if any(ord(ch) > 127 for ch in title):
        title_part = title.replace(" ", "")
    else:
        title_part = title.lower()
        title_part = re.sub(r"[^0-9a-z]+", "-", title_part)
        title_part = re.sub(r"-+", "-", title_part).strip("-")
        if not title_part:
            title_part = "document"
    return f"{date_token}-{title_part}_{locale}.json"


def convert_doc(doc_path: str, *, assets_dir: str, locale: str,
                traditional_to_simplified: bool) -> Dict[str, Any]:
    """執行單次 DOCX -> NCJ 轉換，並覆寫 doc.locale。"""
    config = to_ncj.Config()
    config.assets_dir = assets_dir
    config.traditional_to_simplified = traditional_to_simplified
    config.use_explicit_markers = True
    ncj = to_ncj.convert_docx_to_ncj(doc_path, config)
    ncj["doc"]["locale"] = locale
    return ncj


def share_assets(ncj: Dict[str, Any], target_slug: str) -> None:
    """將 NCJ 中的資產路徑指向指定 slug 目錄。"""
    for asset in ncj.get("assets", []):
        basename = Path(asset["filename"]).name
        asset["filename"] = f"{target_slug}/{basename}"


def write_outputs(pairs: Iterable[Tuple[str, Dict[str, Any]]], output_dir: Path) -> None:
    """寫入所有 JSON 輸出並打印摘要。"""
    output_dir.mkdir(parents=True, exist_ok=True)
    for filename, payload in pairs:
        target = output_dir / filename
        target.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"Wrote {target} ({len(payload.get('blocks', []))} blocks, "
              f"{len(payload.get('assets', []))} assets)")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Convert weekly Traditional + English DOCX into three locale-specific JSON files.")
    parser.add_argument("--traditional-doc", required=True, help="繁體原始 DOCX 路徑")
    parser.add_argument("--english-doc", required=True, help="英文原始 DOCX 路徑")
    parser.add_argument("--output-dir", default=".", help="JSON 輸出目錄（預設當前目錄）")
    parser.add_argument("--assets-dir", default="assets", help="資產存放根目錄（預設 assets）")
    parser.add_argument("--keep-duplicate-assets", action="store_true",
                        help="保留英文 DOCX 產生的獨立資產目錄（預設刪除以共用繁體目錄）")
    args = parser.parse_args()

    output_dir = Path(args.output_dir)
    assets_dir = Path(args.assets_dir)

    traditional_slug = slug_from_doc_path(args.traditional_doc)
    english_slug = slug_from_doc_path(args.english_doc)

    # 1) 繁體 JSON
    hk_ncj = convert_doc(
        args.traditional_doc,
        assets_dir=str(assets_dir),
        locale="zh-HK",
        traditional_to_simplified=False,
    )

    # 2) 簡體 JSON
    cn_ncj = convert_doc(
        args.traditional_doc,
        assets_dir=str(assets_dir),
        locale="zh-CN",
        traditional_to_simplified=True,
    )

    # 3) 英文 JSON，完成後共用資產目錄
    en_ncj = convert_doc(
        args.english_doc,
        assets_dir=str(assets_dir),
        locale="en-US",
        traditional_to_simplified=False,
    )
    share_assets(en_ncj, traditional_slug)

    if not args.keep_duplicate_assets:
        duplicate_dir = assets_dir / english_slug
        if duplicate_dir.exists():
            shutil.rmtree(duplicate_dir)

    date_token = derive_date_token(hk_ncj["doc"], hk_ncj["doc"].get("source_file", ""))

    outputs = [
        (build_output_filename(date_token, hk_ncj["doc"]["title"], "zh-HK"), hk_ncj),
        (build_output_filename(date_token, cn_ncj["doc"]["title"], "zh-CN"), cn_ncj),
        (build_output_filename(date_token, en_ncj["doc"]["title"], "en-US"), en_ncj),
    ]

    write_outputs(outputs, output_dir)
    shared_assets_path = assets_dir / traditional_slug
    print(f"Assets stored under {shared_assets_path} (shared by all locales).")


if __name__ == "__main__":
    main()
