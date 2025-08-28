#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Pandoc JSON AST (from `pandoc -t json --extract-media=assets`) ->
Normalized Content JSON (NCJ) tailored for downstream layout engines.

Usage:
  pandoc input.docx -t json --extract-media=assets \
  | python to_ncj.py --source "input.docx" --style-map style.yml > content.json
"""
import sys, json, re, os, hashlib
from typing import List, Dict, Any, Tuple, Optional

# --------- tiny YAML-ish parser for a simple style map ----------
def load_style_map(path: Optional[str]) -> Dict[str,str]:
    if not path: return {}
    if not os.path.exists(path): return {}
    mapping = {}
    in_styles = False
    with open(path, 'r', encoding='utf-8') as f:
        for raw in f:
            line = raw.strip()
            if not line or line.startswith('#'): continue
            if line.startswith('styles:'):
                in_styles = True
                continue
            if in_styles:
                m = re.match(r'^("?)(.+?)\1\s*:\s*([A-Za-z0-9_\-]+)$', line)
                if m:
                    key = m.group(2)
                    val = m.group(3)
                    mapping[key] = val
    return mapping

# --------- helpers ----------
def stringify_inlines(inlines: List[Any]) -> str:
    out = []
    def walk(xs):
        for x in xs or []:
            # Ensure x is a dict-like inline; skip if malformed
            if not isinstance(x, dict):
                continue
            t = x.get('t'); c = x.get('c')
            if t == 'Str':
                out.append(c)
            elif t == 'Space':
                out.append(' ')
            elif t in ('SoftBreak','LineBreak'):
                out.append('\n')
            elif t == 'Code':
                # Code inline: [attr, text]
                out.append(c[1] if isinstance(c, list) and len(c) > 1 else '')
            elif t in ('Emph','Strong','Underline','Strikeout','Superscript','Subscript','SmallCaps'):
                # These wrap a list of inlines directly
                walk(c if isinstance(c, list) else [])
            elif t == 'Span':
                # Span: [attr, inlines]
                walk(c[1] if isinstance(c, list) and len(c) > 1 else [])
            elif t == 'Link':
                # Link: [attr, inlines, target]
                walk(c[1] if isinstance(c, list) and len(c) > 1 else [])
            elif t == 'Image':
                # Image: [attr, alt-inlines, target]
                walk(c[1] if isinstance(c, list) and len(c) > 1 else [])
            elif t == 'Quoted':
                # Quoted: [quote-type, inlines]
                walk(c[1] if isinstance(c, list) and len(c) > 1 else [])
            elif t == 'Cite':
                # Cite: [citations, inlines]
                walk(c[1] if isinstance(c, list) and len(c) > 1 else [])
            # else: ignore other inline types
    walk(inlines)
    return ''.join(out).strip()

def attr_tuple(attr: Any) -> Tuple[str, List[str], Dict[str,str]]:
    if not isinstance(attr, list) or len(attr) != 3:
        return ("", [], {})
    ident = attr[0] or ""
    classes = attr[1] or []
    kv = {k:v for k,v in (attr[2] or [])}
    return ident, classes, kv

def sha256_of_file(path: str) -> str:
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(65536), b''):
            h.update(chunk)
    return h.hexdigest()

# --------- normalization helpers ----------
def normalize_credit(text: str) -> str:
    if not text: return ""
    # Remove 来源:/Source: (both half-width and full-width colons)
    text = re.sub(r'^\s*(来源|Source)\s*[：:]\s*', '', text, flags=re.I)
    # Remove leading/trailing spaces and ending punctuation
    text = text.strip().rstrip('.,;，。；')
    return text

# caption/title heuristics and doc title/date detection
DOC_TITLE_RE = re.compile(r'^\s*(\d{6})\s*-\s*(.+)')
def is_short_sentence(s: str) -> bool:
    return bool(s) and len(s.strip()) <= 45
def parse_date_from_yyMMdd(yyMMdd: str) -> Optional[str]:
    try:
        yy = int(yyMMdd[0:2]); mm = int(yyMMdd[2:4]); dd = int(yyMMdd[4:6])
        yyyy = 2000 + yy
        return f"{yyyy:04d}-{mm:02d}-{dd:02d}"
    except Exception:
        return None

# --------- main conversion ----------
def convert(pandoc_ast: Dict[str,Any], source_file: str, style_map_path: Optional[str]) -> Dict[str,Any]:
    blocks = pandoc_ast.get('blocks', [])
    meta = pandoc_ast.get('meta', {})
    style_map = load_style_map(style_map_path)

    CAPTION_HINT = re.compile(r'^(图|Figure|Fig\.?|图表|Chart|Graph)\s*\d+[:：．. ]|^【图|^图：', re.I)
    CREDIT_HINT  = re.compile(r'^(来源|Source)[:：]', re.I)

    assets = {}
    warnings = []
    out_blocks = []
    doc_title_text: Optional[str] = None
    doc_date_text: Optional[str] = None

    # --------- helpers for inline images ----------
    def extract_inline_images_from_inlines(inlines):
        imgs = []
        for x in inlines:
            if x.get('t') == 'Image':
                attr, alt, target = x['c']
                src, title = target
                ident, classes, kv = attr_tuple(attr)
                imgs.append({
                    "src": src,
                    "title": stringify_inlines(alt) or None,
                    "ident": ident,
                    "classes": classes,
                    "kv": kv,
                })
        return imgs

    # --------- re-stream blocks ----------
    simple_stream = []
    for b in blocks:
        t = b.get('t'); c = b.get('c')
        if t == 'Header':
            level, attr, inlines = c
            txt = stringify_inlines(inlines)
            # Detect doc title/date from very first textual block
            if doc_title_text is None and not simple_stream:
                m = DOC_TITLE_RE.match(txt or '')
                if m:
                    doc_title_text = txt.strip()
                    doc_date_text = parse_date_from_yyMMdd(m.group(1))
                    continue
            simple_stream.append(('header', {'level':level, 'text':txt}))
        elif t in ('Para','Plain'):
            imgs = extract_inline_images_from_inlines(c)
            if imgs:
                simple_stream.append(('imagepara', {'images':imgs}))
            else:
                txt = stringify_inlines(c)
                if doc_title_text is None and not simple_stream:
                    m = DOC_TITLE_RE.match(txt or '')
                    if m:
                        doc_title_text = txt.strip()
                        doc_date_text = parse_date_from_yyMMdd(m.group(1))
                        continue
                simple_stream.append(('para', {'text':txt}))
        elif t == 'Table':
            # skip tables in output
            continue
        elif t in ('BlockQuote','Div'):
            inner = c[1] if t=='Div' else c
            for ib in inner:
                if ib.get('t') in ('Para','Plain'):
                    imgs = extract_inline_images_from_inlines(ib.get('c'))
                    if imgs:
                        simple_stream.append(('imagepara', {'images':imgs}))
                    else:
                        simple_stream.append(('para', {'text': stringify_inlines(ib.get('c'))}))
                elif ib.get('t') == 'Header':
                    level, attr, inlines = ib.get('c')
                    simple_stream.append(('header', {'level':level, 'text': stringify_inlines(inlines)}))
        # ignore others

    # --------- consume stream ----------
    # Note: keep title/date in doc meta only; no doc_title block in content
    j = 0
    group_counter = 1
    while j < len(simple_stream):
        kind, data = simple_stream[j]
        if kind == 'header':
            # Treat headers as regular paragraphs under the new constraint
            text = (data.get('text') or '').strip()
            if text:
                out_blocks.append({"type":"paragraph","text":text})
        elif kind == 'para':
            text = (data.get('text') or '').strip()
            if text:
                out_blocks.append({"type":"paragraph","text":text})
        elif kind == 'imagepara':
            images = data['images']
            group_len = len(images)
            group_id = None
            # 1) Title absorption: prefer previous short paragraph; fallback to next short paragraph
            title_text: Optional[str] = None
            title_idx: Optional[int] = None
            # prev
            if j-1 >= 0 and simple_stream[j-1][0] == 'para':
                prev_txt = (simple_stream[j-1][1].get('text') or '').strip()
                if prev_txt and is_short_sentence(prev_txt) and not CREDIT_HINT.search(prev_txt):
                    title_text = prev_txt
                    title_idx = j-1
                    simple_stream[title_idx] = ('other', {})  # consume
                    # Remove previously emitted duplicate paragraph if it was just output
                    if out_blocks and out_blocks[-1].get('type') == 'paragraph' and out_blocks[-1].get('text') == prev_txt:
                        out_blocks.pop()
            # next fallback
            if title_text is None and j+1 < len(simple_stream) and simple_stream[j+1][0] == 'para':
                next_txt_for_title = (simple_stream[j+1][1].get('text') or '').strip()
                if next_txt_for_title and is_short_sentence(next_txt_for_title) and not CREDIT_HINT.search(next_txt_for_title):
                    title_text = next_txt_for_title
                    title_idx = j+1
                    simple_stream[title_idx] = ('other', {})  # consume
            # 2) Credit absorption: search within 2 paragraphs forward, else within 2 paragraphs backward (skip title_idx)
            credit_text: Optional[str] = None
            # forward search up to 2 para blocks
            seen_para = 0
            k = j + 1
            while k < len(simple_stream) and seen_para < 2 and credit_text is None:
                kkind, kdata = simple_stream[k]
                if kkind == 'para':
                    if k != (title_idx or -1):
                        ktxt = (kdata.get('text') or '').strip()
                        if ktxt and CREDIT_HINT.search(ktxt):
                            m = re.split(r'[:：]\s*', ktxt, maxsplit=1)
                            absorbed_credit = m[1] if len(m) > 1 else ktxt
                            credit_text = normalize_credit(absorbed_credit)
                            simple_stream[k] = ('other', {})  # consume
                            break
                    seen_para += 1
                k += 1
            # backward search up to 2 para blocks (skip title_idx) if not found
            if credit_text is None:
                seen_para = 0
                k = j - 1
                while k >= 0 and seen_para < 2 and credit_text is None:
                    if simple_stream[k][0] == 'para' and k != (title_idx or -1):
                        ktxt = (simple_stream[k][1].get('text') or '').strip()
                        if ktxt and CREDIT_HINT.search(ktxt):
                            m = re.split(r'[:：]\s*', ktxt, maxsplit=1)
                            absorbed_credit = m[1] if len(m) > 1 else ktxt
                            credit_text = normalize_credit(absorbed_credit)
                            simple_stream[k] = ('other', {})  # consume
                            # Remove previously emitted duplicate paragraph if present in out_blocks
                            for idx in range(len(out_blocks) - 1, -1, -1):
                                ob = out_blocks[idx]
                                if ob.get('type') == 'paragraph' and ob.get('text') == ktxt:
                                    out_blocks.pop(idx)
                                    break
                            break
                        seen_para += 1
                    if simple_stream[k][0] == 'para':
                        # count only paragraph kinds
                        pass
                    k -= 1
            # 3) Generate group metadata for ALL images (single or multiple)
            group_id = f"grp_{group_counter:04d}"
            group_counter += 1
            
            for idx_img, img in enumerate(images):
                src = img['src']
                asset_id = None
                if src and os.path.exists(src):
                    sha = sha256_of_file(src)
                    asset_id = f"img_{sha[:12]}"
                    if asset_id not in assets:
                        assets[asset_id] = {"asset_id": asset_id, "filename": src, "sha256": sha}
                figure = {
                    "type":"figure",
                    "image":{"asset_id": asset_id or src or "", "alt": img.get('title') or None},
                    "title": None,
                    "credit": None
                }
                # Always assign group metadata (for consistent plugin logic)
                figure["group_id"] = group_id
                figure["group_seq"] = idx_img + 1
                figure["group_len"] = group_len
                
                # Assign title to group head and credit to group tail
                if title_text and idx_img == 0:
                    figure["title"] = title_text
                if credit_text and idx_img == group_len - 1:
                    figure["credit"] = credit_text
                out_blocks.append(figure)
        j += 1

    ncj = {
        "doc": {
            "title": (doc_title_text or (meta.get('title', {}).get('c') if isinstance(meta.get('title'), dict) else None)),
            "date": doc_date_text,
            "locale": "zh-CN",
            "version": "v1",
            "source_file": source_file
        },
        "style_map": load_style_map(style_map_path),
        "blocks": out_blocks,
        "assets": list(assets.values()),
        "report": {"warnings": warnings}
    }
    return ncj

def main():
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument('--source', default='unknown.docx')
    ap.add_argument('--style-map', default=None)
    args = ap.parse_args()
    ast = json.loads(sys.stdin.read())
    ncj = convert(ast, args.source, args.style_map)
    json.dump(ncj, sys.stdout, ensure_ascii=False, indent=2)

if __name__ == '__main__':
    main()
