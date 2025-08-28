# DOCX → NCJ (Normalized Content JSON) minimal toolkit

## 0) Requirements
- Pandoc is already installed (you said via Homebrew).
- No Python dependencies needed (pure stdlib).

## 1) Convert .docx to Pandoc AST + extract images
```bash
pandoc "your.docx" -t json --extract-media=assets > ast.json
```

## 2) Transform AST → NCJ
```bash
python to_ncj.py --source "your.docx" --style-map style.yml < ast.json > content.json
```

The tool:
- Preserves order, keeps **semantics** only (h1/h2/paragraph/figure/table).
- Detects **images** and absorbs nearby **captions** (e.g., lines starting with "图1:"/"Figure 1:"/"Chart 2:" etc.) and **credits** (lines starting with "来源:"/"Source:").
- Writes `assets[]` with SHA-256 for images (from `--extract-media=assets`).

## 3) Feed `content.json` to your Figma automation
Your Figma plugin/agent renders `blocks[]` sequentially into the Content Group (vertical Auto Layout, no font size changes). If overflow, extend background height.

## Notes
- You can tailor regex in `to_ncj.py` for caption/credit detection.
- For fancier mapping (e.g., custom Word styles), keep updating `style.yml` and extend the classifier in the script if you need strict style-based routing.
- Tables are emitted as a stub; if your reports rely on tables, add a handler for Pandoc's Table node (version-specific).
