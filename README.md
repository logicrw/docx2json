# DOCX to JSON Converter

## Figure Grouping Algorithm

The converter implements a two-phase figure grouping algorithm to accurately detect related images:

### Phase 1: Same-Paragraph Grouping (Row Layout)
- Multiple images within the same paragraph are automatically grouped with `layout='row'`
- Typical in side-by-side image scenarios

### Phase 2: Adjacent-Paragraph Grouping (Column Layout)  
- Images in consecutive paragraphs with minimal text gaps are grouped with `layout='column'`
- Only groups if gap ≤ `max_gap_paras` and no substantial text (>max_title_len chars) between images
- Exception: If combined width ≤ `page_width_ratio` * page_width, uses `layout='row'` 

### Title and Credit Attribution
- **Title**: Assigned to group's first figure from nearest short text (≤max_title_len chars) above/below
- **Credit**: Assigned to group's last figure from nearest "来源:/Source:" pattern above/below

### CLI Parameters
```bash
python to_ncj.py "input.docx" [options]
  --max_title_len 45        # Max chars for title detection (default: 45)
  --max_gap_paras 1         # Max paragraph gap for grouping (default: 1) 
  --page_width_ratio 0.95   # Width ratio for row layout detection (default: 0.95)
  --debug                   # Include grouping reasoning in output
```