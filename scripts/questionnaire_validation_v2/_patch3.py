"""
Fix: replace compare_hint_changes and validate_crop_harvest to return DataFrames.
"""
import json, sys

NB = "c:/Users/edoar/WORK/FAO/repo/questionnaire_validation_revision/scripts/questionnaire_validation_v2/questionnaire_validation_kobo.ipynb"

with open(NB, encoding="utf-8") as f:
    nb = json.load(f)

cells = {c["id"]: c for c in nb["cells"] if c.get("id")}

def get_src(cid):
    s = cells[cid]["source"]
    return "".join(s) if isinstance(s, list) else s

def set_src(cid, s):
    cells[cid]["source"] = s

# ── Fix compare_hint_changes in abeb4da3 ──────────────────────────────────────
src = get_src("abeb4da3")

OLD_HINT = '''def compare_hint_changes(current: pl.DataFrame, reference: pl.DataFrame) -> list[dict]:
    """Flag questions where the hint text differs from the template (medium severity)."""
    ref_hints = reference.filter(pl.col("hint") != "").select(["Q Name", "hint"])
    if ref_hints.is_empty():
        return []
    joined = current.join(ref_hints, on="Q Name", how="inner", suffix="_ref")
    changed = joined.filter(pl.col("hint") != pl.col("hint_ref"))
    issues = []
    for row in changed.iter_rows(named=True):
        issues.append({
            "issue_type": "hint_changed",
            "Q Name":     row["Q Name"],
            "severity":   "medium",
            "detail":     f"hint changed: {row[\'hint_ref\']!r} \u2192 {row[\'hint\']!r}",
            "reference":  row["hint_ref"],
        })
    return issues'''

NEW_HINT = '''def compare_hint_changes(current: pl.DataFrame, reference: pl.DataFrame) -> pl.DataFrame:
    """Flag questions where the hint text differs from the template (medium severity)."""
    EMPTY = {
        "issue_type": pl.Utf8, "set_name": pl.Utf8, "Q Name": pl.Utf8,
        "field": pl.Utf8, "current": pl.Utf8, "reference": pl.Utf8,
        "severity": pl.Utf8, "excel_row": pl.Int64,
    }
    ref_hints = reference.filter(pl.col("hint") != "").select(["Q Name", "hint"])
    if ref_hints.is_empty():
        return pl.DataFrame(schema=EMPTY)
    joined = (
        current.select(["Q Name", "hint", "excel_row"])
        .join(ref_hints, on="Q Name", how="inner", suffix="_ref")
        .filter(pl.col("hint") != pl.col("hint_ref"))
    )
    if joined.is_empty():
        return pl.DataFrame(schema=EMPTY)
    return (
        joined.with_columns([
            pl.lit("hint_changed").alias("issue_type"),
            pl.lit("").alias("set_name"),
            pl.col("hint").alias("current"),
            pl.col("hint_ref").alias("reference"),
            pl.lit("medium").alias("severity"),
            pl.lit("hint").alias("field"),
        ])
        .select(list(EMPTY.keys()))
    )'''

if OLD_HINT in src:
    src = src.replace(OLD_HINT, NEW_HINT, 1)
    set_src("abeb4da3", src)
    print("OK abeb4da3 compare_hint_changes fixed")
else:
    print("WARN abeb4da3 OLD_HINT not found verbatim — trying partial match")
    # find and replace from function def to end of function
    start = src.find("def compare_hint_changes")
    if start == -1:
        print("ERROR compare_hint_changes not found at all"); sys.exit(1)
    # find next def after it
    next_def = src.find("\ndef ", start + 4)
    if next_def == -1:
        next_def = len(src)
    src = src[:start] + NEW_HINT + src[next_def:]
    set_src("abeb4da3", src)
    print("OK abeb4da3 compare_hint_changes replaced via partial match")

# ── Fix validate_crop_harvest in 6df04d5a ────────────────────────────────────
src = get_src("6df04d5a")

# Find and replace the whole function
start = src.find("def validate_crop_harvest")
if start == -1:
    print("ERROR validate_crop_harvest not found"); sys.exit(1)

next_def = src.find("\ndef ", start + 4)
if next_def == -1:
    next_def = len(src)

NEW_HARVEST = '''def validate_crop_harvest(survey: pl.DataFrame, crop_harvest_cfg: dict) -> pl.DataFrame:
    """
    PASS if only the minimal set is present (no extra full-set questions), OR
    if the full set is completely present (extra questions beyond full are fine).
    FAIL for any partial overlap.
    Returns a DataFrame with ISSUE_SCHEMA columns.
    """
    EMPTY = {
        "issue_type": pl.Utf8, "set_name": pl.Utf8, "Q Name": pl.Utf8,
        "field": pl.Utf8, "current": pl.Utf8, "reference": pl.Utf8,
        "severity": pl.Utf8, "excel_row": pl.Int64,
    }
    if not crop_harvest_cfg:
        return pl.DataFrame(schema=EMPTY)

    minimal = set(crop_harvest_cfg.get("minimal", []))
    full    = set(crop_harvest_cfg.get("full", []))
    present = set(survey["Q Name"].to_list())

    if full.issubset(present):
        return pl.DataFrame(schema=EMPTY)

    extra_full = (full - minimal) & present
    if (minimal & present) == minimal and not extra_full:
        return pl.DataFrame(schema=EMPTY)

    missing_from_full = full - present
    partial_found     = (full - minimal) & present
    detail_parts = []
    if missing_from_full:
        detail_parts.append(f"missing from full set: {sorted(missing_from_full)}")
    if partial_found:
        detail_parts.append(f"partial full-set Qs present: {sorted(partial_found)}")
    detail = "; ".join(detail_parts) or "partial crop-harvest set"
    ref    = f"minimal={sorted(minimal)}, full={sorted(full)}"

    return pl.DataFrame([{
        "issue_type": "crop_harvest_violation",
        "set_name":   "CRP_HARV",
        "Q Name":     "crp_harv_*",
        "field":      "",
        "current":    detail,
        "reference":  ref,
        "severity":   "high",
        "excel_row":  -1,
    }], schema=EMPTY)'''

src = src[:start] + NEW_HARVEST + src[next_def:]
set_src("6df04d5a", src)
print("OK 6df04d5a validate_crop_harvest fixed")

# ── Save ──────────────────────────────────────────────────────────────────────
with open(NB, "w", encoding="utf-8") as f:
    json.dump(nb, f, ensure_ascii=False, indent=1)

print("\nFix applied. Notebook saved.")
