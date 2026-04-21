"""
Patch: add hint comparison + validate_crop_harvest to the KoBO validation notebook.
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

errors = []

# ── 1. ecb8a490  read_kobo_survey: add hint ───────────────────────────────────
src = get_src("ecb8a490")

old1 = '"appearance": pl.Utf8, "calculation": pl.Utf8,'
new1 = '"appearance": pl.Utf8, "calculation": pl.Utf8, "hint": pl.Utf8,'
if old1 in src:
    src = src.replace(old1, new1, 1); print("OK ecb8a490 EMPTY hint")
elif '"hint": pl.Utf8' in src:
    print("INFO ecb8a490 EMPTY hint already present")
else:
    errors.append("ecb8a490 EMPTY hint anchor not found")

old2 = '"calculation"       : str(get(row, "calculation") or "").strip(),'
new2 = ('"calculation"       : str(get(row, "calculation") or "").strip(),\n'
        '            "hint"             : str(get(row, "hint")        or "").strip(),')
if old2 in src:
    src = src.replace(old2, new2, 1); print("OK ecb8a490 records hint")
elif '"hint"' in src:
    print("INFO ecb8a490 records hint already present")
else:
    errors.append("ecb8a490 records hint anchor not found")

set_src("ecb8a490", src)

# ── 2. abeb4da3  comparison functions: append compare_hint_changes ────────────
src = get_src("abeb4da3")
if "def compare_hint_changes" not in src:
    src += '''

def compare_hint_changes(current: pl.DataFrame, reference: pl.DataFrame) -> list[dict]:
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
    return issues
'''
    set_src("abeb4da3", src); print("OK abeb4da3 compare_hint_changes appended")
else:
    print("INFO abeb4da3 compare_hint_changes already present")

# ── 3. 6df04d5a  critical-sets: append validate_crop_harvest ─────────────────
src = get_src("6df04d5a")
if "def validate_crop_harvest" not in src:
    src += '''

def validate_crop_harvest(survey: pl.DataFrame, crop_harvest_cfg: dict) -> list[dict]:
    """
    PASS if only the minimal set is present (no extra full-set questions), OR
    if the full set is completely present (extra questions beyond full are fine).
    FAIL for any partial overlap.
    """
    if not crop_harvest_cfg:
        return []
    minimal = set(crop_harvest_cfg.get("minimal", []))
    full    = set(crop_harvest_cfg.get("full", []))
    present = set(survey["Q Name"].to_list())

    if full.issubset(present):
        return []

    extra_full = (full - minimal) & present
    if (minimal & present) == minimal and not extra_full:
        return []

    missing_from_full = full - present
    partial_found     = (full - minimal) & present
    detail_parts = []
    if missing_from_full:
        detail_parts.append(f"missing from full set: {sorted(missing_from_full)}")
    if partial_found:
        detail_parts.append(f"partial full-set Qs present: {sorted(partial_found)}")

    return [{
        "issue_type": "crop_harvest_violation",
        "Q Name":     "crp_harv_*",
        "severity":   "high",
        "detail":     "; ".join(detail_parts) or "partial crop-harvest set",
        "reference":  f"minimal={sorted(minimal)}, full={sorted(full)}",
    }]
'''
    set_src("6df04d5a", src); print("OK 6df04d5a validate_crop_harvest appended")
else:
    print("INFO 6df04d5a validate_crop_harvest already present")

# ── 4. 21f27652  vanilla cols: add hint ───────────────────────────────────────
src = get_src("21f27652")
old_v = '_VANILLA_COLS_SURVEY  = ["label", "constraint"]'
new_v = '_VANILLA_COLS_SURVEY  = ["label", "constraint", "hint"]'
if old_v in src:
    src = src.replace(old_v, new_v, 1); set_src("21f27652", src)
    print("OK 21f27652 hint added to _VANILLA_COLS_SURVEY")
elif '"hint"' in src:
    print("INFO 21f27652 hint already in _VANILLA_COLS_SURVEY")
else:
    errors.append("21f27652 _VANILLA_COLS_SURVEY anchor not found")

# ── 5. e9a3563e  pipeline ─────────────────────────────────────────────────────
src = get_src("e9a3563e")

# 5a: hint_issues call after calculation_issues
old5a = 'calculation_issues = compare_calculation_changes(current_survey, reference_survey)'
new5a = ('calculation_issues = compare_calculation_changes(current_survey, reference_survey)\n'
         'hint_issues        = compare_hint_changes(current_survey, reference_survey)')
if old5a in src:
    src = src.replace(old5a, new5a, 1); print("OK e9a3563e hint_issues call")
elif "hint_issues" in src:
    print("INFO e9a3563e hint_issues already present")
else:
    errors.append("e9a3563e calc_issues anchor not found")

# 5b: harvest_issues call after count_issues
old5b = 'count_issues    = validate_prefix_counts(current_survey, rules.get("min_count_sets", {}))'
new5b = ('count_issues    = validate_prefix_counts(current_survey, rules.get("min_count_sets", {}))\n'
         'harvest_issues  = validate_crop_harvest(current_survey, rules.get("crop_harvest", {}))')
if old5b in src:
    src = src.replace(old5b, new5b, 1); print("OK e9a3563e harvest_issues call")
elif "harvest_issues" in src:
    print("INFO e9a3563e harvest_issues already present")
else:
    errors.append("e9a3563e count_issues anchor not found")

# 5c: add hint_issues and harvest_issues to concat
old5c = ('     required_issues, appearance_issues, calculation_issues,\n'
         '     option_label_issues, option_pres_issues,\n'
         '     critical_issues, count_issues, relevant_issues],')
new5c = ('     required_issues, appearance_issues, calculation_issues,\n'
         '     hint_issues,\n'
         '     option_label_issues, option_pres_issues,\n'
         '     critical_issues, count_issues, harvest_issues, relevant_issues],')
if old5c in src:
    src = src.replace(old5c, new5c, 1); print("OK e9a3563e concat updated")
elif "hint_issues" in src and "harvest_issues" in src:
    print("INFO e9a3563e concat already has hint/harvest")
else:
    errors.append("e9a3563e concat anchor not found")

# 5d: update print line to include hint count
old5d = ('print(f"Required: {required_issues.height}  '
         'Appearance: {appearance_issues.height}  Calculation: {calculation_issues.height}")')
new5d = ('print(f"Required: {required_issues.height}  Appearance: {appearance_issues.height}  '
         'Calculation: {calculation_issues.height}  Hint: {hint_issues.height}")')
if old5d in src:
    src = src.replace(old5d, new5d, 1); print("OK e9a3563e print line updated")
elif "Hint:" in src:
    print("INFO e9a3563e print line already updated")
# non-fatal if missing

set_src("e9a3563e", src)

# ── 6. 9e9338a9  export cell ──────────────────────────────────────────────────
src = get_src("9e9338a9")

# 6a: add hint_changed to QUESTION_CHANGE_TYPES
old6a = '"type_changed", "required_changed", "appearance_changed", "calculation_changed",'
new6a = '"type_changed", "required_changed", "appearance_changed", "calculation_changed",\n    "hint_changed",'
if old6a in src:
    src = src.replace(old6a, new6a, 1); print("OK 9e9338a9 hint_changed in QUESTION_CHANGE_TYPES")
elif '"hint_changed"' in src:
    print("INFO 9e9338a9 hint_changed already in QUESTION_CHANGE_TYPES")
else:
    errors.append("9e9338a9 QUESTION_CHANGE_TYPES anchor not found")

# 6b: add crop_harvest_violation to CRITICAL_ISSUE_TYPES
old6b = '"missing_critical_question", "advisory_question", "partial_critical_set",\n    "critical_mandatory_mismatch", "below_minimum_count",'
new6b = ('"missing_critical_question", "advisory_question", "partial_critical_set",\n'
         '    "critical_mandatory_mismatch", "below_minimum_count", "crop_harvest_violation",')
if old6b in src:
    src = src.replace(old6b, new6b, 1); print("OK 9e9338a9 crop_harvest_violation in CRITICAL_ISSUE_TYPES")
elif '"crop_harvest_violation"' in src:
    print("INFO 9e9338a9 crop_harvest_violation already in CRITICAL_ISSUE_TYPES")
else:
    errors.append("9e9338a9 CRITICAL_ISSUE_TYPES anchor not found")

# 6c: add labels
old6c = '    "constraint_modified"      : "Constraint changed",'
new6c = ('    "constraint_modified"      : "Constraint changed",\n'
         '    "hint_changed"             : "Hint text changed",\n'
         '    "crop_harvest_violation"   : "Crop harvest rule violation",')
if old6c in src:
    src = src.replace(old6c, new6c, 1); print("OK 9e9338a9 _ISSUE_LABELS updated")
elif '"hint_changed"' in src:
    print("INFO 9e9338a9 _ISSUE_LABELS already updated")
else:
    errors.append("9e9338a9 _ISSUE_LABELS anchor not found")

# 6d: add CRP_HARV row in summary (after min_count_sets loop, before skip logic)
old6d = '    # Skip logic summary row'
new6d = '''    # Crop harvest row
    if _rules.get("crop_harvest"):
        _harv_fail = all_issues.filter(pl.col("issue_type") == "crop_harvest_violation").height > 0
        _harv_det = ("Partial crop-harvest set — see issues" if _harv_fail
                     else "Crop harvest rule satisfied")
        _set_row(ws, r, "CRP_HARV", not _harv_fail, _harv_det); r += 1

    # Skip logic summary row'''
if old6d in src:
    src = src.replace(old6d, new6d, 1); print("OK 9e9338a9 CRP_HARV summary row added")
elif "CRP_HARV" in src:
    print("INFO 9e9338a9 CRP_HARV already in summary")
else:
    errors.append("9e9338a9 summary CRP_HARV anchor not found")

# 6e: update write_question_changes_sheet description to mention hint
old6e = '"QUESTION CHANGES  Presence, mandatory, label, type, required, appearance, calculation, constraint, choices list", 8)'
new6e = '"QUESTION CHANGES  Presence, mandatory, label, type, hint, required, appearance, calculation, constraint, choices list", 8)'
if old6e in src:
    src = src.replace(old6e, new6e, 1); print("OK 9e9338a9 question changes description")
elif "hint" in src:
    print("INFO 9e9338a9 hint already in question changes description")
# non-fatal

set_src("9e9338a9", src)

# ── Final ─────────────────────────────────────────────────────────────────────
if errors:
    for e in errors:
        print(f"ERROR: {e}")
    sys.exit(1)

with open(NB, "w", encoding="utf-8") as f:
    json.dump(nb, f, ensure_ascii=False, indent=1)

print("\nAll patches applied. Notebook saved.")
