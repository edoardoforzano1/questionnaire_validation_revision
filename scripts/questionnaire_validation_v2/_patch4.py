"""
Fix: Summary sheet exact_sets loop uses _crit (critical_issues param) which is empty/wrong.
Use all_issues instead — consistent with CRP_HARV and SKIP LOGIC rows.
Also: advisory_question (required=false) should not flip status to FAIL.
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

src = get_src("9e9338a9")

OLD_LOOP = '''    for set_name in _rules.get("exact_sets", {}):
        set_issues = _crit.filter(pl.col("set_name") == set_name)
        passed = set_issues.height == 0
        if passed:
            details = "All required questions present"
        else:
            missing = [q for q in set_issues.filter(
                pl.col("issue_type").is_in(["missing_critical_question", "partial_critical_set"])
            )["Q Name"].to_list() if q]
            details = f"Missing: {', '.join(missing)}" if missing else "Issues found"
        _set_row(ws, r, set_name, passed, details); r += 1'''

NEW_LOOP = '''    _EXACT_FAIL_TYPES = ["missing_critical_question", "partial_critical_set",
                          "critical_mandatory_mismatch"]
    for set_name in _rules.get("exact_sets", {}):
        set_fail = all_issues.filter(
            (pl.col("set_name") == set_name) & pl.col("issue_type").is_in(_EXACT_FAIL_TYPES)
        )
        set_warn = all_issues.filter(
            (pl.col("set_name") == set_name) & (pl.col("issue_type") == "advisory_question")
        )
        passed = set_fail.height == 0
        if passed and set_warn.height == 0:
            details = "All required questions present"
        elif passed:
            warn_qs = [q for q in set_warn["Q Name"].to_list() if q]
            details = f"Required OK  (advisory missing: {', '.join(warn_qs)})"
        else:
            missing = [q for q in set_fail["Q Name"].to_list() if q]
            details = f"Missing: {', '.join(missing)}" if missing else "Issues found"
        _set_row(ws, r, set_name, passed, details); r += 1'''

if OLD_LOOP in src:
    src = src.replace(OLD_LOOP, NEW_LOOP, 1)
    set_src("9e9338a9", src)
    print("OK 9e9338a9: exact_sets loop fixed to use all_issues")
else:
    print("ERROR: exact_sets loop anchor not found")
    sys.exit(1)

with open(NB, "w", encoding="utf-8") as f:
    json.dump(nb, f, ensure_ascii=False, indent=1)

print("Notebook saved.")
