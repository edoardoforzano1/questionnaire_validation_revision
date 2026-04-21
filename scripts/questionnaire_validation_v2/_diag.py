"""
Diagnostic: run directly on the test file to check if hdds + relevant are caught.
"""
import sys, re, yaml
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
import openpyxl, polars as pl
from pathlib import Path

TEST_FILE = "C:/Temp/DIEM_Monitoring questionnaire_kobo_en_AFG_20260415_test.xlsx"
CFG_FILE  = "c:/Users/edoar/WORK/FAO/repo/questionnaire_validation_revision/scripts/questionnaire_validation_v2/validation_config.yaml"
CRIT_FILE = "c:/Users/edoar/WORK/FAO/repo/questionnaire_validation_revision/scripts/critical_sets.yaml"
TEMPLATES = "c:/Users/edoar/WORK/FAO/repo/questionnaire_validation_revision/Questionnaires/Questionnaire Templates"

# ── 1. find reference template ────────────────────────────────────────────────
candidates = sorted(Path(TEMPLATES).glob("*kobo*EN*template*.xlsx"),
                    key=lambda p: p.stat().st_mtime, reverse=True)
if not candidates:
    print("ERROR: no template found"); sys.exit(1)
ref_path = candidates[0]
print(f"Template : {ref_path.name}")
print(f"Test file: {Path(TEST_FILE).name}")

# ── 2. raw read of survey sheets ──────────────────────────────────────────────
def raw_survey(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb["survey"]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows:
        return {}
    hdr = [str(h).strip() if h is not None else "" for h in rows[0]]
    records = {}
    for row in rows[1:]:
        d = dict(zip(hdr, row))
        name = str(d.get("name") or "").strip()
        if name:
            records[name] = d
    return records

print("\nReading survey sheets...")
curr = raw_survey(TEST_FILE)
ref  = raw_survey(str(ref_path))

print(f"Current questions : {len(curr)}")
print(f"Reference questions: {len(ref)}")

# ── 3. Check HDDS ─────────────────────────────────────────────────────────────
print("\n--- HDDS check ---")
print(f"'hdds' in current : {'hdds' in curr}")
print(f"'hdds' in reference: {'hdds' in ref}")
if 'hdds' not in curr and 'hdds' in ref:
    print("EXPECTED: hdds is missing from test file -> should produce missing_critical_question")

# ── 4. Check broken relevant ──────────────────────────────────────────────────
print("\n--- Broken relevant references in current ---")
REF_RE = re.compile(r"\$\{([^}]+)\}")
curr_names = set(curr.keys())
broken_found = []
for qname, row in curr.items():
    rel = str(row.get("relevant") or "").strip()
    if not rel:
        continue
    refs = REF_RE.findall(rel)
    broken = [v for v in refs if v not in curr_names]
    if broken:
        broken_found.append((qname, broken, rel[:120]))

if broken_found:
    print(f"Found {len(broken_found)} question(s) with broken relevant references:")
    for qname, broken, rel in broken_found:
        print(f"  Q: {qname}  broken: {broken}")
        print(f"     relevant: {rel}")
else:
    print("No broken relevant references found in test file.")
    print("-> Did you save the test file after adding the broken variable?")

# ── 5. Check critical_sets yaml ───────────────────────────────────────────────
print("\n--- Critical sets yaml ---")
with open(CRIT_FILE, encoding='utf-8') as f:
    rules = yaml.safe_load(f)
print(f"exact_sets   : {list(rules.get('exact_sets', {}).keys())}")
print(f"min_count_sets: {list(rules.get('min_count_sets', {}).keys())}")
print(f"crop_harvest  : {rules.get('crop_harvest', {})}")

# ── 6. Simulate validate_critical_sets logic ──────────────────────────────────
print("\n--- Simulate validate_critical_sets ---")
present = set(curr.keys())
for set_name, set_rules in rules.get("exact_sets", {}).items():
    print(f"  Set: {set_name}")
    for rule in set_rules:
        qn = rule["q_name"]
        req = rule.get("required", True)
        in_file = qn in present
        print(f"    {qn}: present={in_file}  required={req}")
        if not in_file and req:
            print(f"    --> SHOULD FLAG as missing_critical_question (high)")
        elif not in_file and not req:
            print(f"    --> SHOULD FLAG as advisory_question (medium)")
