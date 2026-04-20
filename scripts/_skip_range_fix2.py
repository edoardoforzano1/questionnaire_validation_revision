import json, re

with open('c:/Users/edoar/WORK/FAO/repo/questionnaire_validation_revision/scripts/newQ_testv4.ipynb', encoding='utf-8') as f:
    nb = json.load(f)

src15 = ''.join(nb['cells'][15]['source'])

# ── Fix 1: rewrite _extract_condition_codes to accept the full statement ──────
# The previous version expected the text BEFORE "go to" (no anchor), which caused
# the lazy regex to fail. Now we pass the full stmt and anchor on "go to".

OLD_EXTRACT = '''def _extract_condition_codes(condition_text: str) -> set[int]:
    """
    Extract option codes from the condition part of a Default-column rule,
    e.g. "if hh_gender = 1-4, go to …"  →  {1, 2, 3, 4}.
    Returns empty set when no range is found (e.g. "For any response").
    """
    match = re.search(
        r'=\\s*([\\d][^;,\\n]*?)(?:\\s*,\\s*go|\\s+go|$)',
        condition_text, re.IGNORECASE,
    )
    if not match:
        return set()
    rng = match.group(1).strip()
    codes: set[int] = set()
    for a, b in re.findall(r'\\b(\\d+)\\s*(?:-|to)\\s*(\\d+)\\b', rng, re.IGNORECASE):
        codes.update(range(*sorted((int(a), int(b) + 1))))
    clean = re.sub(r'\\b\\d+\\s*(?:-|to)\\s*\\d+\\b', \' \', rng, flags=re.IGNORECASE)
    codes.update(int(n) for n in re.findall(r\'\\b\\d+\\b\', clean))
    return codes'''

NEW_EXTRACT = '''def _extract_condition_codes(stmt: str) -> set[int]:
    """
    Extract option codes from a full Default-column rule statement by finding
    the range between '=' and 'go to', e.g.:
      "if hh_gender = 1-4, go to hh_education"  →  {1, 2, 3, 4}
      "For any response, go to X"               →  {} (any response = no range)
    Receives the FULL statement (not sliced) so that 'go to' can anchor the match.
    """
    match = re.search(
        r\'=\\s*([\\d][^;=\\n]*?)\\s*,?\\s*go\\s+to\',
        stmt, re.IGNORECASE,
    )
    if not match:
        return set()
    rng = match.group(1).strip()
    codes: set[int] = set()
    for a, b in re.findall(r\'\\b(\\d+)\\s*(?:-|to)\\s*(\\d+)\\b\', rng, re.IGNORECASE):
        codes.update(range(*sorted((int(a), int(b) + 1))))
    clean = re.sub(r\'\\b\\d+\\s*(?:-|to)\\s*\\d+\\b\', \' \', rng, flags=re.IGNORECASE)
    codes.update(int(n) for n in re.findall(r\'\\b\\d+\\b\', clean))
    return codes'''

assert OLD_EXTRACT in src15, '_extract_condition_codes old body not found'
src15 = src15.replace(OLD_EXTRACT, NEW_EXTRACT)

# ── Fix 2: pass full `stmt` to _extract_condition_codes (not the sliced part) ─
OLD_CALL = '''        # Extract option codes from the condition part (text before "go to")
        condition_text = stmt[:go_match.start()]
        option_codes   = _extract_condition_codes(condition_text)'''

NEW_CALL = '''        # Extract option codes from the full statement (regex anchors on "go to")
        option_codes = _extract_condition_codes(stmt)'''

assert OLD_CALL in src15, 'condition_text call not found'
src15 = src15.replace(OLD_CALL, NEW_CALL)

nb['cells'][15]['source'] = src15

with open('c:/Users/edoar/WORK/FAO/repo/questionnaire_validation_revision/scripts/newQ_testv4.ipynb', 'w', encoding='utf-8') as f:
    json.dump(nb, f, ensure_ascii=False, indent=1)

print("Saved. Verifying...")

# Quick smoke-test the fixed function in isolation
exec_globals = {"re": re}
exec("""
import re

def _normalize_skip_text(value):
    text = "" if value is None else str(value)
    return re.sub(r"\\s+", " ", text).strip()

def _extract_condition_codes(stmt):
    match = re.search(r'=\\s*([\\d][^;=\\n]*?)\\s*,?\\s*go\\s+to', stmt, re.IGNORECASE)
    if not match:
        return set()
    rng = match.group(1).strip()
    codes = set()
    for a, b in re.findall(r'\\b(\\d+)\\s*(?:-|to)\\s*(\\d+)\\b', rng, re.IGNORECASE):
        codes.update(range(*sorted((int(a), int(b)+1))))
    clean = re.sub(r'\\b\\d+\\s*(?:-|to)\\s*\\d+\\b', ' ', rng, flags=re.IGNORECASE)
    codes.update(int(n) for n in re.findall(r'\\b\\d+\\b', clean))
    return codes

def _extract_skip_codes_for_target(skip_text, target):
    codes = set()
    for part in re.split(r'[\\r\\n;]+', str(skip_text or "")):
        if target not in part or "=" not in part:
            continue
        left = part.split("=", 1)[0]
        for a, b in re.findall(r'\\b(\\d+)\\s*(?:-|to)\\s*(\\d+)\\b', left, re.IGNORECASE):
            codes.update(range(*sorted((int(a), int(b)+1))))
        clean = re.sub(r'\\b\\d+\\s*(?:-|to)\\s*\\d+\\b', ' ', left, flags=re.IGNORECASE)
        codes.update(int(n) for n in re.findall(r'\\b\\d+\\b', clean))
    return codes

# Tests
cases = [
    ("if hh_gender = 1-4, go to hh_education.", {1,2,3,4}),
    ("if X = 1-8, go to ls_proddif",           {1,2,3,4,5,6,7,8}),
    ("For any response, go to ls_proddif",       set()),
    ("if X = 1, go to ls_proddif_2",            {1}),
    ("if X = 12-13, go to fish_intro",          {12,13}),
]
all_ok = True
for stmt, expected in cases:
    got = _extract_condition_codes(stmt)
    ok = got == expected
    if not ok: all_ok = False
    print(f"  {'OK  ' if ok else 'FAIL'}: {stmt!r} → {got} (expected {expected})")

# Range comparison simulation
def_text  = "if hh_gender = 1-4, go to hh_education."
skip_text = "1-3 = hh_education"
expected  = _extract_condition_codes(def_text)   # {1,2,3,4}
actual    = _extract_skip_codes_for_target(skip_text, "hh_education")  # {1,2,3}
mismatch  = actual and actual != expected
print(f"\\n  hh_gender range mismatch detected: {mismatch}  (expected={expected}, actual={actual})")
if not mismatch: all_ok = False

print()
print("All smoke-tests passed!" if all_ok else "Some smoke-tests FAILED.")
""", exec_globals)
