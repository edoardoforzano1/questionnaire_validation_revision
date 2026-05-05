# DIEM Questionnaire Validation v2

Automated validation of DIEM household questionnaires for two data-collection tools: **GeoPoll** and **KoBo**. The process compares a current questionnaire against a reference (either the latest official template or the previous round's validated file) and produces a structured Excel report flagging issues by severity.

---

## How it works

Both tools share the same configuration file (`validation_config.yaml`) and the same validation concept: load the current questionnaire, load the reference, compare them question-by-question and option-by-option, classify any deviation as an issue, and export a report.

### Reference modes

The key choice each run is **what to compare against**:

| Mode | When to use | Reference file |
|---|---|---|
| `latest_template` | Validating a new questionnaire against the official template | Auto-detected from `templates_dir` (newest matching file) |
| `previous_round` | Checking consistency between consecutive rounds | Specified explicitly as `previous_round_file` |

**Example — `latest_template`:** Country team submits a new GeoPoll questionnaire for Yemen. The notebook finds the newest Arabic GeoPoll template in the templates folder and checks whether all mandatory questions are present with the correct labels.

**Example — `previous_round`:** Round 11 KoBo questionnaire for DRC is ready. The notebook loads the validated Round 10 file and flags any questions that were removed, made optional, or had their skip logic broken.

---

## Quick start

1. Open `validation_config.yaml` and set the fields that change each run:

```yaml
tool: "geopoll"            # kobo | geopoll
questionnaire_file: "DRC_R11 HH questionnaire_Geopoll.xlsx"
language: "fr"
iso3: "COD"
reference_mode: "previous_round"
previous_round_file: "validated_questionnaire_geopoll_fr_COD_20250423_r10.xlsx"
working_dir: "C:/Temp"
```

2. Open the matching notebook (`questionnaire_validation_geopoll.ipynb` or `questionnaire_validation_kobo.ipynb`).
3. Run all cells. Output is written to `output_dir/geopoll_output/` or `output_dir/kobo_output/`.

> **Never edit the notebook code.** All configuration lives in `validation_config.yaml`.

---

## What gets validated

Both tools run the following checks:

| Check | Severity |
|---|---|
| Non-optional question removed from current | **High** |
| Non-optional question added to current | **Medium** |
| Mandatory flag changed | **High** |
| Answer option label changed | **Medium** |
| Optional question (`o_*`) added or removed | Info |
| Critical question sets incomplete or misconfigured | **High** |

### GeoPoll-specific checks

- **Prefix counts**: certain question groups (CS coping strategies, HDDS) must contain an exact number of questions. Too few → High.
- **Crop/harvest rule**: a questionnaire must include either the minimal or the full crop/harvest question set — partial sets are not allowed.
- **Skip pattern references**: all variables referenced in skip conditions must exist in the questionnaire. Broken references → High.

### KoBo-specific checks

- **Relevant (skip logic) broken references**: if a `relevant` expression references a variable that no longer exists, the question will silently never show. Flagged as High.
- **Relevant modifications**: changes to skip logic expressions between rounds. Flagged as Medium.
- **Placeholder normalization**: template questions contain `#placeholder#` tokens (e.g. for country-specific crop names). Before comparing labels, placeholders are restored, and the report shows the actual filled values alongside the template text.
- **Validated questionnaire production**: in `previous_round` mode the KoBo notebook also produces a *validated questionnaire* file — a copy of the current questionnaire with:
  - Crop choices injected from the local "Crop list" sheet (country-specific)
  - Admin1/Admin2 choices refreshed live from FAO AGOL (no credentials required)
  - Template labels restored for placeholder questions

---

## Output files

| File | Contents |
|---|---|
| `validation_report_*.xlsx` | Multi-sheet report (see below) |
| `validated_questionnaire_*.xlsx` | *(KoBo only)* Production-ready questionnaire with injected lists |

### Report sheets

**GeoPoll report (4 sheets):**
- **Summary** — per-check-set PASS/FAIL with issue counts
- **Critical Sets** — structural issues (mandatory flags, critical question groups)
- **Template Changes** — presence, mandatory, and option-label differences
- **Optional Questions** — informational log of optional question movements

**KoBo report (5 sheets):**
- **Summary**
- **Critical Sets**
- **Relevant Changes** — broken references (High) and modified expressions (Medium)
- **Question Changes** — presence, mandatory, label, constraint, and choices-list differences
- **Option Changes** — option additions, removals, and label changes

Color coding is consistent across both reports: red = High, orange = Medium, blue = Info, green = PASS.

---

## Config profiles

For teams that switch frequently between countries or rounds, reusable YAML profiles can live in `<output_dir>/config_profiles/` (or any folder set via `config_profiles_dir`). A profile contains only the fields that differ from the base config — everything else inherits from `validation_config.yaml`.

```yaml
# config_profiles/geopoll_ar_yem_latest.yaml
tool: "geopoll"
questionnaire_file: "validated_questionnaire_geopoll_ar_YEM_20251013_TEST.xlsx"
language: "ar"
iso3: "YEM"
reference_mode: "latest_template"
```

Activate by setting `config_profile: "geopoll_ar_yem_latest.yaml"` in `validation_config.yaml` and running the notebook. The resolved configuration is saved to `<output_dir>/configuration/` for traceability.
