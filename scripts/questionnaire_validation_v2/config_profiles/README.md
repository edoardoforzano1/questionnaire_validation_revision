# Config Profiles

Put reusable YAML profiles here for quick switching.

Use `config_profile` in `validation_config.yaml` to activate one profile.
Profiles can contain full config or only fields to override.

## Default profile folder
If `config_profiles_dir` is empty, notebooks now look in:
- `<output_dir>/config_profiles/`

## Quick use
1. Put profile file in `<output_dir>/config_profiles/` (or set `config_profiles_dir`).
2. Set `config_profile` in `validation_config.yaml` (for example `example_geopoll_latest.yaml`).
3. Run notebook normally.
4. The notebook writes a resolved run snapshot to:
   - Shared: `<output_dir>/configuration/`

Example:
```yaml
questionnaire_file: "validated_questionnaire_geopoll_ar_YEM_20251013124512_TEST.xlsx"
language: "ar"
iso3: "YEM"
reference_mode: "latest_template"
tool: "geopoll"
```
