# Agent Instructions for Release Operations

## DEV/release workflow

Development is done in `DEV_pptx_yaml_i18n.xlsm` (excluded from version control via `.gitignore`).
The committed `pptx_yaml_i18n.xlsm` is always a clean release build with no devkit modules.

### Starting development

1. Open `pptx_yaml_i18n.xlsm`.
2. Run `CallInitDevMode` → creates `DEV_pptx_yaml_i18n.xlsm` and imports all `devkit_*` modules from `src/`.
3. Close the original without saving; continue work in the `DEV_` workbook.

### Syncing source files

Run `ExportAllModulesFormsSheetMaps` (or use the Launcher) to update `src/` and `sheet/`.

### Creating a release build

Run `CallSaveAsRelease` from the `DEV_` workbook → generates a clean `pptx_yaml_i18n.xlsm` with all devkit modules removed. This file is the release artifact to commit and upload.

---

This repository uses the following default release policy.

## Default policy

- Unless explicitly excluded by the user, treat all tracked/untracked changes shown by git status as release candidates (same intent as Stage All in SOURCE CONTROL).
- Before commit, always show the candidate file list and wait for user confirmation.
- Update CHANGELOG before tagging.
- Create a version tag only after commit is pushed.
- Create GitHub Release from the tag.
- Always upload release asset: pptx_yaml_i18n.xlsm.
- After release, verify all of the following:
  - git working tree is clean
  - tag exists on origin
  - release page exists
  - release includes asset pptx_yaml_i18n.xlsm

## Release decision rule

- Create a new release when users should download a new workbook.
- In this repository, treat these as release-worthy by default:
  - Changes that modify behavior, UI, import/export results, or compatibility.
  - Any update to the distributed workbook file (`pptx_yaml_i18n.xlsm`).
- Do not create a release for repository-only maintenance updates:
  - Developer workflow docs, agent instructions, CI/config cleanup, or internal notes.
- If uncertain, ask the user whether this change should be user-facing and therefore released.

## Standard execution sequence

1. Show git status --short and candidate file list.
2. Ask for confirmation (or use user-provided include/exclude list).
3. Update CHANGELOG.md for the target version.
4. git add (all candidates unless user limits scope).
5. git commit with a clear message.
6. git push origin main.
7. git tag <version> and git push origin <version>.
8. gh release create <version> pptx_yaml_i18n.xlsm with release notes.
9. Final verification summary.

## Minimal user prompt template

Use this template for fast and reliable execution:

"Run the standard release workflow in AGENTS.md for vX.Y.Z.
Use default scope (all pending changes), attach pptx_yaml_i18n.xlsm,
and show candidate files before commit."

## Optional override template

"Run standard release workflow for vX.Y.Z.
Exclude: <paths>
Include only: <paths>
Release notes focus: <summary>."
