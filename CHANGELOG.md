# Changelog

All notable changes to this project will be documented in this file.

## [Unreleased]

## [1.0.4] - 2026-05-07

### Changed
- Updated `xlsm_devkit` integration so `ImportAllModules` no longer depends on the active VBE code pane being `xlsm_devkit`
- Synced workbook update in `pptx_yaml_i18n.xlsm` for the revised import workflow
- Added exported `src/Sheet4.bas` to match the current workbook project

## [1.0.3] - 2026-05-07

### Changed
- Simplified Sheet1 formulas in workbook (`pptx_yaml_i18n.xlsm`)
- Updated `sheet/Sheet1.md` to reflect formula changes

## [1.0.2] - 2026-05-06

### Changed
- Synced repository files for workbook update: `pptx_yaml_i18n.xlsm`, `sheet/Sheet2.md`, `sheet/Sheet4.md`, and `sheet/Sheet5.md`

## [1.0.1] - 2026-05-06

### Changed
- Refreshed `README` and `README.ja` sheets in the workbook with up-to-date content
- Added `LICENSE` sheet to the workbook

## [1.0.0] - 2026-05-05

### Added
- Export PowerPoint text to YAML (`ppt2yaml`)
- Apply YAML back to a PowerPoint copy
- Group and table cell support in key model (`sNN.shapePath.rNNcNN.pNN`)
- Duplicate sibling shape name disambiguation (`_2`, `_3`)
- Speaker notes export/apply
- Multilingual UI via `i18n` worksheet and `LanguageManager`
- System locale auto-detection on workbook open
- Main sheet protection (UI-only) reapplied on each open
- Apply report: applied keys, PowerPoint-only keys, YAML-only keys, malformed lines

[Unreleased]: https://github.com/minipoisson/pptx_yaml_excel/compare/v1.0.2...HEAD
[1.0.2]: https://github.com/minipoisson/pptx_yaml_excel/compare/v1.0.1...v1.0.2
[1.0.1]: https://github.com/minipoisson/pptx_yaml_excel/compare/v1.0.0...v1.0.1
[1.0.0]: https://github.com/minipoisson/pptx_yaml_excel/releases/tag/v1.0.0