# Sheet Configuration
- VBA CodeName: Sheet2
- Excel UI Name: README

| Address | Name | Value / Label | Formula | Style |
| :--- | :--- | :--- | :--- | :--- |
| A1 | - | # pptx_yaml_excel | - | FG:#800000 |
| A2 | - |  | - | BG:#FFFFFF |
| A3 | - | Excel VBA toolchain to extract translatable text from PowerPoint into YAML, edit it, and apply it back to a presentation copy. | - | Normal |
| A4 | - |  | - | BG:#FFFFFF |
| A5 | - | ## Download | - | FG:#800000 |
| A6 | - |  | - | BG:#FFFFFF |
| A7 | - | **Users**: Download the latest `pptx_yaml_i18n.xlsm` from the link below. No installation required — open the file in Excel and enable macros. | - | Normal |
| A8 | - |  | - | BG:#FFFFFF |
| A9 | - | [**Download pptx_yaml_i18n.xlsm (latest release)**](https://github.com/minipoisson/pptx_yaml_excel/releases/latest/download/pptx_yaml_i18n.xlsm) | - | Normal |
| A10 | - |  | - | BG:#FFFFFF |
| A11 | - | **Developers**: Clone this repository. It serves as a working example of an [xlsm_devkit](https://github.com/minipoisson/xlsm_devkit)-based project, with `src/*.bas` and `sheet/*.md` under version control. | - | Normal |
| A12 | - |  | - | BG:#FFFFFF |
| A13 | - | ## Scope and Assumptions | - | FG:#800000 |
| A14 | - |  | - | BG:#FFFFFF |
| A15 | - | - Primary use case is translation, not iterative slide editing. | - | Normal |
| A16 | - | - YAML keys are derived from slide structure, shape names, table coordinates, and paragraph order. | - | Normal |
| A17 | - | - If the source presentation structure changes, export YAML again before applying. | - | Normal |
| A18 | - | - Applying an old YAML file to a structurally edited deck may produce unexpected key mismatches. | - | Normal |
| A19 | - |  | - | BG:#FFFFFF |
| A20 | - | ## Workflow | - | FG:#800000 |
| A21 | - |  | - | BG:#FFFFFF |
| A22 | - | 1. Select source PowerPoint and export YAML. | - | Normal |
| A23 | - | 2. Edit YAML values. | - | Normal |
| A24 | - | 3. Select edited YAML and base PowerPoint. | - | Normal |
| A25 | - | 4. Apply YAML to create an output PowerPoint copy. | - | Normal |
| A26 | - |  | - | BG:#FFFFFF |
| A27 | - | The tool reports: | - | Normal |
| A28 | - |  | - | BG:#FFFFFF |
| A29 | - | - applied keys | - | Normal |
| A30 | - | - keys that exist only in PowerPoint | - | Normal |
| A31 | - | - keys that exist only in YAML | - | Normal |
| A32 | - | - malformed YAML lines | - | Normal |
| A33 | - |  | - | BG:#FFFFFF |
| A34 | - | ## YAML Key Model | - | FG:#800000 |
| A35 | - |  | - | BG:#FFFFFF |
| A36 | - | Key format is: | - | Normal |
| A37 | - |  | - | BG:#FFFFFF |
| A38 | - | ``` | - | Normal |
| A39 | - | sNN.shapePath.pNN | - | Normal |
| A40 | - | ``` | - | Normal |
| A41 | - |  | - | BG:#FFFFFF |
| A42 | - | - `sNN`: slide index | - | Normal |
| A43 | - | - `shapePath`: shape-name-based path (group hierarchy included) | - | Normal |
| A44 | - | - table cells are encoded as `.rNNcNN` | - | Normal |
| A45 | - | - `pNN`: paragraph index | - | Normal |
| A46 | - |  | - | BG:#FFFFFF |
| A47 | - | Duplicate sibling shape names are disambiguated with suffixes like `_2`, `_3`. | - | Normal |
| A48 | - |  | - | BG:#FFFFFF |
| A49 | - | ## Main Sheet Contract | - | FG:#800000 |
| A50 | - |  | - | BG:#FFFFFF |
| A51 | - | The Main sheet is intentionally address-based for transparent maintenance. | - | Normal |
| A52 | - |  | - | BG:#FFFFFF |
| A53 | - | - `B5`: source PowerPoint | - | Normal |
| A54 | - | - `B8`: exported YAML path | - | Normal |
| A55 | - | - `B11`: edited YAML path | - | Normal |
| A56 | - | - `B14`: base PowerPoint for apply | - | Normal |
| A57 | - | - `B17`: output PowerPoint path | - | Normal |
| A58 | - | - `B20`: output folder | - | Normal |
| A59 | - |  | - | BG:#FFFFFF |
| A60 | - | ## Main Sheet Protection Policy | - | FG:#800000 |
| A61 | - |  | - | BG:#FFFFFF |
| A62 | - | For release builds, the Main sheet is protected without a password to reduce accidental edits. | - | Normal |
| A63 | - | This protection is an operational safeguard and not a security boundary. | - | Normal |
| A64 | - |  | - | BG:#FFFFFF |
| A65 | - | - user-editable cells: `D1`, `B5`, `B11`, `B14` | - | Normal |
| A66 | - | - locked cells include formula/UI cells and internal output cells: `B8`, `B17`, `B20` | - | Normal |
| A67 | - | - on each open, `Workbook_Open` reapplies sheet protection with `UserInterfaceOnly:=True` | - | Normal |
| A68 | - |  | - | BG:#FFFFFF |
| A69 | - | Because `UserInterfaceOnly` is not persisted by Excel, protection settings are re-applied every time the workbook opens. | - | Normal |
| A70 | - | If macros are disabled, this runtime re-apply step will not run. | - | Normal |
| A71 | - |  | - | BG:#FFFFFF |
| A72 | - | ## Localization Design | - | FG:#800000 |
| A73 | - |  | - | BG:#FFFFFF |
| A74 | - | - UI labels are localized by worksheet formulas. | - | Normal |
| A75 | - | - VBA localizes only runtime dialog messages. | - | Normal |
| A76 | - | - VBA reads language state from `Sheet1!D1` and language column index from `Sheet1!F1`. | - | Normal |
| A77 | - | - Translation text is stored in the `i18n` worksheet. | - | Normal |
| A78 | - |  | - | BG:#FFFFFF |
| A79 | - | ## Development | - | FG:#800000 |
| A80 | - |  | - | BG:#FFFFFF |
| A81 | - | This project was developed using [xlsm_devkit](https://github.com/minipoisson/xlsm_devkit), a development toolkit for Excel VBA projects. | - | Normal |
| A82 | - |  | - | BG:#FFFFFF |
| A83 | - | - `src/*.bas` — VBA module files (UTF-8) managed by xlsm_devkit's `ExportAllModules` and `ImportAllModules` macros | - | Normal |
| A84 | - | - `sheet/*.md` — Markdown sheet maps generated by xlsm_devkit's `ExportAllSheetMapsToMD` macro | - | Normal |
| A85 | - |  | - | BG:#FFFFFF |
| A86 | - | ### Prerequisites | - | FG:#800000 |
| A87 | - |  | - | BG:#FFFFFF |
| A88 | - | - Windows + Microsoft Excel with VBA | - | Normal |
| A89 | - | - Enable **Trust access to the VBA project object model** in Excel: | - | Normal |
| A90 | - |     ``` | - | Normal |
| A91 | - |     File → Options → Trust Center → Trust Center Settings | - | Normal |
| A92 | - |       → Macro Settings → Trust access to the VBA project object model | - | Normal |
| A93 | - |     ``` | - | Normal |
| A94 | - | - xlsm_devkit handles UTF-8 ↔ ANSI conversion so that `.bas` files on disk are always UTF-8 while remaining compatible with the VBE. | - | Normal |
