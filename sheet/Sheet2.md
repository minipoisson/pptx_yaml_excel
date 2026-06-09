# Sheet Configuration
- VBA CodeName: Sheet2
- Excel UI Name: README

| Address | Name | Value / Label | Formula | Style |
| :--- | :--- | :--- | :--- | :--- |
| A1 | - | # pptx_yaml_excel | - | FG:#800000; Bold |
| A2 | - |  | - | BG:#FFFFFF |
| A3 | - | Excel VBA toolchain to extract translatable text from PowerPoint into YAML, edit it, and apply it back to a presentation copy. | - | - |
| A4 | - |  | - | BG:#FFFFFF |
| A5 | - | ## Download | - | FG:#800000; Bold |
| A6 | - |  | - | BG:#FFFFFF |
| A7 | - | **Users**: Download the latest `pptx_yaml_i18n.xlsm` from the link below. No installation required — open the file in Excel and enable macros. | - | - |
| A8 | - |  | - | BG:#FFFFFF |
| A9 | - | [**Download pptx_yaml_i18n.xlsm (latest release)**](https://github.com/minipoisson/pptx_yaml_excel/releases/latest/download/pptx_yaml_i18n.xlsm) | - | - |
| A10 | - |  | - | BG:#FFFFFF |
| A11 | - | **Developers**: Clone this repository. It serves as a working example of an [xlsm_devkit](https://github.com/minipoisson/xlsm_devkit)-based project, with `src/*.bas` and `sheet/*.md` under version control. | - | - |
| A12 | - |  | - | BG:#FFFFFF |
| A13 | - | ## Scope and Assumptions | - | FG:#800000; Bold |
| A14 | - |  | - | BG:#FFFFFF |
| A15 | - | - Primary use case is translation, not iterative slide editing. | - | - |
| A16 | - | - YAML keys are derived from slide structure, shape names, table coordinates, and paragraph order. | - | - |
| A17 | - | - If the source presentation structure changes, export YAML again before applying. | - | - |
| A18 | - | - Applying an old YAML file to a structurally edited deck may produce unexpected key mismatches. | - | - |
| A19 | - |  | - | BG:#FFFFFF |
| A20 | - | ## Workflow | - | FG:#800000; Bold |
| A21 | - |  | - | BG:#FFFFFF |
| A22 | - | 1. Select source PowerPoint and export YAML. | - | - |
| A23 | - | 2. Edit YAML values. | - | - |
| A24 | - | 3. Select edited YAML and base PowerPoint. | - | - |
| A25 | - | 4. Apply YAML to create an output PowerPoint copy. | - | - |
| A26 | - |  | - | BG:#FFFFFF |
| A27 | - | The tool reports: | - | - |
| A28 | - |  | - | BG:#FFFFFF |
| A29 | - | - applied keys | - | - |
| A30 | - | - keys that exist only in PowerPoint | - | - |
| A31 | - | - keys that exist only in YAML | - | - |
| A32 | - | - malformed YAML lines | - | - |
| A33 | - |  | - | BG:#FFFFFF |
| A34 | - | ## YAML Key Model | - | FG:#800000; Bold |
| A35 | - |  | - | BG:#FFFFFF |
| A36 | - | Key format is: | - | - |
| A37 | - |  | - | BG:#FFFFFF |
| A38 | - | ``` | - | - |
| A39 | - | sNN.shapePath.pNN | - | - |
| A40 | - | ``` | - | - |
| A41 | - |  | - | BG:#FFFFFF |
| A42 | - | - `sNN`: slide index | - | - |
| A43 | - | - `shapePath`: shape-name-based path (group hierarchy included) | - | - |
| A44 | - | - table cells are encoded as `.rNNcNN` | - | - |
| A45 | - | - `pNN`: paragraph index | - | - |
| A46 | - |  | - | BG:#FFFFFF |
| A47 | - | Duplicate sibling shape names are disambiguated with suffixes like `_2`, `_3`. | - | - |
| A48 | - |  | - | BG:#FFFFFF |
| A49 | - | ## Main Sheet Contract | - | FG:#800000; Bold |
| A50 | - |  | - | BG:#FFFFFF |
| A51 | - | The Main sheet is intentionally address-based for transparent maintenance. | - | - |
| A52 | - |  | - | BG:#FFFFFF |
| A53 | - | - `B5`: source PowerPoint | - | - |
| A54 | - | - `B8`: exported YAML path | - | - |
| A55 | - | - `B11`: edited YAML path | - | - |
| A56 | - | - `B14`: base PowerPoint for apply | - | - |
| A57 | - | - `B17`: output PowerPoint path | - | - |
| A58 | - | - `B20`: output folder | - | - |
| A59 | - |  | - | BG:#FFFFFF |
| A60 | - | ## Main Sheet Protection Policy | - | FG:#800000; Bold |
| A61 | - |  | - | BG:#FFFFFF |
| A62 | - | For release builds, the Main sheet is protected without a password to reduce accidental edits. | - | - |
| A63 | - | This protection is an operational safeguard and not a security boundary. | - | - |
| A64 | - |  | - | BG:#FFFFFF |
| A65 | - | - user-editable cells: `D1`, `B5`, `B11`, `B14` | - | - |
| A66 | - | - locked cells include formula/UI cells and internal output cells: `B8`, `B17`, `B20` | - | - |
| A67 | - | - on each open, `Workbook_Open` reapplies sheet protection with `UserInterfaceOnly:=True` | - | - |
| A68 | - |  | - | BG:#FFFFFF |
| A69 | - | Because `UserInterfaceOnly` is not persisted by Excel, protection settings are re-applied every time the workbook opens. | - | - |
| A70 | - | If macros are disabled, this runtime re-apply step will not run. | - | - |
| A71 | - |  | - | BG:#FFFFFF |
| A72 | - | ## Localization Design | - | FG:#800000; Bold |
| A73 | - |  | - | BG:#FFFFFF |
| A74 | - | - UI labels are localized by worksheet formulas. | - | - |
| A75 | - | - VBA localizes only runtime dialog messages. | - | - |
| A76 | - | - VBA reads language state from `Sheet1!D1` and language column index from `Sheet1!F1`. | - | - |
| A77 | - | - Translation text is stored in the `i18n` worksheet. | - | - |
| A78 | - |  | - | BG:#FFFFFF |
| A79 | - | ## Development | - | FG:#800000; Bold |
| A80 | - |  | - | BG:#FFFFFF |
| A81 | - | This project was developed using [xlsm_devkit](https://github.com/minipoisson/xlsm_devkit), a development toolkit for Excel VBA projects. | - | - |
| A82 | - |  | - | BG:#FFFFFF |
| A83 | - | - `src/*.bas` — VBA module files (UTF-8) managed by xlsm_devkit's `ExportAllModules` and `ImportAllModules` macros | - | - |
| A84 | - | - `sheet/*.md` — Markdown sheet maps generated by xlsm_devkit's `ExportAllSheetMapsToMD` macro | - | - |
| A85 | - |  | - | BG:#FFFFFF |
| A86 | - | ### Prerequisites | - | FG:#800000; Bold |
| A87 | - |  | - | BG:#FFFFFF |
| A88 | - | - Windows + Microsoft Excel with VBA | - | - |
| A89 | - | - Enable **Trust access to the VBA project object model** in Excel: | - | - |
| A90 | - |     ``` | - | - |
| A91 | - |     File → Options → Trust Center → Trust Center Settings | - | - |
| A92 | - |       → Macro Settings → Trust access to the VBA project object model | - | - |
| A93 | - |     ``` | - | - |
| A94 | - | - xlsm_devkit handles UTF-8 ↔ ANSI conversion so that `.bas` files on disk are always UTF-8 while remaining compatible with the VBE. | - | - |
