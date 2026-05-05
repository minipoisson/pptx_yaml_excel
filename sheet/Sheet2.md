# Sheet Configuration
- VBA CodeName: Sheet2
- Excel UI Name: README

| Address | Name | Value / Label | Formula | Style |
| :--- | :--- | :--- | :--- | :--- |
| A1 | - | # pptx_yaml_excel | - | FG:#800000 |
| A2 | - |  | - | BG:#FFFFFF |
| A3 | - | Excel VBA toolchain to extract translatable text from PowerPoint into YAML, edit it, and apply it back to a presentation copy. | - | Normal |
| A4 | - |  | - | BG:#FFFFFF |
| A5 | - | ## Scope and Assumptions | - | FG:#800000 |
| A6 | - |  | - | BG:#FFFFFF |
| A7 | - | - Primary use case is translation, not iterative slide editing. | - | Normal |
| A8 | - | - YAML keys are derived from slide structure, shape names, table coordinates, and paragraph order. | - | Normal |
| A9 | - | - If the source presentation structure changes, export YAML again before applying. | - | Normal |
| A10 | - | - Applying an old YAML file to a structurally edited deck may produce unexpected key mismatches. | - | Normal |
| A11 | - |  | - | BG:#FFFFFF |
| A12 | - | ## Workflow | - | FG:#800000 |
| A13 | - |  | - | BG:#FFFFFF |
| A14 | - | 1. Select source PowerPoint and export YAML. | - | Normal |
| A15 | - | 2. Edit YAML values. | - | Normal |
| A16 | - | 3. Select edited YAML and base PowerPoint. | - | Normal |
| A17 | - | 4. Apply YAML to create an output PowerPoint copy. | - | Normal |
| A18 | - |  | - | BG:#FFFFFF |
| A19 | - | The tool reports: | - | Normal |
| A20 | - |  | - | BG:#FFFFFF |
| A21 | - | - applied keys | - | Normal |
| A22 | - | - keys that exist only in PowerPoint | - | Normal |
| A23 | - | - keys that exist only in YAML | - | Normal |
| A24 | - | - malformed YAML lines | - | Normal |
| A25 | - |  | - | BG:#FFFFFF |
| A26 | - | ## YAML Key Model | - | FG:#800000 |
| A27 | - |  | - | BG:#FFFFFF |
| A28 | - | Key format is: | - | Normal |
| A29 | - |  | - | BG:#FFFFFF |
| A30 | - | ``` | - | Normal |
| A31 | - | sNN.shapePath.pNN | - | Normal |
| A32 | - | ``` | - | Normal |
| A33 | - |  | - | BG:#FFFFFF |
| A34 | - | - `sNN`: slide index | - | Normal |
| A35 | - | - `shapePath`: shape-name-based path (group hierarchy included) | - | Normal |
| A36 | - | - table cells are encoded as `.rNNcNN` | - | Normal |
| A37 | - | - `pNN`: paragraph index | - | Normal |
| A38 | - |  | - | BG:#FFFFFF |
| A39 | - | Duplicate sibling shape names are disambiguated with suffixes like `_2`, `_3`. | - | Normal |
| A40 | - |  | - | BG:#FFFFFF |
| A41 | - | ## Main Sheet Contract | - | FG:#800000 |
| A42 | - |  | - | BG:#FFFFFF |
| A43 | - | The Main sheet is intentionally address-based for transparent maintenance. | - | Normal |
| A44 | - |  | - | BG:#FFFFFF |
| A45 | - | - `B5`: source PowerPoint | - | Normal |
| A46 | - | - `B8`: exported YAML path | - | Normal |
| A47 | - | - `B11`: edited YAML path | - | Normal |
| A48 | - | - `B14`: base PowerPoint for apply | - | Normal |
| A49 | - | - `B17`: output PowerPoint path | - | Normal |
| A50 | - | - `B20`: output folder | - | Normal |
| A51 | - |  | - | BG:#FFFFFF |
| A52 | - | ## Main Sheet Protection Policy | - | FG:#800000 |
| A53 | - |  | - | BG:#FFFFFF |
| A54 | - | For release builds, the Main sheet is protected without a password to reduce accidental edits. | - | Normal |
| A55 | - | This protection is an operational safeguard and not a security boundary. | - | Normal |
| A56 | - |  | - | BG:#FFFFFF |
| A57 | - | - user-editable cells: `D1`, `B5`, `B11`, `B14` | - | Normal |
| A58 | - | - locked cells include formula/UI cells and internal output cells: `B8`, `B17`, `B20` | - | Normal |
| A59 | - | - on each open, `Workbook_Open` reapplies sheet protection with `UserInterfaceOnly:=True` | - | Normal |
| A60 | - |  | - | BG:#FFFFFF |
| A61 | - | Because `UserInterfaceOnly` is not persisted by Excel, protection settings are re-applied every time the workbook opens. | - | Normal |
| A62 | - | If macros are disabled, this runtime re-apply step will not run. | - | Normal |
| A63 | - |  | - | BG:#FFFFFF |
| A64 | - | ## Localization Design | - | FG:#800000 |
| A65 | - |  | - | BG:#FFFFFF |
| A66 | - | - UI labels are localized by worksheet formulas. | - | Normal |
| A67 | - | - VBA localizes only runtime dialog messages. | - | Normal |
| A68 | - | - VBA reads language state from `Sheet1!D1` and language column index from `Sheet1!F1`. | - | Normal |
| A69 | - | - Translation text is stored in the `i18n` worksheet. | - | Normal |
| A70 | - |  | - | BG:#FFFFFF |
| A71 | - | ## Environment Notes | - | FG:#800000 |
| A72 | - |  | - | BG:#FFFFFF |
| A73 | - | - Windows + Excel VBA environment is required. | - | Normal |
| A74 | - | - For module import/export macros, enable: | - | Normal |
| A75 | - |     - Trust access to the VBA project object model | - | Normal |
| A76 | - | - VBA component import/export uses system ANSI code page; this project converts between UTF-8 files and ANSI for VBE compatibility. | - | Normal |
