# Sheet Configuration
- VBA CodeName: Sheet1
- Excel UI Name: Main

| Address | Name | Value / Label | Formula | Style |
| :--- | :--- | :--- | :--- | :--- |
| A1 | - | PowerPoint-YAML Conversion Tool | `=INDEX(i18n!$A$1:$CZ$503,MATCH(MID(CELL("filename",A1),FIND("]",CELL("filename",A1))+1,99)&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | FontSize:16 |
| C1 | - | 🌐: | - | Normal |
| D1 | - | English | - | BG:#FFF2CC; List:=i18n!$B$2:$CZ$2 |
| E1 | - | column no.: | `=INDEX(i18n!$A$1:$CZ$503,MATCH(MID(CELL("filename",E1),FIND("]",CELL("filename",E1))+1,99)&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | Normal |
| F1 | - | 2 | `=MATCH(D1,i18n!2:2,0)` | Normal |
| A2 | - | Quick steps: Select source PowerPoint -> extract text to YAML -> select edited YAML -> apply it to create a new PowerPoint | `=INDEX(i18n!$A$1:$CZ$503,MATCH(MID(CELL("filename",A2),FIND("]",CELL("filename",A2))+1,99)&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | Normal |
| B3 | - | Clear All Below | `=INDEX(i18n!$A$1:$CZ$503,MATCH(MID(CELL("filename",B3),FIND("]",CELL("filename",B3))+1,99)&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | Normal |
| B4 | - | Please enter the path below. | `=INDEX(i18n!$A$1:$CZ$503,MATCH(MID(CELL("filename",B4),FIND("]",CELL("filename",B4))+1,99)&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | Normal |
| A5 | - | 1. Source PowerPoint: | `=INDEX(i18n!$A$1:$CZ$503,MATCH(MID(CELL("filename",A5),FIND("]",CELL("filename",A5))+1,99)&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | BG:#FFD966 |
| B5 | - |  | - | BG:#FFF2CC |
| C5 | - | Browse | `=INDEX(i18n!$A$1:$CZ$503,MATCH(MID(CELL("filename",C5),FIND("]",CELL("filename",C5))+1,99)&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | Normal |
| B7 | - | It will use the same name as the source PowerPoint. | `=INDEX(i18n!$A$1:$CZ$503,MATCH(MID(CELL("filename",B7),FIND("]",CELL("filename",B7))+1,99)&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | Normal |
| A8 | - | 2. Extracted YAML: | `=INDEX(i18n!$A$1:$CZ$503,MATCH(MID(CELL("filename",A8),FIND("]",CELL("filename",A8))+1,99)&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | BG:#FFD966 |
| B8 | - |  | - | BG:#D9D9D9 |
| B10 | - | Please enter the path below. | `=INDEX(i18n!$A$1:$CZ$503,MATCH(MID(CELL("filename",B10),FIND("]",CELL("filename",B10))+1,99)&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | Normal |
| A11 | - | 3. Edited YAML: | `=INDEX(i18n!$A$1:$CZ$503,MATCH(MID(CELL("filename",A11),FIND("]",CELL("filename",A11))+1,99)&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | BG:#FFD966 |
| B11 | - |  | - | BG:#FFF2CC |
| C11 | - | Browse | `=INDEX(i18n!$A$1:$CZ$503,MATCH(MID(CELL("filename",C11),FIND("]",CELL("filename",C11))+1,99)&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | Normal |
| B13 | - | Please enter the path below. | `=INDEX(i18n!$A$1:$CZ$503,MATCH(MID(CELL("filename",B13),FIND("]",CELL("filename",B13))+1,99)&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | Normal |
| A14 | - | 4. Base PowerPoint: | `=INDEX(i18n!$A$1:$CZ$503,MATCH(MID(CELL("filename",A14),FIND("]",CELL("filename",A14))+1,99)&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | BG:#FFD966 |
| B14 | - |  | - | BG:#FFF2CC |
| C14 | - | Browse | `=INDEX(i18n!$A$1:$CZ$503,MATCH(MID(CELL("filename",C14),FIND("]",CELL("filename",C14))+1,99)&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | Normal |
| D14 | - | Same as Source | `=INDEX(i18n!$A$1:$CZ$503,MATCH(MID(CELL("filename",D14),FIND("]",CELL("filename",D14))+1,99)&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | Normal |
| B16 | - | It will use the same name as the edited YAML. | `=INDEX(i18n!$A$1:$CZ$503,MATCH(MID(CELL("filename",B16),FIND("]",CELL("filename",B16))+1,99)&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | Normal |
| A17 | - | 5. Output PowerPoint: | `=INDEX(i18n!$A$1:$CZ$503,MATCH(MID(CELL("filename",A17),FIND("]",CELL("filename",A17))+1,99)&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | BG:#FFD966 |
| B17 | - |  | - | BG:#D9D9D9 |
| B19 | - | After creation, check the folder below | `=INDEX(i18n!$A$1:$CZ$503,MATCH(MID(CELL("filename",B19),FIND("]",CELL("filename",B19))+1,99)&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | Normal |
| A20 | - | 6. Output Folder | `=INDEX(i18n!$A$1:$CZ$503,MATCH(MID(CELL("filename",A20),FIND("]",CELL("filename",A20))+1,99)&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | BG:#FFD966 |
| B20 | - |  | - | BG:#D9D9D9 |
| B21 | - | After completion, the output folder opens and the created file is selected | `=INDEX(i18n!$A$1:$CZ$503,MATCH(MID(CELL("filename",B21),FIND("]",CELL("filename",B21))+1,99)&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | Normal |
| B22 | - | Create Output PowerPoint | `=INDEX(i18n!$A$1:$CZ$503,MATCH(MID(CELL("filename",B22),FIND("]",CELL("filename",B22))+1,99)&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | Normal |

## Shapes

| Address | Name | Label | Formula | OnAction | Style |
| :--- | :--- | :--- | :--- | :--- | :--- |
| C5 | Bevel 1 | Browse | `=C5` | pptx_yaml_i18n.xlsm!Sheet1.btnSpecifyOrgPpt_Click | BG:#D2D2D2; FG:#000000 |
| C11 | Bevel 2 | Browse | `=C11` | pptx_yaml_i18n.xlsm!Sheet1.btnSpecifyYaml_Click | BG:#D2D2D2; FG:#000000 |
| C14 | Bevel 3 | Browse | `=C14` | pptx_yaml_i18n.xlsm!Sheet1.btnSpecifyBasePpt_Click | BG:#D2D2D2; FG:#000000 |
| D14 | Bevel 4 | Same as Source | `=D14` | pptx_yaml_i18n.xlsm!Sheet1.btnCopyOrgToBase_Click | BG:#D2D2D2; FG:#000000 |
| B22 | Bevel 5 | Create Output PowerPoint | `=B22` | pptx_yaml_i18n.xlsm!Sheet1.btnApplyYamlToPpt_Click | BG:#D2D2D2; FG:#000000 |
| B3 | Bevel 6 | Clear All Below | `=B3` | pptx_yaml_i18n.xlsm!Sheet1.btnClearAll_Click | BG:#B1CBE9; FG:#000000 |
