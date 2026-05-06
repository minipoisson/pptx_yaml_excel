# Sheet Configuration
- VBA CodeName: Sheet1
- Excel UI Name: Main

| Address | Name | Value / Label | Formula | Style |
| :--- | :--- | :--- | :--- | :--- |
| A1 | - | PowerPoint⇔YAML変換ツール | `=INDEX(i18n!$A$1:$CZ$500,MATCH($G$1&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | FontSize:16 |
| C1 | - | 🌐: | - | Normal |
| D1 | - | 日本語 | - | BG:#FFF2CC; List:=i18n!$B$2:$CZ$2 |
| E1 | - | 列番号: | `=INDEX(i18n!$A$1:$CZ$500,MATCH($G$1&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | Normal |
| F1 | - | 11 | `=MATCH(D1,i18n!2:2,0)` | Normal |
| G1 | - | Main | `=MID(CELL("filename",G1),FIND("]",CELL("filename",G1))+1,99)` | Normal |
| A2 | - | 簡単な使い方：原本とするPowerPointを指定→テキストを抽出してYAMLに保存→そのYAMLを編集したものを指定→それを反映させて新たなPowerPointを作成 | `=INDEX(i18n!$A$1:$CZ$500,MATCH($G$1&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | Normal |
| B3 | - | 以下をすべてクリアする | `=INDEX(i18n!$A$1:$CZ$500,MATCH(MID(CELL("filename",B3),FIND("]",CELL("filename",B3))+1,99)&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | Normal |
| B4 | - | ↓パスをご入力ください↓ | `=INDEX(i18n!$A$1:$CZ$500,MATCH($G$1&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | Normal |
| A5 | - | 1.原本とするPowerPoint: | `=INDEX(i18n!$A$1:$CZ$500,MATCH($G$1&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | BG:#FFD966 |
| B5 | - |  | - | BG:#FFF2CC |
| C5 | - | 参照 | `=INDEX(i18n!$A$1:$CZ$500,MATCH($G$1&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | Normal |
| B7 | - | 原本PowerPointと同じ名前になります。 | `=INDEX(i18n!$A$1:$CZ$500,MATCH($G$1&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | Normal |
| A8 | - | 2.抽出されたYAML: | `=INDEX(i18n!$A$1:$CZ$500,MATCH($G$1&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | BG:#FFD966 |
| B8 | - |  | - | BG:#D9D9D9 |
| B10 | - | ↓パスをご入力ください↓ | `=INDEX(i18n!$A$1:$CZ$500,MATCH($G$1&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | Normal |
| A11 | - | 3.編集済みのYAML: | `=INDEX(i18n!$A$1:$CZ$500,MATCH($G$1&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | BG:#FFD966 |
| B11 | - |  | - | BG:#FFF2CC |
| C11 | - | 参照 | `=INDEX(i18n!$A$1:$CZ$500,MATCH($G$1&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | Normal |
| B13 | - | ↓パスをご入力ください↓ | `=INDEX(i18n!$A$1:$CZ$500,MATCH($G$1&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | Normal |
| A14 | - | 4.ベースにするPowerPoint: | `=INDEX(i18n!$A$1:$CZ$500,MATCH($G$1&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | BG:#FFD966 |
| B14 | - |  | - | BG:#FFF2CC |
| C14 | - | 参照 | `=INDEX(i18n!$A$1:$CZ$500,MATCH($G$1&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | Normal |
| D14 | - | 原本と同じ | `=INDEX(i18n!$A$1:$CZ$500,MATCH($G$1&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | Normal |
| B16 | - | 編集済みYAMLと同じ名前になります。 | `=INDEX(i18n!$A$1:$CZ$500,MATCH($G$1&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | Normal |
| A17 | - | 5.反映済みのPowerPoint: | `=INDEX(i18n!$A$1:$CZ$500,MATCH($G$1&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | BG:#FFD966 |
| B17 | - |  | - | BG:#D9D9D9 |
| B19 | - | 作成後、↓このフォルダ↓をご確認ください | `=INDEX(i18n!$A$1:$CZ$500,MATCH($G$1&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | Normal |
| A20 | - | 6.フォルダに出力 | `=INDEX(i18n!$A$1:$CZ$500,MATCH($G$1&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | BG:#FFD966 |
| B20 | - |  | - | BG:#D9D9D9 |
| B21 | - | 出力完了後、フォルダが開いて出力されたファイルが選択されます | `=INDEX(i18n!$A$1:$CZ$500,MATCH($G$1&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | Normal |
| B22 | - | 反映済みPowerPointを作成する | `=INDEX(i18n!$A$1:$CZ$500,MATCH($G$1&"!"&ADDRESS(ROW(),COLUMN(),4),i18n!$A:$A,0),$F$1)` | Normal |

## Shapes

| Address | Name | Label | Formula | OnAction | Style |
| :--- | :--- | :--- | :--- | :--- | :--- |
| C5 | Bevel 1 | 参照 | `=C5` | pptx_yaml_i18n.xlsm!Sheet1.btnSpecifyOrgPpt_Click | BG:#D2D2D2; FG:#000000 |
| C11 | Bevel 2 | 参照 | `=C11` | pptx_yaml_i18n.xlsm!Sheet1.btnSpecifyYaml_Click | BG:#D2D2D2; FG:#000000 |
| C14 | Bevel 3 | 参照 | `=C14` | pptx_yaml_i18n.xlsm!Sheet1.btnSpecifyBasePpt_Click | BG:#D2D2D2; FG:#000000 |
| D14 | Bevel 4 | 原本と同じ | `=D14` | pptx_yaml_i18n.xlsm!Sheet1.btnCopyOrgToBase_Click | BG:#D2D2D2; FG:#000000 |
| B22 | Bevel 5 | 反映済みPowerPointを作成する | `=B22` | pptx_yaml_i18n.xlsm!Sheet1.btnApplyYamlToPpt_Click | BG:#D2D2D2; FG:#000000 |
| B3 | Bevel 6 | 以下をすべてクリアする | `=B3` | pptx_yaml_i18n.xlsm!Sheet1.btnClearAll_Click | BG:#B1CBE9; FG:#000000 |
