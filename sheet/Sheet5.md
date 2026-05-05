# Sheet Configuration
- VBA CodeName: Sheet5
- Excel UI Name: README.ja

| Address | Name | Value / Label | Formula | Style |
| :--- | :--- | :--- | :--- | :--- |
| A1 | - | # pptx_yaml_excel | - | FG:#800000 |
| A2 | - |  | - | BG:#FFFFFF |
| A3 | - | PowerPoint のテキストを YAML に書き出し、編集してプレゼンテーションに適用しなおす Excel VBA ツールです。 | - | Normal |
| A4 | - |  | - | BG:#FFFFFF |
| A5 | - | ## 用途と前提 | - | FG:#800000 |
| A6 | - |  | - | BG:#FFFFFF |
| A7 | - | - 主な用途は翻訳作業であり、スライドを繰り返し編集するための差分管理ツールではありません。 | - | Normal |
| A8 | - | - YAML のキーはスライド構造・図形名・表の座標・段落順から生成されます。 | - | Normal |
| A9 | - | - 原本 PowerPoint の構造（図形の追加・削除・改名など）が変わった場合は、YAML を再抽出してください。 | - | Normal |
| A10 | - | - 構造変更後の古い YAML をそのまま適用すると、キーの不整合が多数発生する可能性があります。 | - | Normal |
| A11 | - |  | - | BG:#FFFFFF |
| A12 | - | ## ワークフロー | - | FG:#800000 |
| A13 | - |  | - | BG:#FFFFFF |
| A14 | - | 1. 原本 PowerPoint を選択して YAML をエクスポートする。 | - | Normal |
| A15 | - | 2. YAML のテキスト値を編集する。 | - | Normal |
| A16 | - | 3. 編集済み YAML とベースとなる PowerPoint を選択する。 | - | Normal |
| A17 | - | 4. YAML を適用して出力 PowerPoint を生成する。 | - | Normal |
| A18 | - |  | - | BG:#FFFFFF |
| A19 | - | 実行後、次の件数と一覧が表示されます。 | - | Normal |
| A20 | - |  | - | BG:#FFFFFF |
| A21 | - | - 適用されたキー | - | Normal |
| A22 | - | - PowerPoint にしか存在しないキー | - | Normal |
| A23 | - | - YAML にしか存在しないキー | - | Normal |
| A24 | - | - 不正形式の YAML 行 | - | Normal |
| A25 | - |  | - | BG:#FFFFFF |
| A26 | - | ## YAML キーの仕様 | - | FG:#800000 |
| A27 | - |  | - | BG:#FFFFFF |
| A28 | - | キーのフォーマット: | - | Normal |
| A29 | - |  | - | BG:#FFFFFF |
| A30 | - | ``` | - | Normal |
| A31 | - | sNN.shapePath.pNN | - | Normal |
| A32 | - | ``` | - | Normal |
| A33 | - |  | - | BG:#FFFFFF |
| A34 | - | - `sNN`: スライド番号（2 桁） | - | Normal |
| A35 | - | - `shapePath`: 図形名ベースのパス（グループ階層を `.` でつなぐ） | - | Normal |
| A36 | - | - 表のセルは `.rNNcNN`（行・列の 0 始まり 2 桁）で表現 | - | Normal |
| A37 | - | - `pNN`: 段落インデックス（0 始まり 2 桁） | - | Normal |
| A38 | - |  | - | BG:#FFFFFF |
| A39 | - | 同じスライド内で図形名が重複する場合、出現順に `_2`、`_3` … のサフィックスで区別されます。 | - | Normal |
| A40 | - |  | - | BG:#FFFFFF |
| A41 | - | ## Main シートのセル配置 | - | FG:#800000 |
| A42 | - |  | - | BG:#FFFFFF |
| A43 | - | Main シートはアプリの UI として機能します。セル番地による運用は透明性・管理のしやすさを優先した意図的な設計です。 | - | Normal |
| A44 | - |  | - | BG:#FFFFFF |
| A45 | - | ｜ セル ｜ 内容 ｜ | - | Normal |
| A46 | - | ｜ :--- ｜ :--- ｜ | - | Normal |
| A47 | - | ｜ B5 ｜ 原本 PowerPoint のパス ｜ | - | Normal |
| A48 | - | ｜ B8 ｜ エクスポートされた YAML のパス ｜ | - | Normal |
| A49 | - | ｜ B11 ｜ 編集済み YAML のパス ｜ | - | Normal |
| A50 | - | ｜ B14 ｜ ベース PowerPoint のパス ｜ | - | Normal |
| A51 | - | ｜ B17 ｜ 出力 PowerPoint のパス ｜ | - | Normal |
| A52 | - | ｜ B20 ｜ 出力フォルダ ｜ | - | Normal |
| A53 | - |  | - | BG:#FFFFFF |
| A54 | - | ## Main シート保護ポリシー | - | FG:#800000 |
| A55 | - |  | - | BG:#FFFFFF |
| A56 | - | 配布版では、Main シートを誤操作防止のためにパスワードなしで保護します。 | - | Normal |
| A57 | - | この保護は運用上の事故防止が目的であり、セキュリティ境界を提供するものではありません。 | - | Normal |
| A58 | - |  | - | BG:#FFFFFF |
| A59 | - | - ユーザー編集可能セル: `D1`, `B5`, `B11`, `B14` | - | Normal |
| A60 | - | - ロック対象セル: 数式/UI セル、および内部更新セル `B8`, `B17`, `B20` | - | Normal |
| A61 | - | - ブック起動時に `Workbook_Open` から `UserInterfaceOnly:=True` で保護を再適用します | - | Normal |
| A62 | - |  | - | BG:#FFFFFF |
| A63 | - | `UserInterfaceOnly` は Excel 保存時に永続化されないため、ブックを開くたびに再設定が必要です。 | - | Normal |
| A64 | - | マクロが無効の場合、この再適用処理は実行されません。 | - | Normal |
| A65 | - |  | - | BG:#FFFFFF |
| A66 | - | ## 多言語対応の仕組み | - | FG:#800000 |
| A67 | - |  | - | BG:#FFFFFF |
| A68 | - | - Main シートの UI ラベルはシート上の数式によって多言語化されています。 | - | Normal |
| A69 | - | - VBA は実行時ダイアログのメッセージのみ多言語対応します。 | - | Normal |
| A70 | - | - VBA は `Sheet1!D1`（言語名）と `Sheet1!F1`（言語列インデックス）を参照します。F1 はシート数式が管理します。 | - | Normal |
| A71 | - | - 翻訳テキストは `i18n` シートに格納されています。 | - | Normal |
| A72 | - |  | - | BG:#FFFFFF |
| A73 | - | ## 実行環境 | - | FG:#800000 |
| A74 | - |  | - | BG:#FFFFFF |
| A75 | - | - Windows + Excel VBA 環境が必要です。 | - | Normal |
| A76 | - | - モジュールのインポート・エクスポートマクロを使用する場合は、Excel のマクロ設定で次を有効にしてください: | - | Normal |
| A77 | - |   - VBA プロジェクト オブジェクト モデルへのアクセスを信頼する | - | Normal |
| A78 | - | - VBA のモジュール入出力はシステムの ANSI コードページを使用します。このプロジェクトでは UTF-8（ファイル）と ANSI（VBE）の相互変換を行っています。 | - | Normal |
