# Sheet Configuration
- VBA CodeName: Sheet5
- Excel UI Name: README.ja

| Address | Name | Value / Label | Formula | Style |
| :--- | :--- | :--- | :--- | :--- |
| A1 | - | # pptx_yaml_excel | - | FG:#800000; Bold |
| A2 | - |  | - | BG:#FFFFFF |
| A3 | - | PowerPoint のテキストを YAML に書き出し、編集してプレゼンテーションに適用しなおす Excel VBA ツールです。 | - | - |
| A4 | - |  | - | BG:#FFFFFF |
| A5 | - | ![スクリーンショット](capture.ja.png) | - | - |
| A6 | - |  | - | BG:#FFFFFF |
| A7 | - | ## ダウンロード | - | FG:#800000; Bold |
| A8 | - |  | - | BG:#FFFFFF |
| A9 | - | **利用者向け**: 以下のリンクから最新の `pptx_yaml_i18n.xlsm` をダウンロードしてください。インストール不要です。Excel で開いてマクロを有効にするだけで使えます。 | - | - |
| A10 | - |  | - | BG:#FFFFFF |
| A11 | - | [**pptx_yaml_i18n.xlsm をダウンロード（最新リリース）**](https://github.com/minipoisson/pptx_yaml_excel/releases/latest/download/pptx_yaml_i18n.xlsm) | - | - |
| A12 | - |  | - | BG:#FFFFFF |
| A13 | - | **開発者向け**: このリポジトリをそのままクローンして使えます。[xlsm_devkit](https://github.com/minipoisson/xlsm_devkit) を使った開発の実例として、`src/*.bas` と `sheet/*.md` をバージョン管理しています。 | - | - |
| A14 | - |  | - | BG:#FFFFFF |
| A15 | - | ## 用途と前提 | - | FG:#800000; Bold |
| A16 | - |  | - | BG:#FFFFFF |
| A17 | - | - 主な用途は翻訳作業であり、スライドを繰り返し編集するための差分管理ツールではありません。 | - | - |
| A18 | - | - YAML のキーはスライド構造・図形名・表の座標・段落順から生成されます。 | - | - |
| A19 | - | - 原本 PowerPoint の構造（図形の追加・削除・改名など）が変わった場合は、YAML を再抽出してください。 | - | - |
| A20 | - | - 構造変更後の古い YAML をそのまま適用すると、キーの不整合が多数発生する可能性があります。 | - | - |
| A21 | - |  | - | BG:#FFFFFF |
| A22 | - | ## ワークフロー | - | FG:#800000; Bold |
| A23 | - |  | - | BG:#FFFFFF |
| A24 | - | 1. 原本 PowerPoint を選択して YAML をエクスポートする。 | - | - |
| A25 | - | 2. YAML のテキスト値を編集する。 | - | - |
| A26 | - | 3. 編集済み YAML とベースとなる PowerPoint を選択する。 | - | - |
| A27 | - | 4. YAML を適用して出力 PowerPoint を生成する。 | - | - |
| A28 | - |  | - | BG:#FFFFFF |
| A29 | - | 実行後、次の件数と一覧が表示されます。 | - | - |
| A30 | - |  | - | BG:#FFFFFF |
| A31 | - | - 適用されたキー | - | - |
| A32 | - | - PowerPoint にしか存在しないキー | - | - |
| A33 | - | - YAML にしか存在しないキー | - | - |
| A34 | - | - 不正形式の YAML 行 | - | - |
| A35 | - |  | - | BG:#FFFFFF |
| A36 | - | ## YAML キーの仕様 | - | FG:#800000; Bold |
| A37 | - |  | - | BG:#FFFFFF |
| A38 | - | キーのフォーマット: | - | - |
| A39 | - |  | - | BG:#FFFFFF |
| A40 | - | ``` | - | - |
| A41 | - | sNN.shapePath.pNN | - | - |
| A42 | - | ``` | - | - |
| A43 | - |  | - | BG:#FFFFFF |
| A44 | - | - `sNN`: スライド番号（2 桁） | - | - |
| A45 | - | - `shapePath`: 図形名ベースのパス（グループ階層を `.` でつなぐ） | - | - |
| A46 | - | - 表のセルは `.rNNcNN`（行・列の 0 始まり 2 桁）で表現 | - | - |
| A47 | - | - `pNN`: 段落インデックス（0 始まり 2 桁） | - | - |
| A48 | - |  | - | BG:#FFFFFF |
| A49 | - | 同じスライド内で図形名が重複する場合、出現順に `_2`、`_3` … のサフィックスで区別されます。 | - | - |
| A50 | - |  | - | BG:#FFFFFF |
| A51 | - | ## Main シートのセル配置 | - | FG:#800000; Bold |
| A52 | - |  | - | BG:#FFFFFF |
| A53 | - | Main シートはアプリの UI として機能します。セル番地による運用は透明性・管理のしやすさを優先した意図的な設計です。 | - | - |
| A54 | - |  | - | BG:#FFFFFF |
| A55 | - | \| セル \| 内容 \| | - | - |
| A56 | - | \| :--- \| :--- \| | - | - |
| A57 | - | \| B5 \| 原本 PowerPoint のパス \| | - | - |
| A58 | - | \| B8 \| エクスポートされた YAML のパス \| | - | - |
| A59 | - | \| B11 \| 編集済み YAML のパス \| | - | - |
| A60 | - | \| B14 \| ベース PowerPoint のパス \| | - | - |
| A61 | - | \| B17 \| 出力 PowerPoint のパス \| | - | - |
| A62 | - | \| B20 \| 出力フォルダ \| | - | - |
| A63 | - |  | - | BG:#FFFFFF |
| A64 | - | ## Main シート保護ポリシー | - | FG:#800000; Bold |
| A65 | - |  | - | BG:#FFFFFF |
| A66 | - | 配布版では、Main シートを誤操作防止のためにパスワードなしで保護します。 | - | - |
| A67 | - | この保護は運用上の事故防止が目的であり、セキュリティ境界を提供するものではありません。 | - | - |
| A68 | - |  | - | BG:#FFFFFF |
| A69 | - | - ユーザー編集可能セル: `D1`, `B5`, `B11`, `B14` | - | - |
| A70 | - | - ロック対象セル: 数式/UI セル、および内部更新セル `B8`, `B17`, `B20` | - | - |
| A71 | - | - ブック起動時に `Workbook_Open` から `UserInterfaceOnly:=True` で保護を再適用します | - | - |
| A72 | - |  | - | BG:#FFFFFF |
| A73 | - | `UserInterfaceOnly` は Excel 保存時に永続化されないため、ブックを開くたびに再設定が必要です。 | - | - |
| A74 | - | マクロが無効の場合、この再適用処理は実行されません。 | - | - |
| A75 | - |  | - | BG:#FFFFFF |
| A76 | - | ## 多言語対応の仕組み | - | FG:#800000; Bold |
| A77 | - |  | - | BG:#FFFFFF |
| A78 | - | - Main シートの UI ラベルはシート上の数式によって多言語化されています。 | - | - |
| A79 | - | - VBA は実行時ダイアログのメッセージのみ多言語対応します。 | - | - |
| A80 | - | - VBA は `Sheet1!D1`（言語名）と `Sheet1!F1`（言語列インデックス）を参照します。F1 はシート数式が管理します。 | - | - |
| A81 | - | - 翻訳テキストは `i18n` シートに格納されています。 | - | - |
| A82 | - |  | - | BG:#FFFFFF |
| A83 | - | ## 開発環境 | - | FG:#800000; Bold |
| A84 | - |  | - | BG:#FFFFFF |
| A85 | - | このプロジェクトは [xlsm_devkit](https://github.com/minipoisson/xlsm_devkit) を使って開発しました。xlsm_devkit は Excel VBA プロジェクト向けの開発支援ツールキットです。 | - | - |
| A86 | - |  | - | BG:#FFFFFF |
| A87 | - | - `src/*.bas` — xlsm_devkit の `ExportAllModules`・`ImportAllModules` マクロで管理する VBA モジュール（UTF-8） | - | - |
| A88 | - | - `sheet/*.md` — xlsm_devkit の `ExportAllSheetMapsToMD` マクロで出力したシートマップ（Markdown） | - | - |
| A89 | - |  | - | BG:#FFFFFF |
| A90 | - | ### 前提条件 | - | FG:#800000; Bold |
| A91 | - |  | - | BG:#FFFFFF |
| A92 | - | - Windows + Microsoft Excel（VBA 有効） | - | - |
| A93 | - | - Excel のマクロ設定で **VBA プロジェクト オブジェクト モデルへのアクセスを信頼する** を有効にしてください: | - | - |
| A94 | - |     ``` | - | - |
| A95 | - |     ファイル → オプション → トラスト センター → トラスト センターの設定 | - | - |
| A96 | - |       → マクロの設定 → VBA プロジェクト オブジェクト モデルへのアクセスを信頼する | - | - |
| A97 | - |     ``` | - | - |
| A98 | - | - xlsm_devkit が UTF-8（ファイル）と ANSI（VBE）の相互変換を担うため、`src/` の `.bas` ファイルは常に UTF-8 で保存されます。 | - | - |
