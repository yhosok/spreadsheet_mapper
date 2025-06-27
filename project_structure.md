# プロジェクト構成

```
spreadsheet_mapper/
├── README.md                 # プロジェクト概要と使用方法
├── src/                      # Google Apps Scriptソースコード
│   ├── main.gs              # メイン処理（変換実行）
│   ├── mapping.gs           # マッピング処理
│   ├── utils.gs             # ユーティリティ関数
│   └── menu.gs              # カスタムメニュー設定
├── sample/                   # サンプルデータ
│   └── sample_data.csv      # テスト用サンプルデータ
├── docs/                     # ドキュメント
│   └── setup_guide.md       # 詳細セットアップガイド
└── project_structure.md     # このファイル
```

## ファイル詳細

### src/main.gs
- `executeFormatConversion()` - メイン変換処理
- `validateSheets()` - 必要シートの存在確認
- `getInputData()` - フォーマットA入力データ取得
- `prepareOutputSheet()` - フォーマットB出力シート準備
- `convertData()` - データ変換処理

### src/mapping.gs
- `getMappingData()` - 項目マッピング取得
- `validateMappingData()` - マッピングデータ検証
- `validateHeaderMapping()` - ヘッダーとマッピングの整合性チェック
- `generateConversionPreview()` - 変換プレビュー生成

### src/utils.gs
- `showSuccess()`, `showError()`, `showWarning()` - メッセージ表示
- `isEmptyOrWhitespace()`, `isEmptyArray()` - 値チェック
- `getCurrentDateTime()` - 日時取得
- `Timer` - パフォーマンス測定クラス

### src/menu.gs
- `onOpen()` - スプレッドシート開時処理
- `createCustomMenu()` - カスタムメニュー作成
- `validateMappingSetup()` - マッピング検証
- `showConversionPreview()` - 変換プレビュー表示
- ヘルプ機能各種
