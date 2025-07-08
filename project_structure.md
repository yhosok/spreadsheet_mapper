# プロジェクト構成（拡張版対応）

```
spreadsheet_mapper/
├── README.md                          # プロジェクト概要と使用方法（拡張版対応）
├── all-in-one.gs                     # 統合版スクリプト（基本版）
├── all-in-one-extended.gs           # 統合版スクリプト（拡張版）
├── src/                              # Google Apps Scriptソースコード
│   ├── main.gs                       # メイン処理（拡張版対応）
│   ├── mapping.gs                    # マッピング処理（拡張版対応）
│   ├── utils.gs                      # ユーティリティ関数
│   └── menu.gs                       # カスタムメニュー設定
├── sample/                           # サンプルデータ
│   ├── sample_data.csv               # テスト用サンプルデータ（基本版）
│   ├── sample_data_extended.csv      # テスト用サンプルデータ（拡張版）
│   └── mapping_example_extended.csv  # マッピング例（拡張版）
├── docs/                             # ドキュメント
│   ├── setup_guide.md                # 詳細セットアップガイド
│   └── extended_features_guide.md    # 拡張機能使用ガイド
└── project_structure.md              # このファイル
```

## 📊 バージョン情報

### v1.0.0（基本版）
- 1対1のフィールドマッピング
- 2列形式のマッピングシート
- 基本的なフォーマット変換機能

### v2.0.0（拡張版）
- 複数フィールド結合機能（CONCAT）
- 固定値設定機能（FIXED）
- 4列形式のマッピングシート対応
- 拡張機能の詳細ガイド

## ファイル詳細

### 統合版ファイル

#### all-in-one.gs（基本版）
- 基本的なフォーマット変換機能
- 2列形式のマッピングシート対応
- 1対1のフィールドマッピングのみ

#### all-in-one-extended.gs（拡張版）
- 全機能を含む統合スクリプト
- 複数フィールド結合・固定値設定対応
- 4列形式のマッピングシート対応

### コアファイル

#### src/main.gs（拡張版対応）
- `executeFormatConversion()` - メイン変換処理
- `validateSheets()` - 必要シートの存在確認
- `getInputData()` - フォーマットA入力データ取得
- `prepareOutputSheet()` - フォーマットB出力シート準備
- `convertData()` - データ変換処理（拡張機能対応）

#### src/mapping.gs（拡張版対応）
- `getMappingData()` - 項目マッピング取得（4列対応）
- `parseMappingRow()` - マッピング行解析（新機能）
- `validateMappingData()` - マッピングデータ検証（拡張対応）
- `validateHeaderMapping()` - ヘッダーとマッピングの整合性チェック
- `generateConversionPreview()` - 変換プレビュー生成
- `convertFieldValue()` - 単一フィールド値変換（新機能）

#### src/utils.gs
- `showSuccess()`, `showError()`, `showWarning()` - メッセージ表示
- `getCurrentDateTime()` - 日時取得
- `Timer` - パフォーマンス測定クラス

#### src/menu.gs
- `onOpen()` - スプレッドシート開時処理
- `createCustomMenu()` - カスタムメニュー作成
- `validateMappingSetup()` - マッピング検証
- `showConversionPreview()` - 変換プレビュー表示
- `showSheetCreationHelp()` - シート作成ヘルプ（拡張機能対応）
- `showUsageGuide()` - 使い方ガイド（拡張機能対応）
- `showVersionInfo()` - バージョン情報（v2.0.0対応）

### サンプルファイル

#### sample/sample_data.csv（基本版）
- 基本的なフィールド構成
- シンプルなサンプルデータ

#### sample/sample_data_extended.csv（拡張版）
- 住所フィールドが分割されたデータ
- 拡張機能のテストに適したデータ

#### sample/mapping_example_extended.csv（拡張版）
- 4列形式のマッピング例
- 結合・固定値設定の使用例

### ドキュメント

#### docs/setup_guide.md
- 詳細なセットアップ手順
- 基本版・拡張版の両方に対応

#### docs/extended_features_guide.md
- 拡張機能の詳細ガイド
- 実践的な使用例
- エラー対応・トラブルシューティング

## 🎯 新機能詳細

### 1. 複数フィールド結合（CONCAT）
- 複数のINPUTフィールドを1つのOUTPUTフィールドに結合
- 区切り文字のカスタマイズ可能
- 空の値は自動的に除外

### 2. 固定値設定（FIXED）
- 特定フィールドに固定値を設定
- 動的な値の設定も可能
- 全行に同じ値を適用

### 3. 4列マッピング形式
- A列: フォーマットA項目名（または特殊形式）
- B列: フォーマットB項目名
- C列: 変換タイプ（NORMAL, CONCAT, FIXED）
- D列: 設定値（区切り文字、固定値など）

### 4. 自動タイプ判定
- マッピング内容から変換タイプを自動判定
- 明示的な指定も可能
- 後方互換性の確保
