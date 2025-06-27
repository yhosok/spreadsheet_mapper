/**
 * カスタムメニュー設定ファイル
 * スプレッドシートにカスタムメニューを追加
 */

/**
 * スプレッドシート開時に実行される関数
 * カスタムメニューを自動で追加
 */
function onOpen() {
  try {
    createCustomMenu();
    console.log('カスタムメニューが正常に作成されました');
  } catch (error) {
    console.error('カスタムメニュー作成エラー:', error);
    // エラーが発生してもスプレッドシートの読み込みは継続
  }
}

/**
 * カスタムメニューを作成
 */
function createCustomMenu() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('📊 フォーマット変換ツール')
    .addItem('🔄 フォーマット変換実行', 'executeFormatConversion')
    .addSeparator()
    .addItem('✅ マッピング検証', 'validateMappingSetup')
    .addItem('👀 変換プレビュー', 'showConversionPreview')
    .addSeparator()
    .addItem('📝 シート作成ヘルプ', 'showSheetCreationHelp')
    .addItem('❓ 使い方ガイド', 'showUsageGuide')
    .addItem('ℹ️ バージョン情報', 'showVersionInfo')
    .addToUi();
}

/**
 * マッピング設定の検証
 * マッピングデータとフォーマットA入力シートの整合性をチェック
 */
function validateMappingSetup() {
  try {
    console.log('マッピング検証を開始します');
    
    // 必要なシートの存在確認
    if (!validateSheets()) {
      return;
    }
    
    // マッピングデータの取得と検証
    const mappingData = getMappingData();
    if (!mappingData || mappingData.length === 0) {
      showError('項目マッピングシートにデータが設定されていません。');
      return;
    }
    
    // マッピングデータの妥当性チェック
    const mappingValidation = validateMappingData(mappingData);
    if (!mappingValidation.isValid) {
      const errorMessage = 'マッピングデータに問題があります:\n\n' + 
                          mappingValidation.errors.join('\n');
      showError(errorMessage);
      return;
    }
    
    // フォーマットA入力シートとの整合性チェック
    const inputData = getInputData();
    if (inputData && inputData.headers && inputData.headers.length > 0) {
      const headerValidation = validateHeaderMapping(inputData.headers, mappingData);
      
      let message = 'マッピング検証結果:\n\n';
      message += `✅ マッピング項目数: ${mappingData.length}件\n`;
      message += `✅ 入力シート列数: ${inputData.headers.length}件\n\n`;
      
      if (!headerValidation.isValid) {
        message += '⚠️ 以下の項目が入力シートに見つかりません:\n';
        message += headerValidation.missingColumns.map(col => `  - ${col}`).join('\n');
        message += '\n\n';
      }
      
      if (headerValidation.warnings.length > 0) {
        message += '📝 注意事項:\n';
        message += headerValidation.warnings.map(warning => `  - ${warning}`).join('\n');
        message += '\n\n';
      }
      
      if (headerValidation.isValid && headerValidation.warnings.length === 0) {
        message += '✅ すべての検証に合格しました！';
        showSuccess(message);
      } else {
        showWarning(message);
      }
    } else {
      showSuccess(`マッピングデータの検証に合格しました。\n項目数: ${mappingData.length}件\n\nフォーマットA入力シートにデータを追加してから再度検証することをお勧めします。`);
    }
    
  } catch (error) {
    console.error('マッピング検証エラー:', error);
    showError(`マッピング検証中にエラーが発生しました: ${error.message}`);
  }
}

/**
 * 変換プレビューの表示
 * 実際の変換処理を実行せずに結果をプレビュー
 */
function showConversionPreview() {
  try {
    console.log('変換プレビューを開始します');
    
    // 必要なシートの存在確認
    if (!validateSheets()) {
      return;
    }
    
    // マッピングデータの取得
    const mappingData = getMappingData();
    if (!mappingData || mappingData.length === 0) {
      showError('項目マッピングシートにデータが設定されていません。');
      return;
    }
    
    // フォーマットA入力データの取得
    const inputData = getInputData();
    if (!inputData || inputData.data.length === 0) {
      showError('フォーマットA入力シートにデータが存在しません。');
      return;
    }
    
    // プレビューデータの生成（最初の3行のみ）
    const preview = generateConversionPreview(inputData, mappingData, 3);
    
    // プレビュー結果の表示
    let message = '🔍 変換プレビュー結果:\n\n';
    message += `📊 入力データ: ${inputData.data.length}行\n`;
    message += `🗂️ マッピング項目: ${mappingData.length}件\n\n`;
    
    message += '📋 フォーマットB出力項目:\n';
    message += preview.outputHeaders.map((header, index) => `  ${index + 1}. ${header}`).join('\n');
    message += '\n\n';
    
    message += '📝 サンプルデータ（最初の3行）:\n';
    for (let i = 0; i < preview.sampleData.length; i++) {
      message += `\n行 ${i + 1}:\n`;
      for (let j = 0; j < preview.outputHeaders.length; j++) {
        const header = preview.outputHeaders[j];
        const value = preview.sampleData[i].output[j] || '(空)';
        message += `  ${header}: ${value}\n`;
      }
    }
    
    if (inputData.data.length > 3) {
      message += `\n... 他 ${inputData.data.length - 3} 行のデータがあります`;
    }
    
    showSuccess(message);
    
  } catch (error) {
    console.error('変換プレビューエラー:', error);
    showError(`変換プレビュー中にエラーが発生しました: ${error.message}`);
  }
}

/**
 * シート作成ヘルプの表示
 */
function showSheetCreationHelp() {
  const message = `📚 必要なシートの作成方法:\n\n` +
    `1️⃣ ${SHEET_NAMES.INPUT}:\n` +
    `   • シート名: 「${SHEET_NAMES.INPUT}」\n` +
    `   • 1行目: ヘッダー行（項目名）\n` +
    `   • 2行目以降: 変換したいデータ\n\n` +
    
    `2️⃣ ${SHEET_NAMES.MAPPING}:\n` +
    `   • シート名: 「${SHEET_NAMES.MAPPING}」\n` +
    `   • A列: フォーマットAの項目名\n` +
    `   • B列: フォーマットBの項目名\n\n` +
    
    `📋 マッピングシートの例（ヘッダーあり）:\n` +
    `   A1: フォーマットA項目名  B1: フォーマットB項目名\n` +
    `   A2: 出荷日             B2: shipping_date\n` +
    `   A3: 顧客名             B3: customer_name\n` +
    `   A4: 商品コード         B4: product_code\n` +
    `   A5: 数量              B5: quantity\n\n` +
    
    `📋 マッピングシートの例（ヘッダーなし）:\n` +
    `   A1: 出荷日             B1: shipping_date\n` +
    `   A2: 顧客名             B2: customer_name\n` +
    `   A3: 商品コード         B3: product_code\n` +
    `   A4: 数量              B4: quantity\n\n` +
    
    `⚠️ 注意事項:\n` +
    `   • シート名は正確に入力してください\n` +
    `   • フォーマットB項目名は英数字とアンダースコアのみ使用\n` +
    `   • 重複する項目名は設定しないでください\n` +
    `   • ヘッダー行は自動検出されます`;
  
  showSuccess(message);
}

/**
 * 使い方ガイドの表示
 */
function showUsageGuide() {
  const message = `📖 使い方ガイド:\n\n` +
    `1️⃣ 事前準備:\n` +
    `   • 必要なシートを作成（シート作成ヘルプ参照）\n` +
    `   • 項目マッピングシートにマッピング情報を設定\n\n` +
    
    `2️⃣ データ変換手順:\n` +
    `   • フォーマットA入力シートにデータを貼り付け\n` +
    `   • 「マッピング検証」で設定を確認\n` +
    `   • 「変換プレビュー」で結果を確認\n` +
    `   • 「フォーマット変換実行」で変換実行\n\n` +
    
    `3️⃣ 結果確認:\n` +
    `   • フォーマットB出力シートが自動生成されます\n` +
    `   • 変換されたデータを確認してください\n\n` +
    
    `🔧 トラブルシューティング:\n` +
    `   • エラーが出る場合は「マッピング検証」を実行\n` +
    `   • シート名が正確か確認\n` +
    `   • データの形式が正しいか確認`;
  
  showSuccess(message);
}

/**
 * バージョン情報の表示
 */
function showVersionInfo() {
  const version = '1.0.0';
  const releaseDate = '2024-01-01';
  
  const message = `📱 フォーマット変換ツール\n\n` +
    `バージョン: ${version}\n` +
    `リリース日: ${releaseDate}\n\n` +
    `🎯 主な機能:\n` +
    `   • CSVフォーマット変換\n` +
    `   • 項目マッピング設定\n` +
    `   • データ検証とプレビュー\n` +
    `   • エラーハンドリング\n\n` +
    `💻 開発者: フォーマット変換チーム\n` +
    `📧 サポート: 管理者までお問い合わせください`;
  
  showSuccess(message);
}
