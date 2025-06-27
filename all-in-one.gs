/**
 * Googleスプレッドシート フォーマット変換ツール - 統合版
 * 
 * このファイル1つで全ての機能を利用できます。
 * Google Apps Scriptエディタにこのコードをコピー&ペーストしてください。
 * 
 * 使用方法:
 * 1. Googleスプレッドシートで「拡張機能」→「Apps Script」を選択
 * 2. このファイルの内容を全てコピーしてApps Scriptエディタに貼り付け
 * 3. ファイルを保存
 * 4. スプレッドシートに戻ると「📊 フォーマット変換ツール」メニューが追加されます
 */

// ===============================
// 設定（カスタマイズ可能）
// ===============================

// シート名の設定
const SHEET_NAMES = {
  INPUT: 'INPUT',      // 入力シート名
  MAPPING: '項目マッピング',       // マッピングシート名
  OUTPUT: 'OUTPUT'      // 出力シート名
};

// マッピングシートの設定
const MAPPING_CONFIG = {
  HAS_HEADER: true     // 項目マッピングシートには必ずヘッダー行があります
};

// ===============================
// メイン処理
// ===============================

/**
 * フォーマット変換のメイン関数
 * カスタムメニューから呼び出される
 */
function executeFormatConversion() {
  try {
    // ログ開始
    console.log('フォーマット変換処理を開始します');
    
    // 必要なシートの存在確認
    if (!validateSheets()) {
      return;
    }
    
    // マッピング情報の取得
    const mappingData = getMappingData();
    if (!mappingData || mappingData.length === 0) {
      showError('項目マッピングシートにデータが設定されていません。');
      return;
    }
    
    // フォーマットA入力データの取得
    const inputData = getInputData();
    if (!inputData || inputData.length === 0) {
      showError('フォーマットA入力シートにデータが存在しません。');
      return;
    }
    
    // フォーマットB出力シートの準備
    const outputSheet = prepareOutputSheet();
    if (!outputSheet) {
      showError('フォーマットB出力シートの準備に失敗しました。');
      return;
    }
    
    // ヘッダー行の作成
    const headerRow = createHeaderRow(mappingData);
    outputSheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
    
    // データ変換と出力
    const convertedData = convertData(inputData, mappingData);
    if (convertedData.length > 0) {
      outputSheet.getRange(2, 1, convertedData.length, convertedData[0].length).setValues(convertedData);
    }
    
    // 完了メッセージ
    showSuccess(`変換が完了しました。${convertedData.length}行のデータを変換しました。`);
    console.log('フォーマット変換処理が正常に完了しました');
    
  } catch (error) {
    console.error('エラーが発生しました:', error);
    showError(`処理中にエラーが発生しました: ${error.message}`);
  }
}

/**
 * 必要なシートの存在確認
 * @return {boolean} 全ての必要シートが存在する場合true
 */
function validateSheets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const requiredSheets = [SHEET_NAMES.INPUT, SHEET_NAMES.MAPPING];
  
  for (const sheetName of requiredSheets) {
    if (!spreadsheet.getSheetByName(sheetName)) {
      showError(`必要なシート「${sheetName}」が存在しません。シートを作成してください。`);
      return false;
    }
  }
  
  return true;
}

/**
 * フォーマットA入力データの取得
 * @return {Object} 入力データ（headers, data）
 */
function getInputData() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = spreadsheet.getSheetByName(SHEET_NAMES.INPUT);
  
  const lastRow = inputSheet.getLastRow();
  const lastCol = inputSheet.getLastColumn();
  
  if (lastRow <= 1) {
    return null; // ヘッダー行のみまたはデータなし
  }
  
  // ヘッダー行とデータ行を取得
  const headerRange = inputSheet.getRange(1, 1, 1, lastCol);
  const dataRange = inputSheet.getRange(2, 1, lastRow - 1, lastCol);
  
  const headers = headerRange.getValues()[0];
  const data = dataRange.getValues();
  
  // ヘッダー情報を含むオブジェクト形式で返す
  return {
    headers: headers,
    data: data
  };
}

/**
 * フォーマットB出力シートの準備
 * 既存シートがある場合はクリア、なければ新規作成
 * @return {GoogleAppsScript.Spreadsheet.Sheet} 準備されたシート
 */
function prepareOutputSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = SHEET_NAMES.OUTPUT;
  
  let outputSheet = spreadsheet.getSheetByName(sheetName);
  
  if (outputSheet) {
    // 既存シートの内容をクリア
    outputSheet.clear();
  } else {
    // 新規シート作成
    outputSheet = spreadsheet.insertSheet(sheetName);
  }
  
  return outputSheet;
}

/**
 * ヘッダー行の作成
 * @param {Array<Array>} mappingData マッピングデータ
 * @return {Array} フォーマットBのヘッダー配列
 */
function createHeaderRow(mappingData) {
  return mappingData.map(mapping => mapping[1]); // B列（フォーマットB項目名）
}

/**
 * データの変換
 * @param {Object} inputData 入力データ（headers, data）
 * @param {Array<Array>} mappingData マッピングデータ
 * @return {Array<Array>} 変換されたデータ
 */
function convertData(inputData, mappingData) {
  const { headers, data } = inputData;
  const convertedData = [];
  
  for (const row of data) {
    const convertedRow = [];
    
    for (const mapping of mappingData) {
      const formatAColumn = mapping[0]; // A列（フォーマットA項目名）
      const columnIndex = headers.indexOf(formatAColumn);
      
      if (columnIndex !== -1) {
        // 対応する列が見つかった場合、その値を使用
        convertedRow.push(row[columnIndex] || '');
      } else {
        // 対応する列が見つからない場合、空文字を設定
        convertedRow.push('');
        console.warn(`警告: フォーマットA項目「${formatAColumn}」が入力データに見つかりません`);
      }
    }
    
    convertedData.push(convertedRow);
  }
  
  return convertedData;
}

// ===============================
// マッピング処理
// ===============================

/**
 * 項目マッピングシートからマッピングデータを取得
 * @return {Array<Array>} マッピングデータ [[フォーマットA項目, フォーマットB項目], ...]
 */
function getMappingData() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const mappingSheet = spreadsheet.getSheetByName(SHEET_NAMES.MAPPING);
    
    if (!mappingSheet) {
      console.error('項目マッピングシートが見つかりません');
      return [];
    }
    
    const lastRow = mappingSheet.getLastRow();
    
    if (lastRow < 2) {
      console.warn('項目マッピングシートにデータがありません（ヘッダー行を除く）');
      return [];
    }
    
    // A列とB列のデータを取得（2行目から開始、ヘッダー行をスキップ）
    const range = mappingSheet.getRange(2, 1, lastRow - 1, 2);
    const values = range.getValues();
    
    // 空行や不正なデータをフィルタリング
    const validMappings = values.filter(row => {
      return row[0] && row[1] && // 両方の列に値が存在
             typeof row[0] === 'string' && 
             typeof row[1] === 'string' &&
             row[0].trim() !== '' && 
             row[1].trim() !== '';
    });
    
    if (validMappings.length === 0) {
      console.warn('有効なマッピングデータが見つかりません');
      return [];
    }
    
    console.log(`${validMappings.length}件のマッピングデータを取得しました`);
    return validMappings;
    
  } catch (error) {
    console.error('マッピングデータの取得中にエラーが発生しました:', error);
    throw new Error(`マッピングデータの取得に失敗しました: ${error.message}`);
  }
}

/**
 * マッピングデータの妥当性を検証
 * @param {Array<Array>} mappingData マッピングデータ
 * @return {Object} 検証結果 {isValid: boolean, errors: Array<string>}
 */
function validateMappingData(mappingData) {
  const errors = [];
  const formatAItems = new Set();
  const formatBItems = new Set();
  
  for (let i = 0; i < mappingData.length; i++) {
    const [formatA, formatB] = mappingData[i];
    const rowNumber = i + 1;
    
    // フォーマットA項目の重複チェック
    if (formatAItems.has(formatA)) {
      errors.push(`行${rowNumber}: フォーマットA項目「${formatA}」が重複しています`);
    } else {
      formatAItems.add(formatA);
    }
    
    // フォーマットB項目の重複チェック
    if (formatBItems.has(formatB)) {
      errors.push(`行${rowNumber}: フォーマットB項目「${formatB}」が重複しています`);
    } else {
      formatBItems.add(formatB);
    }
    
    // 項目名の妥当性チェック
    if (formatA.length > 100) {
      errors.push(`行${rowNumber}: フォーマットA項目名が長すぎます（100文字以内）`);
    }
    
    if (formatB.length > 100) {
      errors.push(`行${rowNumber}: フォーマットB項目名が長すぎます（100文字以内）`);
    }
    
    // 特殊文字のチェック（フォーマットB項目名）
    if (!/^[a-zA-Z0-9_]+$/.test(formatB)) {
      errors.push(`行${rowNumber}: フォーマットB項目名「${formatB}」に使用できない文字が含まれています（英数字とアンダースコアのみ）`);
    }
  }
  
  return {
    isValid: errors.length === 0,
    errors: errors
  };
}

/**
 * フォーマットA入力シートのヘッダーとマッピングデータの整合性をチェック
 * @param {Array<string>} inputHeaders 入力シートのヘッダー
 * @param {Array<Array>} mappingData マッピングデータ
 * @return {Object} チェック結果 {isValid: boolean, missingColumns: Array<string>, warnings: Array<string>}
 */
function validateHeaderMapping(inputHeaders, mappingData) {
  const missingColumns = [];
  const warnings = [];
  
  for (const mapping of mappingData) {
    const formatAColumn = mapping[0];
    
    if (!inputHeaders.includes(formatAColumn)) {
      missingColumns.push(formatAColumn);
    }
  }
  
  // 入力シートにあるがマッピングにない列をチェック
  for (const header of inputHeaders) {
    const isMapped = mappingData.some(mapping => mapping[0] === header);
    if (!isMapped) {
      warnings.push(`入力シートの列「${header}」はマッピングに定義されていません`);
    }
  }
  
  return {
    isValid: missingColumns.length === 0,
    missingColumns: missingColumns,
    warnings: warnings
  };
}

/**
 * データ変換のプレビューを生成（デバッグ用）
 * @param {Object} inputData 入力データ
 * @param {Array<Array>} mappingData マッピングデータ
 * @param {number} maxRows プレビューする最大行数（デフォルト: 5）
 * @return {Object} プレビューデータ
 */
function generateConversionPreview(inputData, mappingData, maxRows = 5) {
  const { headers, data } = inputData;
  const previewData = {
    mappingInfo: mappingData,
    inputHeaders: headers,
    outputHeaders: mappingData.map(mapping => mapping[1]),
    sampleData: []
  };
  
  const rowsToProcess = Math.min(data.length, maxRows);
  
  for (let i = 0; i < rowsToProcess; i++) {
    const inputRow = data[i];
    const outputRow = [];
    
    for (const mapping of mappingData) {
      const formatAColumn = mapping[0];
      const columnIndex = headers.indexOf(formatAColumn);
      
      if (columnIndex !== -1) {
        outputRow.push(inputRow[columnIndex] || '');
      } else {
        outputRow.push('');
      }
    }
    
    previewData.sampleData.push({
      input: inputRow,
      output: outputRow
    });
  }
  
  return previewData;
}

// ===============================
// ユーティリティ関数
// ===============================

/**
 * 成功メッセージを表示
 * @param {string} message 表示するメッセージ
 */
function showSuccess(message) {
  try {
    SpreadsheetApp.getUi().alert('成功', message, SpreadsheetApp.getUi().ButtonSet.OK);
    console.log(`成功: ${message}`);
  } catch (error) {
    console.log(`成功メッセージ表示エラー: ${error.message}`);
    console.log(`成功: ${message}`);
  }
}

/**
 * エラーメッセージを表示
 * @param {string} message 表示するエラーメッセージ
 */
function showError(message) {
  try {
    SpreadsheetApp.getUi().alert('エラー', message, SpreadsheetApp.getUi().ButtonSet.OK);
    console.error(`エラー: ${message}`);
  } catch (error) {
    console.error(`エラーメッセージ表示エラー: ${error.message}`);
    console.error(`エラー: ${message}`);
  }
}

/**
 * 警告メッセージを表示
 * @param {string} message 表示する警告メッセージ
 */
function showWarning(message) {
  try {
    SpreadsheetApp.getUi().alert('警告', message, SpreadsheetApp.getUi().ButtonSet.OK);
    console.warn(`警告: ${message}`);
  } catch (error) {
    console.warn(`警告メッセージ表示エラー: ${error.message}`);
    console.warn(`警告: ${message}`);
  }
}

/**
 * 現在の日時を文字列として取得
 * @param {string} format フォーマット形式（'datetime', 'date', 'time'）
 * @return {string} フォーマットされた日時文字列
 */
function getCurrentDateTime(format = 'datetime') {
  const now = new Date();
  
  switch (format) {
    case 'date':
      return Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    case 'time':
      return Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');
    case 'datetime':
    default:
      return Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  }
}

// ===============================
// カスタムメニュー設定
// ===============================

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
  const releaseDate = '2025-06-27';
  
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
