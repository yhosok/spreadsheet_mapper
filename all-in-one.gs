/**
 * Googleスプレッドシート フォーマット変換ツール - 統合版 v2.0.2
 * 
 * このファイル1つで全ての機能を利用できます。
 * Google Apps Scriptエディタにこのコードをコピー&ペーストしてください。
 * 
 * 機能:
 * - 基本的な1対1フィールドマッピング
 * - 複数フィールド結合機能（フィールド1+フィールド2+フィールド3）
 * - 固定値設定機能（[FIXED:値]）
 * - 2列・4列マッピング形式の両方に対応
 * - 自動的な後方互換性
 * - 緩和されたバリデーション（文字種制限なし）
 * - 空値データの許容
 * - 不存在フィールドの警告表示（エラーにしない）
 * - 選択的マッピング（OUTPUTフィールドが空の場合はスキップ）
 * 
 * 使用方法:
 * 1. Googleスプレッドシートで「拡張機能」→「Apps Script」を選択
 * 2. このファイルの内容を全てコピーしてApps Scriptエディタに貼り付け
 * 3. ファイルを保存
 * 4. スプレッドシートに戻ると「📊 フォーマット変換ツール」メニューが追加されます
 * 
 * 更新履歴:
 * - v2.0.2: 選択的マッピング機能追加（OUTPUTフィールドが空の場合はスキップ）
 * - v2.0.1: バリデーション緩和、空値許容、不存在フィールド対応
 * - v2.0.0: 拡張機能統合、複数フィールド結合、固定値設定機能追加
 * - v1.0.0: 基本的なフォーマット変換機能
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
    if (!inputData) {
      showError('フォーマットA入力シートにヘッダー行が存在しません。');
      return;
    }
    // データ行が空でも処理継続（ヘッダーのみでも可）
    
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
  
  if (lastRow < 1) {
    return null; // シート自体にデータなし
  }
  
  // ヘッダー行を取得
  const headerRange = inputSheet.getRange(1, 1, 1, lastCol);
  const headers = headerRange.getValues()[0];
  
  // データ行の取得（空でも許容）
  let data = [];
  if (lastRow > 1) {
    const dataRange = inputSheet.getRange(2, 1, lastRow - 1, lastCol);
    data = dataRange.getValues();
  }
  // ヘッダーのみでデータ行が空の場合もdata=[]として処理継続
  
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
 * @param {Array<Object>} mappingData マッピングデータ
 * @return {Array} フォーマットBのヘッダー配列
 */
function createHeaderRow(mappingData) {
  return mappingData.map(mapping => mapping.targetField);
}

/**
 * データの変換
 * @param {Object} inputData 入力データ（headers, data）
 * @param {Array<Object>} mappingData マッピングデータ
 * @return {Array<Array>} 変換されたデータ
 */
function convertData(inputData, mappingData) {
  const { headers, data } = inputData;
  const convertedData = [];
  
  for (const row of data) {
    const convertedRow = [];
    
    for (const mapping of mappingData) {
      const convertedValue = convertFieldValue(row, headers, mapping);
      convertedRow.push(convertedValue);
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
 * @return {Array<Object>} マッピングデータ配列
 * 各要素: {
 *   sourceFields: Array<string>,  // ソースフィールド名配列
 *   targetField: string,          // ターゲットフィールド名
 *   type: string,                 // 変換タイプ ('NORMAL', 'CONCAT', 'FIXED')
 *   config: Object                // 設定値 (separator, fixedValue等)
 * }
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
    
    const lastCol = mappingSheet.getLastColumn();
    const columnCount = Math.max(lastCol, 4); // 最低4列確保
    
    // 全データを取得（2行目から開始、ヘッダー行をスキップ）
    const range = mappingSheet.getRange(2, 1, lastRow - 1, columnCount);
    const values = range.getValues();
    
    // 空行や不正なデータをフィルタリング & 構造化
    const validMappings = [];
    
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const [sourceFieldsStr, targetField, conversionType, configValue] = row;
      
      // ソースフィールドのみ必須チェック（ターゲットフィールドは空でも可）
      if (!sourceFieldsStr || 
          typeof sourceFieldsStr !== 'string' ||
          sourceFieldsStr.trim() === '') {
        continue;
      }
      
      // ターゲットフィールドが空の場合はスキップ（エラーにしない）
      if (!targetField || 
          typeof targetField !== 'string' ||
          targetField.trim() === '') {
        console.log(`行 ${i + 2}: ターゲットフィールドが空のためスキップします（ソース: ${sourceFieldsStr}）`);
        continue;
      }
      
      // マッピング情報の構造化
      const mappingInfo = parseMappingRow(sourceFieldsStr.trim(), targetField.trim(), conversionType, configValue);
      
      if (mappingInfo) {
        validMappings.push(mappingInfo);
      }
    }
    
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
 * マッピング行の解析
 * @param {string} sourceFieldsStr ソースフィールド文字列
 * @param {string} targetField ターゲットフィールド名
 * @param {string} conversionType 変換タイプ（オプション）
 * @param {string} configValue 設定値（オプション）
 * @return {Object|null} 解析されたマッピング情報
 */
function parseMappingRow(sourceFieldsStr, targetField, conversionType, configValue) {
  try {
    let mappingInfo = {
      sourceFields: [],
      targetField: targetField,
      type: 'NORMAL',
      config: {}
    };
    
    // 固定値の判定
    if (sourceFieldsStr.startsWith('[FIXED') && sourceFieldsStr.includes(']')) {
      mappingInfo.type = 'FIXED';
      
      // [FIXED:値] 形式から値を抽出
      const match = sourceFieldsStr.match(/\[FIXED(?::(.+))?\]/);
      if (match && match[1]) {
        mappingInfo.config.fixedValue = match[1];
      } else if (configValue) {
        mappingInfo.config.fixedValue = configValue.toString();
      } else {
        mappingInfo.config.fixedValue = '';
      }
      
      mappingInfo.sourceFields = [];
      
    } else if (sourceFieldsStr.includes('+')) {
      // 複数フィールド結合の判定
      mappingInfo.type = 'CONCAT';
      mappingInfo.sourceFields = sourceFieldsStr.split('+').map(field => field.trim()).filter(field => field.length > 0);
      
      // 区切り文字の設定（デフォルト：空文字）
      mappingInfo.config.separator = configValue && configValue.toString() || '';
      
    } else {
      // 通常のマッピング
      mappingInfo.type = 'NORMAL';
      mappingInfo.sourceFields = [sourceFieldsStr];
    }
    
    // 変換タイプが明示的に指定されている場合は上書き
    if (conversionType && typeof conversionType === 'string' && conversionType.trim() !== '') {
      const specifiedType = conversionType.trim().toUpperCase();
      if (['NORMAL', 'CONCAT', 'FIXED'].includes(specifiedType)) {
        mappingInfo.type = specifiedType;
      }
    }
    
    return mappingInfo;
    
  } catch (error) {
    console.error(`マッピング行の解析エラー: ${error.message}`);
    return null;
  }
}

/**
 * 単一フィールドの値を変換
 * @param {Array} inputRow 入力行データ
 * @param {Array<string>} headers ヘッダー配列
 * @param {Object} mapping マッピング情報
 * @return {string} 変換後の値
 */
function convertFieldValue(inputRow, headers, mapping) {
  switch (mapping.type) {
    case 'FIXED':
      return mapping.config.fixedValue || '';
      
    case 'CONCAT':
      const values = [];
      for (const sourceField of mapping.sourceFields) {
        const columnIndex = headers.indexOf(sourceField);
        if (columnIndex !== -1) {
          const value = inputRow[columnIndex];
          // 空値（null, undefined, ''）も許容し、空文字として処理
          if (value !== null && value !== undefined) {
            values.push(value.toString());
          }
        }
        // フィールドが存在しない場合は無視（エラーにしない）
      }
      return values.join(mapping.config.separator || '');
      
    case 'NORMAL':
    default:
      if (mapping.sourceFields.length > 0) {
        const sourceField = mapping.sourceFields[0];
        const columnIndex = headers.indexOf(sourceField);
        if (columnIndex !== -1) {
          const value = inputRow[columnIndex];
          // 空値も許容
          return value !== null && value !== undefined ? value.toString() : '';
        }
      }
      // フィールドが存在しない場合は空文字を返す
      return '';
  }
}

/**
 * マッピングデータの妥当性を検証（緩和版）
 * @param {Array<Object>} mappingData マッピングデータ
 * @return {Object} 検証結果 {isValid: boolean, errors: Array<string>}
 */
function validateMappingData(mappingData) {
  const errors = [];
  const targetFields = new Set();
  
  for (let i = 0; i < mappingData.length; i++) {
    const mapping = mappingData[i];
    const rowNumber = i + 2; // ヘッダー行を考慮
    
    // ターゲットフィールドの重複チェック
    if (targetFields.has(mapping.targetField)) {
      errors.push(`行${rowNumber}: フォーマットB項目「${mapping.targetField}」が重複しています`);
    } else {
      targetFields.add(mapping.targetField);
    }
    
    // 文字種制限と長さ制限を除去（緩和）
    // フィールド名の妥当性チェックは実施しない
    
    // タイプ別の検証
    switch (mapping.type) {
      case 'NORMAL':
        if (mapping.sourceFields.length !== 1) {
          errors.push(`行${rowNumber}: NORMAL タイプでは1つのソースフィールドが必要です`);
        }
        break;
        
      case 'CONCAT':
        if (mapping.sourceFields.length < 2) {
          errors.push(`行${rowNumber}: CONCAT タイプでは複数のソースフィールドが必要です`);
        }
        break;
        
      case 'FIXED':
        if (!mapping.config.fixedValue && mapping.config.fixedValue !== '') {
          errors.push(`行${rowNumber}: FIXED タイプでは固定値の設定が必要です`);
        }
        break;
        
      default:
        errors.push(`行${rowNumber}: 不明な変換タイプ「${mapping.type}」です`);
    }
    
    // ソースフィールド名の制限を除去（緩和）
  }
  
  return {
    isValid: errors.length === 0,
    errors: errors
  };
}

/**
 * フォーマットA入力シートのヘッダーとマッピングデータの整合性をチェック（緩和版）
 * @param {Array<string>} inputHeaders 入力シートのヘッダー
 * @param {Array<Object>} mappingData マッピングデータ
 * @return {Object} チェック結果 {isValid: boolean, missingColumns: Array<string>, warnings: Array<string>}
 */
function validateHeaderMapping(inputHeaders, mappingData) {
  const missingColumns = [];
  const warnings = [];
  const usedHeaders = new Set();
  
  for (const mapping of mappingData) {
    // 固定値の場合はヘッダーチェック不要
    if (mapping.type === 'FIXED') {
      continue;
    }
    
    // ソースフィールドの存在チェック（エラーではなく警告として処理）
    for (const sourceField of mapping.sourceFields) {
      if (!inputHeaders.includes(sourceField)) {
        warnings.push(`マッピングで指定された項目「${sourceField}」が入力シートに見つかりません`);
      } else {
        usedHeaders.add(sourceField);
      }
    }
  }
  
  // 入力シートにあるがマッピングにない列をチェック
  for (const header of inputHeaders) {
    if (!usedHeaders.has(header)) {
      warnings.push(`【警告】入力シートの列「${header}」はマッピングに定義されていません（この列は出力されません）`);
    }
  }
  
  // 緩和版では常にvalidとして処理継続
  return {
    isValid: true,  // 常にtrueにして処理継続
    missingColumns: missingColumns,
    warnings: warnings
  };
}

/**
 * データ変換のプレビューを生成（デバッグ用）
 * @param {Object} inputData 入力データ
 * @param {Array<Object>} mappingData マッピングデータ
 * @param {number} maxRows プレビューする最大行数（デフォルト: 5）
 * @return {Object} プレビューデータ
 */
function generateConversionPreview(inputData, mappingData, maxRows = 5) {
  const { headers, data } = inputData;
  const previewData = {
    mappingInfo: mappingData,
    inputHeaders: headers,
    outputHeaders: mappingData.map(mapping => mapping.targetField),
    sampleData: []
  };
  
  const rowsToProcess = Math.min(data.length, maxRows);
  
  for (let i = 0; i < rowsToProcess; i++) {
    const inputRow = data[i];
    const outputRow = [];
    
    for (const mapping of mappingData) {
      const convertedValue = convertFieldValue(inputRow, headers, mapping);
      outputRow.push(convertedValue);
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
    
    // フォーマットA入力シートとの整合性チェック（緩和版）
    const inputData = getInputData();
    if (inputData && inputData.headers && inputData.headers.length > 0) {
      const headerValidation = validateHeaderMapping(inputData.headers, mappingData);
      
      let message = 'マッピング検証結果:\n\n';
      message += `✅ マッピング項目数: ${mappingData.length}件\n`;
      message += `✅ 入力シート列数: ${inputData.headers.length}件\n\n`;
      
      // 緩和版では不存在項目もエラーではなく警告として表示
      if (headerValidation.warnings.length > 0) {
        message += '⚠️ 警告事項（処理は正常に継続されます）:\n';
        message += headerValidation.warnings.map(warning => `  - ${warning}`).join('\n');
        message += '\n\n';
      }
      
      message += '✅ 検証完了\n';
      message += headerValidation.warnings.length > 0 ? 
        '📝 警告がありますが、処理は正常に実行されます。' : 
        '📝 すべてのチェックに合格しました。';
      showSuccess(message);
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
    if (!inputData) {
      showError('フォーマットA入力シートにヘッダー行が存在しません。');
      return;
    }
    
    if (inputData.data.length === 0) {
      showWarning('フォーマットA入力シートにデータ行がありません（ヘッダーのみ）。\nデータを追加してから再度プレビューしてください。');
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
    `   • B列: フォーマットBの項目名\n` +
    `   • C列: 変換タイプ（オプション）\n` +
    `   • D列: 設定値（オプション）\n\n` +
    
    `📋 基本マッピング例:\n` +
    `   A1: 出荷日             B1: shipping_date\n` +
    `   A2: 顧客名             B2: customer_name\n\n` +
    
    `📋 拡張マッピング例:\n` +
    `   A1: 住所1+住所2+住所3   B1: full_address    C1: CONCAT   D1: 、\n` +
    `   A2: [FIXED:処理済み]    B2: status          C2: FIXED    D2:\n` +
    `   A3: [FIXED]           B3: created_by      C3: FIXED    D3: システム\n\n` +
    
    `⚠️ 注意事項:\n` +
    `   • シート名は正確に入力してください\n` +
    `   • フォーマットB項目名は英数字とアンダースコアのみ使用\n` +
    `   • 重複する項目名は設定しないでください\n` +
    `   • ヘッダー行は自動検出されます\n` +
    `   • 2列・4列どちらの形式でも動作します`;
  
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
    
    `3️⃣ 拡張機能:\n` +
    `   • 複数フィールド結合: フィールド1+フィールド2+フィールド3\n` +
    `   • 固定値設定: [FIXED:値] または [FIXED]\n` +
    `   • 区切り文字指定: 設定値列で区切り文字を指定\n\n` +
    
    `4️⃣ 結果確認:\n` +
    `   • フォーマットB出力シートが自動生成されます\n` +
    `   • 変換されたデータを確認してください\n\n` +
    
    `🔧 トラブルシューティング:\n` +
    `   • エラーが出る場合は「マッピング検証」を実行\n` +
    `   • シート名が正確か確認\n` +
    `   • データの形式が正しいか確認\n` +
    `   • 拡張機能の記法が正しいか確認`;
  
  showSuccess(message);
}

/**
 * バージョン情報の表示
 */
function showVersionInfo() {
  const version = '2.0.0';
  const releaseDate = '2025-07-08';
  
  const message = `📱 フォーマット変換ツール\n\n` +
    `バージョン: ${version}\n` +
    `リリース日: ${releaseDate}\n\n` +
    `🎯 主な機能:\n` +
    `   • CSVフォーマット変換\n` +
    `   • 項目マッピング設定\n` +
    `   • 複数フィールド結合機能\n` +
    `   • 固定値設定機能\n` +
    `   • データ検証とプレビュー\n` +
    `   • エラーハンドリング\n\n` +
    `💻 開発者: フォーマット変換チーム\n` +
    `📧 サポート: 管理者までお問い合わせください`;
  
  showSuccess(message);
}
