/**
 * メイン処理ファイル
 * フォーマットA入力シートからフォーマットB出力シートへの変換を実行
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
 * @return {Array<Array>} 入力データの配列（ヘッダー行を除く）
 */
function getInputData() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = spreadsheet.getSheetByName(SHEET_NAMES.INPUT);
  
  const lastRow = inputSheet.getLastRow();
  const lastCol = inputSheet.getLastColumn();
  
  if (lastRow <= 1) {
    return []; // ヘッダー行のみまたはデータなし
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
