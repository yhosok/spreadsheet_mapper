/**
 * ユーティリティ関数ファイル
 * 共通で使用される汎用的な関数群
 */

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
 * 確認ダイアログを表示
 * @param {string} title ダイアログのタイトル
 * @param {string} message 確認メッセージ
 * @return {boolean} ユーザーがOKを選択した場合true
 */
function showConfirmation(title, message) {
  try {
    const response = SpreadsheetApp.getUi().alert(
      title, 
      message, 
      SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
    );
    return response === SpreadsheetApp.getUi().Button.OK;
  } catch (error) {
    console.error(`確認ダイアログ表示エラー: ${error.message}`);
    return false;
  }
}

/**
 * 文字列が空または空白のみかチェック
 * @param {string} str チェックする文字列
 * @return {boolean} 空または空白のみの場合true
 */
function isEmptyOrWhitespace(str) {
  return !str || typeof str !== 'string' || str.trim().length === 0;
}

/**
 * 配列が空かチェック
 * @param {Array} arr チェックする配列
 * @return {boolean} 配列が空の場合true
 */
function isEmptyArray(arr) {
  return !Array.isArray(arr) || arr.length === 0;
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

/**
 * シート名の妥当性をチェック
 * @param {string} sheetName チェックするシート名
 * @return {boolean} 妥当な場合true
 */
function isValidSheetName(sheetName) {
  if (isEmptyOrWhitespace(sheetName)) {
    return false;
  }
  
  // Googleスプレッドシートで使用できない文字をチェック
  const invalidChars = ['[', ']', '*', '?', ':', '\\', '/'];
  for (const char of invalidChars) {
    if (sheetName.includes(char)) {
      return false;
    }
  }
  
  // 長さチェック（100文字以内）
  return sheetName.length <= 100;
}

/**
 * データを安全にログ出力（大量データの場合は制限）
 * @param {string} label ログのラベル
 * @param {any} data ログ出力するデータ
 * @param {number} maxLength 最大出力文字数（デフォルト: 1000）
 */
function safeLog(label, data, maxLength = 1000) {
  try {
    let logString = `${label}: `;
    
    if (typeof data === 'object') {
      logString += JSON.stringify(data);
    } else {
      logString += String(data);
    }
    
    if (logString.length > maxLength) {
      logString = logString.substring(0, maxLength) + '... (truncated)';
    }
    
    console.log(logString);
  } catch (error) {
    console.log(`${label}: [ログ出力エラー: ${error.message}]`);
  }
}

/**
 * 配列を指定されたサイズのチャンクに分割
 * @param {Array} array 分割する配列
 * @param {number} chunkSize チャンクサイズ
 * @return {Array<Array>} 分割された配列の配列
 */
function chunkArray(array, chunkSize) {
  if (!Array.isArray(array) || chunkSize <= 0) {
    return [array];
  }
  
  const chunks = [];
  for (let i = 0; i < array.length; i += chunkSize) {
    chunks.push(array.slice(i, i + chunkSize));
  }
  return chunks;
}

/**
 * 重複を除去した配列を返す
 * @param {Array} array 重複を除去する配列
 * @return {Array} 重複が除去された配列
 */
function uniqueArray(array) {
  if (!Array.isArray(array)) {
    return [];
  }
  
  return [...new Set(array)];
}

/**
 * シートが存在するかチェック
 * @param {string} sheetName チェックするシート名
 * @return {boolean} シートが存在する場合true
 */
function sheetExists(sheetName) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    return spreadsheet.getSheetByName(sheetName) !== null;
  } catch (error) {
    console.error(`シート存在チェックエラー: ${error.message}`);
    return false;
  }
}

/**
 * スプレッドシートの情報を取得
 * @return {Object} スプレッドシートの基本情報
 */
function getSpreadsheetInfo() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();
    
    return {
      id: spreadsheet.getId(),
      name: spreadsheet.getName(),
      url: spreadsheet.getUrl(),
      sheetCount: sheets.length,
      sheetNames: sheets.map(sheet => sheet.getName()),
      lastModified: getCurrentDateTime()
    };
  } catch (error) {
    console.error(`スプレッドシート情報取得エラー: ${error.message}`);
    return null;
  }
}

/**
 * パフォーマンス測定のためのタイマークラス
 */
class Timer {
  constructor() {
    this.startTime = null;
    this.endTime = null;
  }
  
  start() {
    this.startTime = new Date();
    console.log(`タイマー開始: ${getCurrentDateTime()}`);
  }
  
  stop() {
    this.endTime = new Date();
    const duration = this.endTime - this.startTime;
    console.log(`タイマー終了: ${getCurrentDateTime()}, 実行時間: ${duration}ms`);
    return duration;
  }
  
  getDuration() {
    if (this.startTime && this.endTime) {
      return this.endTime - this.startTime;
    }
    return null;
  }
}
