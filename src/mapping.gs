/**
 * マッピング処理ファイル
 * 項目マッピングシートからマッピング情報を取得・処理
 */

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
 * @param {Array<Array>} inputData 入力データ
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
