/**
 * マッピング処理ファイル
 * 項目マッピングシートからマッピング情報を取得・処理
 * 
 * 拡張機能:
 * 1. 複数フィールド結合：フィールド1+フィールド2+フィールド3
 * 2. 固定値設定：[FIXED:値]
 * 3. 4列形式対応：A列(フォーマットA項目名), B列(フォーマットB項目名), C列(変換タイプ), D列(設定値)
 */

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
      
      // ターゲットフィールドが空の場合はスキップ（エラーにしない） - 選択的マッピング
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
 * マッピングデータの妥当性を検証
 * @param {Array<Object>} mappingData マッピングデータ
 * @return {Object} 検証結果 {isValid: boolean, errors: Array<string>}
 */
function validateMappingData(mappingData) {
  const errors = [];
  const warnings = [];
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
    
    // 緩和版：フィールド名の文字種・長さ制限を撤廃
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
        // 緩和版：固定値が空文字列でも許可
        if (mapping.config.fixedValue === undefined || mapping.config.fixedValue === null) {
          warnings.push(`行${rowNumber}: FIXED タイプで固定値が未設定です（空文字として扱います）`);
        }
        break;
        
      default:
        errors.push(`行${rowNumber}: 不明な変換タイプ「${mapping.type}」です`);
    }
    
    // 緩和版：ソースフィールド名の妥当性チェックを除去
  }
  
  return {
    isValid: errors.length === 0,
    errors: errors,
    warnings: warnings
  };
}

/**
 * フォーマットA入力シートのヘッダーとマッピングデータの整合性をチェック
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
    
    // ソースフィールドの存在チェック（緩和版：エラーではなく警告として処理）
    for (const sourceField of mapping.sourceFields) {
      if (!inputHeaders.includes(sourceField)) {
        warnings.push(`【警告】マッピングで指定された項目「${sourceField}」が入力シートに見つかりません（空値として処理されます）`);
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
  
  // 緩和版では警告があっても処理継続（常にvalidとして返す）
  return {
    isValid: true,  // 緩和版：常にtrueにして処理継続
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
      // 緩和版：固定値が未設定の場合は空文字を返す
      return mapping.config.fixedValue !== undefined && mapping.config.fixedValue !== null 
        ? mapping.config.fixedValue.toString() 
        : '';
      
    case 'CONCAT':
      const values = [];
      for (const sourceField of mapping.sourceFields) {
        const columnIndex = headers.indexOf(sourceField);
        if (columnIndex !== -1) {
          const value = inputRow[columnIndex];
          // 緩和版：空欄データも許容し、非空の値のみを結合対象とする
          if (value !== null && value !== undefined && value !== '') {
            values.push(value.toString());
          }
        }
      }
      return values.join(mapping.config.separator || '');
      
    case 'NORMAL':
    default:
      if (mapping.sourceFields.length > 0) {
        const sourceField = mapping.sourceFields[0];
        const columnIndex = headers.indexOf(sourceField);
        if (columnIndex !== -1) {
          const value = inputRow[columnIndex];
          // 緩和版：空欄データ許容（null, undefined, 空文字はそのまま空文字を返す）
          return value !== null && value !== undefined ? value.toString() : '';
        }
      }
      return '';  // 緩和版：フィールドが見つからない場合は空文字を返す
  }
}
