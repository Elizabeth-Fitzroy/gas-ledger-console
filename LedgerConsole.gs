const LEDGER_CONSOLE = {
  sheets: {
    ledger: 'Ledger',
    barcode: 'BarcodeScan',
    console: 'Console'
  },
  consoleStartRow: 2,
  headerSearchRows: 3,
  ledger: {
    purchaseId: ['仕入れID', 'Purchase ID', 'PurchaseID'],
    isbn: ['ISBN'],
    author: ['著者', 'Author'],
    series: ['シリーズ', 'Series'],
    title: ['書名', 'Title'],
    subtitle: ['Subtitle', 'サブタイトル'],
    confirmTitle: ['確認用書名', '確認用タイトル'],
    condition: ['コンディション', 'Condition'],
    status: ['ステータス', 'Status'],
    stockDate: ['棚卸日', 'Stock Date', 'Inventory Date'],
    price: ['販売価格', '価格', 'Price'],
    transactionId: ['取引ID', '取引 ID', 'Transaction ID'],
    destination: ['配送先', 'Destination'],
    trackingNumber: ['送り状番号等', '送り状番号', '追跡番号'],
    memo: ['備考', 'Memo'],
    inStockDate: ['In stock', 'In Stock', 'Instock'],
    readyDate: ['Ready'],
    listOnDate: ['List on', 'List On', 'Liston'],
    soldDate: ['Sold'],
    notSaleDate: ['NotSale', 'Not Sale']
  },
  console: {
    rowId: ['Ledger_Row_ID', 'Ledger Row ID'],
    purchaseId: ['仕入れID', 'Purchase ID', 'PurchaseID'],
    isbn: ['ISBN'],
    author: ['著者', 'Author'],
    series: ['シリーズ', 'Series'],
    title: ['書名', 'Title'],
    currentCondition: ['現在のコンディション', 'Current Condition'],
    currentStatus: ['現在のステータス', 'Current Status'],
    currentDate: ['現在の棚卸日', 'Current Stock Date', 'Current Inventory Date'],
    newStatus: ['新ステータス', 'New Status'],
    newCondition: ['新コンディション', 'New Condition'],
    newDate: ['新棚卸日', 'New Stock Date', 'New Inventory Date'],
    price: ['販売価格', '価格', 'Price'],
    transactionId: ['取引ID', '取引 ID', 'Transaction ID'],
    destination: ['配送先', 'Destination'],
    trackingNumber: ['送り状番号等', '送り状番号', '追跡番号'],
    memo: ['備考', 'Memo']
  }
};

/**
 * あいまい検索対応テキスト正規化関数
 * - 半角 / 全角の統一 (NFKC)
 * - 大文字 / 小文字の統一 (toLowerCase)
 * - ひらがな ⇄ カタカナの同一視 (全角カタカナへ統一)
 * - 前後トリム
 */
function normalizeText(value) {
  if (value === null || value === undefined) return '';
  let str = String(value).toLowerCase().normalize('NFKC').trim();
  // ひらがな (U+3041〜U+3096) を カタカナ (U+30A1〜U+30F6) に変換 (差分: 0x60)
  return str.replace(/[\u3041-\u3096]/g, function(ch) {
    return String.fromCharCode(ch.charCodeAt(0) + 0x60);
  });
}

function normalizeIsbn(value) {
  return normalizeText(value).replace(/[^0-9X]/gi, '').toUpperCase();
}

function getTodayFormatted() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
}

function findHeaderRowAndMap(sheet, aliases, requiredKeys) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (!lastRow || !lastColumn) throw new Error(`シート「${sheet.getName()}」にデータがありません。`);

  const rowsToScan = Math.min(LEDGER_CONSOLE.headerSearchRows, lastRow);
  const values = sheet.getRange(1, 1, rowsToScan, lastColumn).getDisplayValues();
  const normalizedAliases = {};
  Object.keys(aliases).forEach(key => {
    normalizedAliases[key] = aliases[key].map(normalizeText);
  });

  for (let rowIndex = 0; rowIndex < values.length; rowIndex++) {
    const map = {};
    values[rowIndex].forEach((value, columnIndex) => {
      const normalizedValue = normalizeText(value);
      Object.keys(normalizedAliases).some(key => {
        if (normalizedAliases[key].indexOf(normalizedValue) !== -1 && map[key] === undefined) {
          map[key] = columnIndex + 1;
          return true;
        }
        return false;
      });
    });
    if (requiredKeys.every(key => map[key] !== undefined)) {
      return { headerRow: rowIndex + 1, map: map, lastColumn: lastColumn };
    }
  }
  throw new Error(`シート「${sheet.getName()}」で必要なヘッダーを特定できません。`);
}

function getSheetColumnSpan(map) {
  return Math.max.apply(null, Object.keys(map).map(key => map[key]));
}

function getOptionalHeaderMap(sheet, aliases) {
  try {
    return findHeaderRowAndMap(sheet, aliases, []).map;
  } catch (error) {
    return {};
  }
}

function safeAlert(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (error) {
    // Webアプリやサイドバーからの呼び出しではUIが存在しない場合があります。
  }
}


function fetchToConsole(params) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const ledgerSheet = spreadsheet.getSheetByName(LEDGER_CONSOLE.sheets.ledger);
    const consoleSheet = spreadsheet.getSheetByName(LEDGER_CONSOLE.sheets.console);
    if (!ledgerSheet || !consoleSheet) {
      return { success: false, count: 0, message: 'Ledger または Console シートが存在しません。' };
    }

    const ledgerInfo = findHeaderRowAndMap(
      ledgerSheet,
      LEDGER_CONSOLE.ledger,
      ['purchaseId', 'isbn', 'author', 'series', 'title', 'condition', 'status', 'stockDate']
    );
    const consoleInfo = findHeaderRowAndMap(
      consoleSheet,
      LEDGER_CONSOLE.console,
      Object.keys(LEDGER_CONSOLE.console)
    );
    const ledgerStartRow = ledgerInfo.headerRow + 1;
    const ledgerLastRow = ledgerSheet.getLastRow();
    const ledgerWidth = Math.max(ledgerSheet.getLastColumn(), getSheetColumnSpan(ledgerInfo.map));
    const ledgerValues = ledgerLastRow >= ledgerStartRow
      ? ledgerSheet.getRange(ledgerStartRow, 1, ledgerLastRow - ledgerStartRow + 1, ledgerWidth).getValues()
      : [];

    const mode = params && params.mode === 'barcode' ? 'barcode' : 'keyword';
    const rawKeyword = (params && params.keyword) ? String(params.keyword) : '';
    let barcodeSet = new Set();

    if (mode === 'barcode') {
      const barcodeSheet = spreadsheet.getSheetByName(LEDGER_CONSOLE.sheets.barcode);
      if (!barcodeSheet) return { success: false, count: 0, message: 'BarcodeScan シートが存在しません。' };
      const barcodeColumn = getOptionalHeaderMap(barcodeSheet, { isbn: ['ISBN'] }).isbn || barcodeSheet.getDataRange().getColumn();
      const barcodeLastRow = barcodeSheet.getLastRow();
      if (barcodeLastRow) {
        barcodeSheet.getRange(1, barcodeColumn, barcodeLastRow, 1).getValues().forEach(row => {
          const isbn = normalizeIsbn(row[0]);
          if (isbn) barcodeSet.add(isbn);
        });
      }
      if (!barcodeSet.size) return { success: false, count: 0, message: 'BarcodeScan に有効なISBNがありません。' };
    } else if (!rawKeyword.trim()) {
      return { success: false, count: 0, message: '検索キーワードを入力してください。' };
    }

    // ハイブリッド AND/OR 検索判定
    const isPipeOr = /[|｜]/.test(rawKeyword);
    let keywords = [];
    let isOrSearch = false;

    if (isPipeOr) {
      // 1. パイプ記号が含まれる場合【最優先】: ユーザーの明示的な指定として OR 検索を強制
      keywords = rawKeyword.split(/[|｜]/).map(normalizeText).filter(Boolean);
      isOrSearch = true;
    } else {
      // 2. パイプ記号がない場合: スペース分割＋UI指定モード判定
      keywords = rawKeyword.split(/[\s　]+/).map(normalizeText).filter(Boolean);
      const reqSearchMode = params && params.searchMode ? String(params.searchMode).toUpperCase() : 'AND';
      isOrSearch = (reqSearchMode === 'OR');
    }

    // 書き込みモード・ステータス条件パラメータのパース
    const writeMode = (params && params.writeMode === 'append') ? 'append' : 'overwrite';
    const isNormalStatus = (params && typeof params.isNormalStatus === 'boolean') ? params.isNormalStatus : true;
    const targetStatuses = (params && Array.isArray(params.statuses)) ? params.statuses : [];
    const normalizedTargetStatuses = new Set(targetStatuses.map(normalizeText));

    const matchedRows = [];

    ledgerValues.forEach((row, index) => {
      // 現在のステータスを取得
      const currentStatus = normalizeText(row[ledgerInfo.map.status - 1]);
      const isSold = (currentStatus === 'sold' || currentStatus === '販売済' || currentStatus === '売却済');

      // 動的ステータス判定ロジック
      // 1. targetStatuses が指定されている場合: 指定ステータスのいずれか、または (isNormalStatus=trueなら) Sold 以外に一致
      // 2. targetStatuses が空で isNormalStatus が true の場合: Sold 以外を対象
      // 3. いずれも未指定 (isNormalStatus=false かつ targetStatuses空): 全ステータスを対象
      let statusMatch = false;

      if (targetStatuses.length > 0) {
        if (normalizedTargetStatuses.has(currentStatus)) {
          statusMatch = true;
        }
        if (isNormalStatus && !isSold) {
          statusMatch = true;
        }
      } else if (isNormalStatus) {
        statusMatch = !isSold;
      } else {
        statusMatch = true;
      }

      if (!statusMatch) return;

      let matches = false;
      if (mode === 'barcode') {
        matches = barcodeSet.has(normalizeIsbn(row[ledgerInfo.map.isbn - 1]));
      } else {
        const targetColumns = [
          ledgerInfo.map.author,
          ledgerInfo.map.series,
          ledgerInfo.map.title,
          ledgerInfo.map.subtitle,
          ledgerInfo.map.confirmTitle
        ].filter(Boolean);

        const combinedText = targetColumns
          .map(col => normalizeText(row[col - 1]))
          .join(' ');

        if (isOrSearch) {
          matches = keywords.some(kw => combinedText.indexOf(kw) !== -1);
        } else {
          matches = keywords.every(kw => combinedText.indexOf(kw) !== -1);
        }
      }

      if (matches) {
        matchedRows.push({ rowId: ledgerStartRow + index, row: row });
      }
    });

    const consoleStartRow = consoleInfo.headerRow + 1;
    const consoleLastRow = consoleSheet.getLastRow();
    const consoleWidth = Math.max(consoleSheet.getLastColumn(), getSheetColumnSpan(consoleInfo.map));

    let writeTargetRow = consoleStartRow;

    if (writeMode === 'append') {
      // 追記モード (append): 既存データをクリアせず、最終行の直後からデータを出力
      if (consoleLastRow >= consoleStartRow) {
        writeTargetRow = consoleLastRow + 1;
      }
    } else {
      // 上書きモード (overwrite): 2行目以降の全コンテンツをクリア
      if (consoleLastRow >= consoleStartRow) {
        consoleSheet.getRange(consoleStartRow, 1, consoleLastRow - consoleStartRow + 1, consoleWidth).clearContent();
      }
    }

    if (!matchedRows.length) return { success: true, count: 0, message: '条件に該当するデータが見つかりませんでした。' };

    const output = matchedRows.map(item => {
      const values = new Array(consoleWidth).fill('');
      const source = item.row;
      const set = (key, value) => {
        if (consoleInfo.map[key]) {
          values[consoleInfo.map[key] - 1] = value;
        }
      };

      set('rowId', item.rowId);
      set('purchaseId', source[ledgerInfo.map.purchaseId - 1]);
      set('isbn', source[ledgerInfo.map.isbn - 1]);
      set('author', source[ledgerInfo.map.author - 1]);
      set('series', source[ledgerInfo.map.series - 1]);
      set('title', source[ledgerInfo.map.title - 1]);
      set('currentCondition', source[ledgerInfo.map.condition - 1]);
      set('currentStatus', source[ledgerInfo.map.status - 1]);
      set('currentDate', source[ledgerInfo.map.stockDate - 1]);

      if (ledgerInfo.map.price) set('price', source[ledgerInfo.map.price - 1]);
      if (ledgerInfo.map.transactionId) set('transactionId', source[ledgerInfo.map.transactionId - 1]);
      if (ledgerInfo.map.destination) set('destination', source[ledgerInfo.map.destination - 1]);
      if (ledgerInfo.map.trackingNumber) set('trackingNumber', source[ledgerInfo.map.trackingNumber - 1]);
      if (ledgerInfo.map.memo) set('memo', source[ledgerInfo.map.memo - 1]);

      return values;
    });

    consoleSheet.getRange(writeTargetRow, 1, output.length, consoleWidth).setValues(output);
    return {
      success: true,
      count: output.length,
      message: `${output.length} 件のデータを Console シートに${writeMode === 'append' ? '追記' : '抽出'}しました。`
    };
  } catch (error) {
    return { success: false, count: 0, message: `抽出失敗: ${error.message}` };
  }
}

function parseDateValue(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyyMMdd');
  }
  return String(value === null || value === undefined ? '' : value).trim();
}

function writeBackToLedger() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const ledgerSheet = spreadsheet.getSheetByName(LEDGER_CONSOLE.sheets.ledger);
    const consoleSheet = spreadsheet.getSheetByName(LEDGER_CONSOLE.sheets.console);
    if (!ledgerSheet || !consoleSheet) return { success: false, message: 'Ledger または Console シートが存在しません。' };

    const ledgerInfo = findHeaderRowAndMap(
      ledgerSheet,
      LEDGER_CONSOLE.ledger,
      ['purchaseId', 'isbn', 'author', 'series', 'title', 'condition', 'status', 'stockDate']
    );
    const consoleInfo = findHeaderRowAndMap(
      consoleSheet,
      LEDGER_CONSOLE.console,
      Object.keys(LEDGER_CONSOLE.console)
    );
    const consoleStartRow = consoleInfo.headerRow + 1;
    const consoleLastRow = consoleSheet.getLastRow();
    if (consoleLastRow < consoleStartRow) return { success: false, message: 'Console シートに書き戻し対象のデータがありません。' };

    const consoleWidth = Math.max(consoleSheet.getLastColumn(), getSheetColumnSpan(consoleInfo.map));
    const consoleValues = consoleSheet.getRange(consoleStartRow, 1, consoleLastRow - consoleStartRow + 1, consoleWidth).getValues();
    const ledgerStartRow = ledgerInfo.headerRow + 1;
    const today = getTodayFormatted();

    // ★ 要件 2: ライトバック時の事前バリデーション（取引ID必須チェック）
    const invalidRows = [];
    consoleValues.forEach(row => {
      const value = key => (consoleInfo.map[key] ? row[consoleInfo.map[key] - 1] : '');
      const rowId = value('rowId');
      const newStatus = normalizeText(value('newStatus'));
      const newCondition = String(value('newCondition') || '').trim();
      const newDate = parseDateValue(value('newDate'));

      const price = value('price');
      const transactionId = String(value('transactionId') || '').trim();
      const destination = value('destination');
      const trackingNumber = value('trackingNumber');
      const memo = value('memo');

      const hasAnyUpdate = newStatus !== '' || newCondition !== '' || newDate !== '' ||
                           price !== '' || transactionId !== '' || destination !== '' ||
                           trackingNumber !== '' || memo !== '';

      if (!rowId || !hasAnyUpdate) return;

      // newStatus が Ready または List on の場合、transactionId が必須
      if (newStatus === 'ready' || newStatus === 'list on' || newStatus === 'liston') {
        if (!transactionId) {
          invalidRows.push(rowId);
        }
      }
    });

    if (invalidRows.length > 0) {
      const errorMessage = `【書き戻し中断】ステータスを Ready または List on に変更する場合、取引IDの入力は必須です。（対象行ID: ${invalidRows.join(', ')}）`;
      safeAlert(errorMessage);
      return { success: false, message: errorMessage };
    }

    // ★ ライトバック書き込み処理（更新対象の列だけをピンポイント書き込み）
    let updatedCount = 0;
    let skippedCount = 0;

    consoleValues.forEach(row => {
      const value = key => (consoleInfo.map[key] ? row[consoleInfo.map[key] - 1] : '');
      const rowId = value('rowId');

      const newStatusRaw = String(value('newStatus') || '').trim();
      const newStatusNorm = normalizeText(newStatusRaw);
      const newCondition = String(value('newCondition') || '').trim();
      const newDate = parseDateValue(value('newDate'));

      const price = value('price');
      const transactionId = value('transactionId');
      const destination = value('destination');
      const trackingNumber = value('trackingNumber');
      const memo = value('memo');

      const hasAnyUpdate = newStatusRaw !== '' || newCondition !== '' || newDate !== '' ||
                           price !== '' || transactionId !== '' || destination !== '' ||
                           trackingNumber !== '' || memo !== '';

      if (!rowId || !hasAnyUpdate) {
        skippedCount++;
        return;
      }

      const targetRow = Number.parseInt(rowId, 10);
      if (!Number.isInteger(targetRow) || targetRow < ledgerStartRow) {
        skippedCount++;
        return;
      }

      // 1. ステータス・コンディション・棚卸日の書き込み
      if (newStatusRaw && ledgerInfo.map.status) {
        ledgerSheet.getRange(targetRow, ledgerInfo.map.status).setValue(newStatusRaw);
      }
      if (newCondition && ledgerInfo.map.condition) {
        ledgerSheet.getRange(targetRow, ledgerInfo.map.condition).setValue(newCondition);
      }
      if (ledgerInfo.map.stockDate) {
        ledgerSheet.getRange(targetRow, ledgerInfo.map.stockDate).setValue(newDate || today);
      }

      // 2. ★ 要件 1: ステータス変更時の日付更新（カラム自動記録）
      if (newStatusNorm === 'in stock' || newStatusNorm === 'instock') {
        if (ledgerInfo.map.inStockDate) ledgerSheet.getRange(targetRow, ledgerInfo.map.inStockDate).setValue(today);
      } else if (newStatusNorm === 'ready') {
        if (ledgerInfo.map.readyDate) ledgerSheet.getRange(targetRow, ledgerInfo.map.readyDate).setValue(today);
      } else if (newStatusNorm === 'list on' || newStatusNorm === 'liston') {
        if (ledgerInfo.map.listOnDate) ledgerSheet.getRange(targetRow, ledgerInfo.map.listOnDate).setValue(today);
      } else if (newStatusNorm === 'sold' || newStatusNorm === '販売済' || newStatusNorm === '売却済') {
        if (ledgerInfo.map.soldDate) ledgerSheet.getRange(targetRow, ledgerInfo.map.soldDate).setValue(today);
      } else if (newStatusNorm === 'notsale' || newStatusNorm === 'not sale') {
        if (ledgerInfo.map.notSaleDate) ledgerSheet.getRange(targetRow, ledgerInfo.map.notSaleDate).setValue(today);
      }

      // 3. 追加項目のピンポイント書き込み
      if (price !== '' && ledgerInfo.map.price) ledgerSheet.getRange(targetRow, ledgerInfo.map.price).setValue(price);
      if (transactionId !== '' && ledgerInfo.map.transactionId) ledgerSheet.getRange(targetRow, ledgerInfo.map.transactionId).setValue(transactionId);
      if (destination !== '' && ledgerInfo.map.destination) ledgerSheet.getRange(targetRow, ledgerInfo.map.destination).setValue(destination);
      if (trackingNumber !== '' && ledgerInfo.map.trackingNumber) ledgerSheet.getRange(targetRow, ledgerInfo.map.trackingNumber).setValue(trackingNumber);
      if (memo !== '' && ledgerInfo.map.memo) ledgerSheet.getRange(targetRow, ledgerInfo.map.memo).setValue(memo);

      updatedCount++;
    });

    consoleSheet.getRange(consoleStartRow, 1, consoleValues.length, consoleWidth).clearContent();
    return {
      success: true,
      message: `【ライトバック完了】\n・更新件数: ${updatedCount} 件\n・スキップ: ${skippedCount} 件`
    };
  } catch (error) {
    return { success: false, message: `ライトバック失敗: ${error.message}` };
  }
}
