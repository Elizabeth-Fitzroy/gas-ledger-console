/**
 * @fileoverview 在庫管理システム「販売・出品コンソール」GASバックエンド
 * @description GoogleスプレッドシートのLedgerシート・P-Masterシートと連携し、
 *              サイドバーUIからの検索・プレビュー・データ更新を提供します。
 *              ヘッダー行（1行目）から各列の位置を完全動的に取得し、列移動や増減に柔軟に対応します。
 */

// ====================================================================
// 1. スプレッドシート UI 起動・メニュー設定
// ====================================================================

/**
 * スプレッドシート起動時にカスタムメニューを追加する関数
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📚 在庫管理')
    .addItem('販売・出品コンソールを開く', 'showSalesConsoleSidebar')
    .addToUi();
}

/**
 * 販売・出品コンソールのサイドバーUIを表示する関数
 */
function showSalesConsoleSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('📚 販売・出品コンソール')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ====================================================================
// 2. ヘッダー動的マッピング・シート取得ユーティリティ
// ====================================================================

/**
 * 対象シートを取得する（シート名の表記ゆれに対応）
 * @param {SpreadsheetApp.Spreadsheet} ss - 対象スプレッドシート
 * @param {string[]} candidateNames - 検索するシート名候補
 * @returns {SpreadsheetApp.Sheet|null} シートオブジェクト
 */
function getSheetByCandidates_(ss, candidateNames) {
  for (let i = 0; i < candidateNames.length; i++) {
    const sheet = ss.getSheetByName(candidateNames[i]);
    if (sheet) return sheet;
  }
  return null;
}

/**
 * シートの1行目から列ヘッダーのインデックス（0始まり）を動的にマッピングする
 * @param {Array<string>} headers - 1行目のヘッダー配列
 * @returns {Object} 列名キーと列インデックス（0始まり）のマッピング辞書
 */
function createHeaderMap_(headers) {
  const map = {};
  if (!headers || !Array.isArray(headers)) return map;

  headers.forEach((header, index) => {
    if (header !== undefined && header !== null) {
      const cleanHeader = String(header).trim();
      if (cleanHeader) {
        map[cleanHeader] = index;
      }
    }
  });

  return map;
}

/**
 * 定義された候補名リストから、該当する列インデックスを取得する
 * @param {Object} headerMap - createHeaderMap_ で作成した辞書
 * @param {string[]} candidates - 探索する列名候補（優先度順）
 * @returns {number} 列インデックス（0始まり）。見つからない場合は -1
 */
function findColIndex_(headerMap, candidates) {
  for (let i = 0; i < candidates.length; i++) {
    const name = candidates[i];
    if (headerMap[name] !== undefined) {
      return headerMap[name];
    }
  }
  return -1;
}

/**
 * Ledgerシートの主要列インデックス構造体をヘッダー行から動的に取得する
 * ※ 列記号や列番号をハードコードせず、1行目のヘッダー名から動的に特定します。
 * @param {Object} headerMap - ヘッダーマッピング辞書
 * @returns {Object} 各フィールドに対応する列番号（0始まり）
 */
function getLedgerColumnIndices_(headerMap) {
  return {
    purchaseId: findColIndex_(headerMap, ['仕入れID', '仕入ID', 'SKU', 'ID']),
    isbn: findColIndex_(headerMap, ['ISBN', 'ISBNコード', 'ISBN13', 'JAN']),
    condition: findColIndex_(headerMap, ['コンディション', '状態', 'Condition', '状態ランク', 'コンディションランク']),
    stickerPrice: findColIndex_(headerMap, ['シール価格', '仕入価格', '仕入れ価格', '原価']),
    receiptPrice: findColIndex_(headerMap, ['レシート価格']),
    receiptNumber: findColIndex_(headerMap, ['レシート番号']),
    status: findColIndex_(headerMap, ['ステータス', 'Status', '販売ステータス']),
    inventoryDate: findColIndex_(headerMap, ['棚卸日', 'CountDate']),
    inStockDate: findColIndex_(headerMap, ['In stock', '仕入日', '入庫日']),
    readyDate: findColIndex_(headerMap, ['Ready', 'セット組日']),
    listOnDate: findColIndex_(headerMap, ['List on', '出品日', '出品日時']),
    soldDate: findColIndex_(headerMap, ['Sold', '売上日', '販売日', '売却日']),
    price: findColIndex_(headerMap, ['販売価格', '売価', 'Price', '出品価格']),
    transactionId: findColIndex_(headerMap, ['取引ID', 'メルカリ商品ID', '商品ID', 'メルカリID', '注文ID', 'TransactionID']),
    shippingInfo: findColIndex_(headerMap, ['配送先', '送付先', 'お届け先', '配送情報', '宛先', 'ShippingInfo']),
    note: findColIndex_(headerMap, ['備考', 'メモ', 'Note', '特記事項']),
    bookTitle: findColIndex_(headerMap, ['書名', '確認用書名', 'タイトル', '書籍名', '商品名']),
    series: findColIndex_(headerMap, ['シリーズ', 'シリーズ名']),
    vol: findColIndex_(headerMap, ['Vol', '巻数', '巻'])
  };
}

// ====================================================================
// 3. 検索機能 API (searchLedger)
// ====================================================================

/**
 * ISBNまたは書名キーワード、仕入れIDでLedgerシートから書籍を検索する
 * @param {string} queryText - 検索キーワード（ISBN、書名の一部、仕入れIDなど）
 * @returns {Object} 検索結果オブジェクト { success: boolean, items: Array, count: number, message: string }
 */
function searchLedger(queryText) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ledgerSheet = getSheetByCandidates_(ss, ['Ledger', 'ledger', '元帳', '台帳']);
    
    if (!ledgerSheet) {
      return {
        success: false,
        items: [],
        count: 0,
        message: '『Ledger』シートが見つかりませんでした。シート名をご確認ください。'
      };
    }

    const data = ledgerSheet.getDataRange().getValues();
    if (data.length < 2) {
      return {
        success: true,
        items: [],
        count: 0,
        message: 'データが存在しません。'
      };
    }

    // 1行目のヘッダーから列位置を動的解決
    const headers = data[0];
    const headerMap = createHeaderMap_(headers);
    const cols = getLedgerColumnIndices_(headerMap);

    if (cols.purchaseId === -1 && cols.isbn === -1) {
      return {
        success: false,
        items: [],
        count: 0,
        message: '『Ledger』シートに「仕入れID」または「ISBN」列が見つかりません。'
      };
    }

    // P-Masterシートが存在すれば、ISBNと書名の辞書を作成（書名補完用）
    const pMasterSheet = getSheetByCandidates_(ss, ['P-Master', 'PMaster', 'Master', 'マスタ']);
    const pMasterDict = pMasterSheet ? buildPMasterDict_(pMasterSheet) : {};

    const rawQuery = String(queryText || '').trim();
    const cleanQuery = rawQuery.toLowerCase().replace(/[-\s]/g, ''); // ハイフン・空白を除去
    const isBlankQuery = cleanQuery.length === 0;

    const results = [];
    const maxResults = 50; // パフォーマンス上限

    // 2行目（インデックス 1）以降を探索
    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      const purchaseId = cols.purchaseId !== -1 ? String(row[cols.purchaseId] || '').trim() : '';
      const isbn = cols.isbn !== -1 ? String(row[cols.isbn] || '').trim() : '';
      const cleanIsbn = isbn.replace(/[-\s]/g, '');

      // 書名の取得（Ledger列またはP-Masterからの補完）
      let title = cols.bookTitle !== -1 ? String(row[cols.bookTitle] || '').trim() : '';
      const series = cols.series !== -1 ? String(row[cols.series] || '').trim() : '';
      const vol = cols.vol !== -1 ? String(row[cols.vol] || '').trim() : '';

      if (!title && cleanIsbn && pMasterDict[cleanIsbn]) {
        title = pMasterDict[cleanIsbn].title || '';
      }

      // 検索マッチング判定
      let isMatch = false;
      if (isBlankQuery) {
        isMatch = true;
      } else {
        const lowerPurchaseId = purchaseId.toLowerCase();
        const lowerTitle = title.toLowerCase();
        const lowerSeries = series.toLowerCase();

        if (
          (cleanIsbn && cleanIsbn.includes(cleanQuery)) ||
          (lowerPurchaseId && lowerPurchaseId.includes(cleanQuery)) ||
          (lowerTitle && lowerTitle.includes(rawQuery.toLowerCase())) ||
          (lowerSeries && lowerSeries.includes(rawQuery.toLowerCase()))
        ) {
          isMatch = true;
        }
      }

      if (isMatch) {
        results.push({
          rowIndex: r + 1, // スプレッドシート上の行番号（1始まり）
          purchaseId: purchaseId,
          isbn: isbn,
          title: title,
          series: series,
          vol: vol,
          status: cols.status !== -1 ? String(row[cols.status] || '').trim() : '',
          condition: cols.condition !== -1 ? String(row[cols.condition] || '').trim() : '',
          stickerPrice: cols.stickerPrice !== -1 ? row[cols.stickerPrice] : '',
          receiptPrice: cols.receiptPrice !== -1 ? row[cols.receiptPrice] : '',
          price: cols.price !== -1 ? row[cols.price] : '',
          transactionId: cols.transactionId !== -1 ? String(row[cols.transactionId] || '').trim() : '',
          shippingInfo: cols.shippingInfo !== -1 ? String(row[cols.shippingInfo] || '').trim() : '',
          note: cols.note !== -1 ? String(row[cols.note] || '').trim() : '',
          inStockDate: cols.inStockDate !== -1 ? formatDateValue_(row[cols.inStockDate]) : '',
          listOnDate: cols.listOnDate !== -1 ? formatDateValue_(row[cols.listOnDate]) : '',
          soldDate: cols.soldDate !== -1 ? formatDateValue_(row[cols.soldDate]) : ''
        });

        if (results.length >= maxResults) {
          break;
        }
      }
    }

    return {
      success: true,
      items: results,
      count: results.length,
      message: results.length > 0 ? `${results.length} 件見つかりました。` : '該当する商品が見つかりませんでした。'
    };

  } catch (error) {
    return {
      success: false,
      items: [],
      count: 0,
      message: '検索処理中にエラーが発生しました: ' + error.toString()
    };
  }
}

// ====================================================================
// 4. データ更新機能 API (updateLedgerItem)
// ====================================================================

/**
 * サイドバーから送信されたデータに基づき、Ledgerシートの該当仕入れID行を更新する
 * ステータスがSoldかつ取引IDが指定されている場合、同じ取引IDを持つセット品行を一括更新します。
 * @param {Object} updateData - 更新データオブジェクト
 * @returns {Object} 処理結果 { success: boolean, message: string, updatedRow: number }
 */
function updateLedgerItem(updateData) {
  try {
    if (!updateData || !updateData.purchaseId) {
      return {
        success: false,
        message: '仕入れIDが指定されていません。'
      };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ledgerSheet = getSheetByCandidates_(ss, ['Ledger', 'ledger', '元帳', '台帳']);
    
    if (!ledgerSheet) {
      return {
        success: false,
        message: '『Ledger』シートが見つかりませんでした。'
      };
    }

    const data = ledgerSheet.getDataRange().getValues();
    if (data.length < 2) {
      return {
        success: false,
        message: 'データが存在しません。'
      };
    }

    // ヘッダー行（1行目）から動的に列インデックスを取得
    const headers = data[0];
    const headerMap = createHeaderMap_(headers);
    const cols = getLedgerColumnIndices_(headerMap);

    if (cols.purchaseId === -1) {
      return {
        success: false,
        message: '『Ledger』シートに「仕入れID」列が見つかりません。'
      };
    }

    const targetId = String(updateData.purchaseId).trim();
    let targetRowIndex = -1; // 0始まり（配列の行インデックス）

    // 指定された仕入れIDのメイン行を特定
    for (let r = 1; r < data.length; r++) {
      if (String(data[r][cols.purchaseId]).trim() === targetId) {
        targetRowIndex = r;
        break;
      }
    }

    if (targetRowIndex === -1) {
      return {
        success: false,
        message: `仕入れID 「${targetId}」 がLedgerシート内に見つかりませんでした。`
      };
    }

    const targetRowNumber = targetRowIndex + 1; // スプレッドシート上の行番号（1始まり）
    const todayStr = getTodayString_();

    const normalizedStatus = String(updateData.status || '').trim();
    const isSoldStatus = normalizedStatus.toLowerCase().includes('sold') || 
                         normalizedStatus.includes('売約') || 
                         normalizedStatus.includes('発送') || 
                         normalizedStatus.includes('売上');
    const inputTransactionId = updateData.transactionId !== undefined ? String(updateData.transactionId).trim() : '';

    // ====================================================================
    // 一括 Sold 更新判定
    // 条件: ステータスが「Sold」かつ「取引ID」が空欄でないこと
    // ====================================================================
    if (isSoldStatus && inputTransactionId !== '' && cols.transactionId !== -1) {
      // シート全体から同一の取引IDを持つ行インデックスを収集
      const matchedRowIndices = [];
      for (let r = 1; r < data.length; r++) {
        const rowTxId = String(data[r][cols.transactionId] || '').trim();
        // 現在の入力取引IDと一致する行、または操作対象のメイン行
        if (rowTxId === inputTransactionId || r === targetRowIndex) {
          matchedRowIndices.push(r);
        }
      }

      // 重複を除去
      const uniqueRowIndices = Array.from(new Set(matchedRowIndices));

      // 各行を更新
      for (let i = 0; i < uniqueRowIndices.length; i++) {
        const rIndex = uniqueRowIndices[i];
        const rowNum = rIndex + 1;
        const isMainTargetRow = (rIndex === targetRowIndex);

        // 1. 販売ステータス (Sold)
        if (cols.status !== -1) {
          ledgerSheet.getRange(rowNum, cols.status + 1).setValue(updateData.status);
        }

        // 2. Sold日付（未設定なら当日日付を補完）
        if (cols.soldDate !== -1) {
          const currentSoldDate = data[rIndex][cols.soldDate];
          if (!currentSoldDate) {
            ledgerSheet.getRange(rowNum, cols.soldDate + 1).setValue(todayStr);
          }
        }

        // 3. 取引ID
        if (cols.transactionId !== -1) {
          ledgerSheet.getRange(rowNum, cols.transactionId + 1).setValue(inputTransactionId);
        }

        // 4. 販売価格（重要: メイン行のみに入力値を保存し、他の行は空欄を維持）
        if (cols.price !== -1) {
          if (isMainTargetRow) {
            if (updateData.price !== undefined && updateData.price !== '') {
              const numPrice = Number(updateData.price);
              ledgerSheet.getRange(rowNum, cols.price + 1).setValue(isNaN(numPrice) ? updateData.price : numPrice);
            } else {
              ledgerSheet.getRange(rowNum, cols.price + 1).setValue('');
            }
          } else {
            // 他のセット品行は販売価格を空欄にする
            ledgerSheet.getRange(rowNum, cols.price + 1).setValue('');
          }
        }

        // 5. 配送先情報（一括反映）
        if (updateData.shippingInfo !== undefined && cols.shippingInfo !== -1) {
          ledgerSheet.getRange(rowNum, cols.shippingInfo + 1).setValue(updateData.shippingInfo);
        }

        // 6. 備考 / メモ（一括反映）
        if (updateData.note !== undefined && cols.note !== -1) {
          ledgerSheet.getRange(rowNum, cols.note + 1).setValue(updateData.note);
        }

        // 7. コンディション（操作対象行のみ個別に更新）
        if (isMainTargetRow && updateData.condition !== undefined && cols.condition !== -1) {
          ledgerSheet.getRange(rowNum, cols.condition + 1).setValue(updateData.condition);
        }
      }

      const totalCount = uniqueRowIndices.length;
      const message = totalCount > 1
        ? `仕入れID 「${targetId}」 および同一取引ID（${inputTransactionId}）のセット品 計 ${totalCount} 件を一括で Sold に更新しました。`
        : `仕入れID 「${targetId}」 の情報（行 ${targetRowNumber}）を Sold に更新しました。`;

      return {
        success: true,
        message: message,
        updatedRow: targetRowNumber,
        batchCount: totalCount
      };

    } else {
      // ====================================================================
      // 通常更新（単一行更新 / 安全ガード）
      // 取引IDが空欄の場合、または Sold 以外のステータス時は選択された1行のみ更新
      // ====================================================================

      // 1. 販売ステータス
      if (updateData.status !== undefined && cols.status !== -1) {
        ledgerSheet.getRange(targetRowNumber, cols.status + 1).setValue(updateData.status);
        
        // ステータスに応じた日付自動補完
        const statusLower = normalizedStatus.toLowerCase();
        if ((statusLower.includes('list') || statusLower.includes('出品')) && cols.listOnDate !== -1) {
          const currentListOn = data[targetRowIndex][cols.listOnDate];
          if (!currentListOn) {
            ledgerSheet.getRange(targetRowNumber, cols.listOnDate + 1).setValue(todayStr);
          }
        } else if (isSoldStatus && cols.soldDate !== -1) {
          const currentSold = data[targetRowIndex][cols.soldDate];
          if (!currentSold) {
            ledgerSheet.getRange(targetRowNumber, cols.soldDate + 1).setValue(todayStr);
          }
        }
      }

      // 2. コンディション
      if (updateData.condition !== undefined && cols.condition !== -1) {
        ledgerSheet.getRange(targetRowNumber, cols.condition + 1).setValue(updateData.condition);
      }

      // 3. 取引ID / メルカリ商品ID
      if (updateData.transactionId !== undefined && cols.transactionId !== -1) {
        ledgerSheet.getRange(targetRowNumber, cols.transactionId + 1).setValue(updateData.transactionId);
      }

      // 4. 販売価格
      if (updateData.price !== undefined && cols.price !== -1) {
        if (updateData.price === '') {
          ledgerSheet.getRange(targetRowNumber, cols.price + 1).setValue('');
        } else {
          const numPrice = Number(updateData.price);
          ledgerSheet.getRange(targetRowNumber, cols.price + 1).setValue(isNaN(numPrice) ? updateData.price : numPrice);
        }
      }

      // 5. 配送先情報
      if (updateData.shippingInfo !== undefined && cols.shippingInfo !== -1) {
        ledgerSheet.getRange(targetRowNumber, cols.shippingInfo + 1).setValue(updateData.shippingInfo);
      }

      // 6. 備考 / メモ
      if (updateData.note !== undefined && cols.note !== -1) {
        ledgerSheet.getRange(targetRowNumber, cols.note + 1).setValue(updateData.note);
      }

      return {
        success: true,
        message: `仕入れID 「${targetId}」 の情報（行 ${targetRowNumber}）を正常に更新しました。`,
        updatedRow: targetRowNumber,
        batchCount: 1
      };
    }

  } catch (error) {
    return {
      success: false,
      message: '更新処理中にエラーが発生しました: ' + error.toString()
    };
  }
}

// ====================================================================
// 5. 補助ユーティリティ関数
// ====================================================================

/**
 * P-Masterシートから { ISBN: { title, series, vol } } の参照マップを作成する
 * @param {SpreadsheetApp.Sheet} pMasterSheet
 * @returns {Object}
 */
function buildPMasterDict_(pMasterSheet) {
  const dict = {};
  try {
    const data = pMasterSheet.getDataRange().getValues();
    if (data.length < 2) return dict;

    const headers = data[0];
    const headerMap = createHeaderMap_(headers);

    const isbnCol = findColIndex_(headerMap, ['ISBN', 'ISBNコード', 'ISBN13', 'JAN', 'Key', 'KEY']);
    const titleCol = findColIndex_(headerMap, ['書名', 'タイトル', '書籍名', 'Title']);
    const seriesCol = findColIndex_(headerMap, ['シリーズ', 'シリーズ名', 'Series']);
    const volCol = findColIndex_(headerMap, ['巻数', 'Vol', '巻']);
    const isbn13Col = findColIndex_(headerMap, ['ISBN13', '計算ISBN13']);

    if (isbnCol === -1 && titleCol === -1) return dict;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const isbnVal = isbnCol !== -1 ? String(row[isbnCol] || '').trim().replace(/[-\s]/g, '') : '';
      const isbn13Val = isbn13Col !== -1 ? String(row[isbn13Col] || '').trim().replace(/[-\s]/g, '') : '';
      const titleVal = titleCol !== -1 ? String(row[titleCol] || '').trim() : '';
      const seriesVal = seriesCol !== -1 ? String(row[seriesCol] || '').trim() : '';
      const volVal = volCol !== -1 ? String(row[volCol] || '').trim() : '';

      const info = { title: titleVal, series: seriesVal, vol: volVal };

      if (isbnVal) dict[isbnVal] = info;
      if (isbn13Val && isbn13Val !== isbnVal) dict[isbn13Val] = info;
    }
  } catch (e) {
    Logger.log('P-Master辞書構築スキップ: ' + e);
  }
  return dict;
}

/**
 * 今日の日付を YYYYMMDD 形式の文字列で取得する
 * @returns {string} 例: "20260826"
 */
function getTodayString_() {
  const d = new Date();
  const year = d.getFullYear();
  const month = ('0' + (d.getMonth() + 1)).slice(-2);
  const day = ('0' + d.getDate()).slice(-2);
  return `${year}${month}${day}`;
}

/**
 * 日付セル値を扱いやすい文字列に整形する
 * @param {any} val - セルの値
 * @returns {string}
 */
function formatDateValue_(val) {
  if (!val) return '';
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = ('0' + (val.getMonth() + 1)).slice(-2);
    const d = ('0' + val.getDate()).slice(-2);
    return `${y}/${m}/${d}`;
  }
  return String(val);
}
