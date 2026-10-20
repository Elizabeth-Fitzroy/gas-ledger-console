// ====================================================================
// 2. 書籍検索・API補完処理（個別・コア処理）
// ====================================================================

function updateUsingOpenBD() {
  updateLookupBookCore("OpenBD");
}
function updateUsingGoogleBooks() {
  updateLookupBookCore("Google");
}
function updateUsingNDL() {
  updateLookupBookCore("NDL");
}
function updateLookupBookIntegrated() {
  updateLookupBookCore("All");
}

/**
 * コア処理：指定されたAPIモードに基づいて実行
 * @param {string} mode - "OpenBD" | "Google" | "NDL" | "All"
 */
function updateLookupBookCore(mode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const scanSheet = ss.getSheetByName("BarcodeScan");
  const masterSheet = ss.getSheetByName("P-Master");   

  if (!scanSheet || !masterSheet) {
    SpreadsheetApp.getUi().alert("シート名を確認してください。");
    return;
  }

  // --- 1. シートの列マッピングを正しく準備する ---
  const scanHeaders = scanSheet.getDataRange().getValues()[0];
  const colMap = {};
  scanHeaders.forEach((header, index) => {
    if (header) {
      colMap[String(header).trim()] = index + 1;
    }
  });

  const masterValues = masterSheet.getDataRange().getValues();
  const masterMap = new Map();
  const masterHeaders = masterValues[0];
  const mColMap = {};
  masterHeaders.forEach((header, index) => {
    if (header) {
      mColMap[String(header).trim()] = index; 
    }
  });

  // 2行目から順にマスタをマップに格納
  for (let r = 1; r < masterValues.length; r++) {
    const row = masterValues[r];
    const isbnMasterCol = mColMap["ISBN"] !== undefined ? mColMap["ISBN"] : 3; 
    const isbnMaster = String(row[isbnMasterCol]).replace(/[-\s]/g, "").trim();
    
    if (!isbnMaster) continue;
    
    if (isbnMaster.length === 13 && isbnMaster.startsWith("978")) {
      let isbn10Key = convertISBN13to10(isbnMaster);
      masterMap.set(isbn10Key, row);
    } else if (isbnMaster.length === 10) {
      let isbn13Key = convertISBN10to13(isbnMaster);
      masterMap.set(isbn13Key, row);
    }
    masterMap.set(isbnMaster, row);
  }
  
  const lastRow = scanSheet.getLastRow();
  if (lastRow < 2) return;

  for (let i = 2; i <= lastRow; i++) {
    let isbnCol = colMap["ISBN"] || 1; 
    let isbnRaw = String(scanSheet.getRange(i, isbnCol).getValue()).replace(/[-\s]/g, "").trim();
    let isbnMatch = isbnRaw.match(/\b(\d{9}[\dX]|978\d{10})\b/);
    
    let statusCol = colMap["ステータス"] || 14;
    let titleCol = colMap["書名"] || 5;
    
    let status = scanSheet.getRange(i, statusCol).getValue().toString();
    let titleVal = scanSheet.getRange(i, titleCol).getValue().toString(); 
    
    if (!isbnMatch || status.indexOf("OK") !== -1 || titleVal !== "") continue;
    
    const isbn = isbnMatch[0];
    
    // --- 2. P-master を最優先チェック ---
    if (masterMap.has(isbn)) {
      const m = masterMap.get(isbn);
      
      const targetCols = ["分類", "著者", "シリーズ", "書名", "SubTitle", "Vol", "出版社", "発行年月", "定価", "判型", "頁数", "仕入価格", "備考"];
      targetCols.forEach(colName => {
        if (colMap[colName] && mColMap[colName] !== undefined) {
          scanSheet.getRange(i, colMap[colName]).setValue(m[mColMap[colName]] || "");
        }
      });
      
      scanSheet.getRange(i, colMap["ステータス"] || 14).setValue("OK: P-master参照");
      continue; 
    }

    // --- 3. 指定されたAPIモードに基づいて叩き分ける ---
    let result = null;
    if (mode === "OpenBD") {
      result = fetchOpenBD(isbn);
    } else if (mode === "Google") {
      result = fetchGoogleBooks(isbn);
    } else if (mode === "NDL") {
      result = fetchNDL(isbn);
    } else {
      result = fetchOpenBD(isbn) || fetchGoogleBooks(isbn) || fetchNDL(isbn);
    }

    if (result) {
      // 汎用データクレンジングを実行
      const cleaned = cleanData(result);
      
      if (colMap["分類"]) scanSheet.getRange(i, colMap["分類"]).setValue(cleaned.cCode || "");
      if (colMap["著者"]) scanSheet.getRange(i, colMap["著者"]).setValue(cleaned.author || "");
      if (colMap["書名"]) scanSheet.getRange(i, colMap["書名"]).setValue(cleaned.title || "");
      if (colMap["SubTitle"]) scanSheet.getRange(i, colMap["SubTitle"]).setValue(cleaned.subTitle || "");
      if (colMap["Vol"]) scanSheet.getRange(i, colMap["Vol"]).setValue(cleaned.volume || "");
      if (colMap["出版社"]) scanSheet.getRange(i, colMap["出版社"]).setValue(cleaned.publisher || "");
      if (colMap["発行年月"]) scanSheet.getRange(i, colMap["発行年月"]).setValue(cleaned.date || "");
      if (colMap["定価"]) scanSheet.getRange(i, colMap["定価"]).setValue(cleaned.price || "");
      if (colMap["判型"]) scanSheet.getRange(i, colMap["判型"]).setValue(cleaned.format || "");
      if (colMap["頁数"]) scanSheet.getRange(i, colMap["頁数"]).setValue(cleaned.pages || "");
      
      scanSheet.getRange(i, colMap["ステータス"] || 14).setValue("OK: " + cleaned.source);
    } else {
      if (mode === "All") {
        scanSheet.getRange(i, colMap["ステータス"] || 14).setValue("NG: API該当なし");
      }
    }
    Utilities.sleep(500); 
  }
}

// --- API取得関数群 ---

function fetchOpenBD(isbn) {
  try {
    const res = UrlFetchApp.fetch("https://api.openbd.jp/v1/get?isbn=" + isbn);
    const data = JSON.parse(res.getContentText());
    if (data && data[0]) {
      const s = data[0].summary;
      const b = data[0].onix;
      let cCode = (b && b.DescriptiveDetail && b.DescriptiveDetail.Subject) ? 
                  b.DescriptiveDetail.Subject.find(sub => sub.SubjectSchemeIdentifier === "78")?.SubjectCode : "";
      let price = (b && b.ProductSupply && b.ProductSupply.SupplyDetail && b.ProductSupply.SupplyDetail.Price) ?
                  b.ProductSupply.SupplyDetail.Price[0].PriceAmount : "";
      let format = (b && b.DescriptiveDetail) ? b.DescriptiveDetail.ProductFormDetail : "";
      let pages = (b && b.DescriptiveDetail && b.DescriptiveDetail.Extent) ? 
                  b.DescriptiveDetail.Extent.find(e => e.ExtentType === "00")?.ExtentValue : "";
      
      return { author: s.author, title: s.title, volume: s.volume, publisher: s.publisher, date: s.pubdate, 
               cCode: cCode, price: price, format: format, pages: pages, source: "OpenBD", subTitle: s.subtitle || "" };
    }
  } catch(e) {}
  return null;
}

function fetchGoogleBooks(isbn) {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('GOOGLE_BOOKS_API_KEY');
  try {
    const url = `https://www.googleapis.com/books/v1/volumes?q=isbn:${isbn}&country=JP${apiKey ? '&key=' + apiKey : ''}`;
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() === 200) {
      const data = JSON.parse(res.getContentText());
      if (data.totalItems > 0) {
        const v = data.items[0].volumeInfo;
        const s = data.items[0].saleInfo;
        let price = (s && s.listPrice) ? s.listPrice.amount : "";
        let pages = v.pageCount || "";
        return { author: (v.authors || []).join(","), title: v.title, volume: "", publisher: v.publisher || "", 
                 date: v.publishedDate || "", cCode: "", price: price, format: "", pages: pages, source: "Google", subTitle: v.subtitle || "" };
      }
    }
  } catch(e) {}
  return null;
}

function fetchNDL(isbn) {
  try {
    const url = `https://ndlsearch.ndl.go.jp/api/sru?operation=searchRetrieve&query=isbn%3D%22${isbn}%22`;
    const xml = UrlFetchApp.fetch(url).getContentText();
    const doc = XmlService.parse(xml);
    
    const nsSrw = XmlService.getNamespace("http://www.loc.gov/zing/srw/");
    const nsRecords = doc.getRootElement().getChild("records", nsSrw);
    if (!nsRecords) return null;
    
    const record = nsRecords.getChild("record", nsSrw);
    if (record) {
      const recordData = record.getChild("recordData", nsSrw);
      if (!recordData) return null;

      const dcData = XmlService.parse(recordData.getValue());
      const dcRoot = dcData.getRootElement();
      
      const nsDc = XmlService.getNamespace("dc", "http://purl.org/dc/elements/1.1/");
      const nsXsi = XmlService.getNamespace("xsi", "http://www.w3.org/2001/XMLSchema-instance");

      let ndc = "";
      const subjects = dcRoot.getChildren("subject", nsDc);
      for (let s of subjects) {
        let typeAttr = s.getAttribute("type", nsXsi);
        if (typeAttr && typeAttr.getValue().toLowerCase().indexOf("ndc") !== -1) {
          ndc = s.getValue();
          break;
        }
      }

      let titleText = dcRoot.getChildText("title", nsDc) || "";
      let volumeText = dcRoot.getChildText("volume", nsDc) || "";
      
      return {
        author: dcRoot.getChildText("creator", nsDc) || "",
        title: titleText,
        volume: volumeText,
        publisher: dcRoot.getChildText("publisher", nsDc) || "",
        date: dcRoot.getChildText("date", nsDc) || "",
        cCode: ndc,
        price: "", 
        format: "", 
        pages: "", 
        source: "NDL",
        subTitle: ""
      };
    }
  } catch(e) {}
  return null;
}

/**
 * 汎用データクレンジング (書名とVolを一体化)
 */
function cleanData(data) {
  if (data.author) {
    data.author = data.author.replace(/[,\.]/g, "").replace(/[\(\[\（\【](.*?)[\]\)\】\〕]/g, "").replace(/\d{4}-\d*/g, "").trim();
  }

  let fullTitle = data.title || "";
  let extractedVolume = data.volume || "";
  let extractedSub = data.subTitle || "";

  if (data.publisher && data.publisher.indexOf("角川") !== -1) {
    data.publisher = "角川書店";
  }

  if (extractedVolume && /^\d{3,}$/.test(extractedVolume)) {
    extractedVolume = "";
  }

  // 1. タイトルに含まれる記号によるサブタイトルの抽出
  const titleSeparators = [ /[:：]\s*/, /\s{2,}/ ];
  for (let sep of titleSeparators) {
    if (sep.test(fullTitle)) {
      let parts = fullTitle.split(sep);
      fullTitle = parts[0].trim();
      if (!extractedSub) {
        extractedSub = parts.slice(1).join(" ").trim();
      }
      break;
    }
  }

// 2. タイトル末尾から巻数を抽出
  const volPatterns = [
    / [(\（\s]第?(\d+|[上下中前後次後]|[ⅠⅡⅢⅣⅤ]|[\u2160-\u216F])[巻巻]?[\)\）]?$/,
    /\s(\d+|[上下中前後次後]|[ⅠⅡⅢⅣⅤ])$/,
    /[第](\d+|[上下中前後次後]|[壱弐参])[巻巻]$/,
    /[\s・]+([上下中前後次後])$/
  ];

  for (let pattern of volPatterns) {
    let match = fullTitle.match(pattern);
    if (match) {
      if (!extractedVolume) extractedVolume = match[1];
      fullTitle = fullTitle.replace(pattern, "").trim();
      break;
    }
  }

  // --------------------------------------------------
  // 3. 巻数（Vol）の漢数字→アラビア数字変換処理（既存処理）
  // --------------------------------------------------
  if (extractedVolume) {
    extractedVolume = extractedVolume.replace(/^第/, "").replace(/巻$/, "").replace(/[\(\)\（\）]/g, "").trim();

    if (/[一二三四五六七八九十百千]/.test(extractedVolume)) {
      const kanjiNums = { '一':1, '二':2, '三':3, '四':4, '五':5, '六':6, '七':7, '八':8, '九':9 };
      const kanjiUnits = { '十':10, '百':100, '千':1000, '万':10000 };

      let total = 0;
      let r = 0;
      let isValid = true;

      for (let i = 0; i < extractedVolume.length; i++) {
        let char = extractedVolume[i];
        if (kanjiNums[char] !== undefined) {
          r = kanjiNums[char];
        } else if (kanjiUnits[char] !== undefined) {
          if (r === 0) r = 1;
          total += r * kanjiUnits[char];
          r = 0;
        } else {
          isValid = false;
          break;
        }
      }
      if (isValid) {
        total += r;
        extractedVolume = String(total);
      }
    }
  }

  // --------------------------------------------------
  // 4. 書名とVolを一体化（書名 + 半角スペース + Vol）
  // --------------------------------------------------
  if (extractedVolume) {
    const volStr = String(extractedVolume).trim();
    if (!fullTitle.endsWith(volStr)) {
      fullTitle = fullTitle ? `${fullTitle} ${volStr}` : volStr;
    }
  }

  // クレンジング結果を格納
  data.title = fullTitle;
  data.subTitle = extractedSub; // サブタイトルは既存どおり保持
  data.volume = "";             // Volは書名に吸収されたためクリア

  return data;
}
/**
 * ISBN13桁(978...)をISBN10桁に変換
 */
function convertISBN13to10(isbn13) {
  const cleaned = isbn13.replace(/[-\s]/g, "");
  if (cleaned.length !== 13 || !cleaned.startsWith("978")) return isbn13;
  
  const core = cleaned.substring(3, 12);
  let sum = 0;
  for (let i = 0; i < 9; i++) {
    sum += parseInt(core[i], 10) * (10 - i);
  }
  let remainder = sum % 11;
  let checkDigit = 11 - remainder;
  
  if (checkDigit === 10) checkDigit = "X";
  else if (checkDigit === 11) checkDigit = "0";
  return core + checkDigit;
}

/**
 * ISBN10桁をISBN13桁(978...)に変換
 */
function convertISBN10to13(isbn10) {
  const cleaned = isbn10.replace(/[-\s]/g, "");
  if (cleaned.length !== 10) return isbn10;
  
  const core = "978" + cleaned.substring(0, 9);
  let sum = 0;
  for (let i = 0; i < 12; i++) {
    sum += parseInt(core[i], 10) * (i % 2 === 0 ? 1 : 3);
  }
  let remainder = sum % 10;
  let checkDigit = (10 - remainder) % 10;
  
  return core + String(checkDigit);
}