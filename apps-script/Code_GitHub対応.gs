const RECEIPT_APP = {
  TZ: 'Asia/Tokyo',
  ROOT_FOLDER: '家計簿レシート',
  PAYERS: {
    'ゆみこ': { dateCol: 4, amountCol: 5, descCol: 6 }, // D/E/F
    '翔平':   { dateCol: 8, amountCol: 9, descCol: 10 } // H/I/J
  }
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📸 レシート')
    .addItem('初期設定', 'setupReceiptApp')
    .addItem('登録ページURLを表示', 'showReceiptAppUrl')
    .addToUi();
}

function setupReceiptApp() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('このコードは家計簿スプレッドシートに紐づけて使ってください。');

  PropertiesService.getScriptProperties().setProperty('HOUSEHOLD_SHEET_ID', ss.getId());
  getOrCreateRootFolder_();

  SpreadsheetApp.getUi().alert(
    '初期設定完了',
    '家計簿との接続設定が完了しました。\n次に Apps Script を「ウェブアプリ」としてデプロイしてください。',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function showReceiptAppUrl() {
  const url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert(
    'レシート登録ページ',
    url || 'まだウェブアプリとしてデプロイされていません。',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('レシート登録')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 画像を保存し、Google Drive OCRで日付・合計金額・店名を抽出する。
 * HTML側からは form 要素をそのまま渡す。
 */
function analyzeReceipt(formObject) {
  if (!formObject || !formObject.receipt) throw new Error('レシート画像を選択してください。');

  const payer = String(formObject.payer || 'ゆみこ');
  if (!RECEIPT_APP.PAYERS[payer]) throw new Error('支払者が不正です。');

  const blob = formObject.receipt;
  const originalName = blob.getName() || 'receipt.jpg';
  const savedName = Utilities.formatDate(new Date(), RECEIPT_APP.TZ, 'yyyyMMdd_HHmmss') + '_' + originalName;
  blob.setName(savedName);

  const folder = getOrCreateMonthFolder_(new Date());
  const receiptFile = folder.createFile(blob);

  let ocrText = '';
  let ocrWarning = '';
  try {
    ocrText = ocrBlob_(blob);
  } catch (err) {
    ocrWarning = 'OCR読取に失敗しました。日付・金額・内容を手入力して登録できます。';
    console.error(err);
  }

  const parsed = parseReceipt_(ocrText);
  const today = Utilities.formatDate(new Date(), RECEIPT_APP.TZ, 'yyyy-MM-dd');

  return {
    payer: payer,
    date: parsed.date || today,
    amount: parsed.amount || '',
    description: parsed.description || '',
    receiptFileId: receiptFile.getId(),
    receiptUrl: receiptFile.getUrl(),
    warning: ocrWarning || parsed.warning || '',
    detectedSheet: getMonthSheetName_(parsed.date || today)
  };
}

/**
 * 確認済みデータを該当月シートへ登録。
 */
function saveReceipt(data) {
  if (!data) throw new Error('登録データがありません。');

  const payer = String(data.payer || '');
  const map = RECEIPT_APP.PAYERS[payer];
  if (!map) throw new Error('支払者を選択してください。');

  const dateString = String(data.date || '').trim();
  const amount = Number(String(data.amount || '').replace(/,/g, ''));
  const description = String(data.description || '').trim();

  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateString)) throw new Error('日付を確認してください。');
  if (!Number.isFinite(amount) || amount <= 0) throw new Error('合計金額を確認してください。');
  if (!description) throw new Error('店名・内容を入力してください。');

  const ss = getHouseholdSpreadsheet_();
  const sheetName = getMonthSheetName_(dateString);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('月タブ「' + sheetName + '」が見つかりません。先に月タブを作成してください。');

  const totalRow = findTotalRow_(sheet);
  const row = findOrCreateEmptyRow_(sheet, totalRow, map);

  const [y, m, d] = dateString.split('-').map(Number);
  const dateObj = new Date(y, m - 1, d, 12, 0, 0);

  sheet.getRange(row, map.dateCol).setValue(dateObj).setNumberFormat('M/d');
  sheet.getRange(row, map.amountCol).setValue(amount).setNumberFormat('#,##0');
  const descCell = sheet.getRange(row, map.descCol);
  descCell.setValue(description);

  if (data.receiptUrl) {
    descCell.setNote('レシート画像: ' + data.receiptUrl);
  }

  SpreadsheetApp.flush();

  return {
    ok: true,
    sheetName: sheetName,
    row: row,
    message: sheetName + ' の ' + row + '行目に登録しました。'
  };
}

function getHouseholdSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('HOUSEHOLD_SHEET_ID');
  if (!id) throw new Error('先にスプレッドシート上の「📸 レシート → 初期設定」を実行してください。');
  return SpreadsheetApp.openById(id);
}

function getMonthSheetName_(dateString) {
  let date;
  if (dateString instanceof Date) {
    date = dateString;
  } else {
    const m = String(dateString).match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) throw new Error('日付形式が不正です。');
    date = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0);
  }
  const yy = Utilities.formatDate(date, RECEIPT_APP.TZ, 'yy');
  const month = Number(Utilities.formatDate(date, RECEIPT_APP.TZ, 'M'));
  return yy + month; // 2026/9 => 269, 2026/10 => 2610
}

function findTotalRow_(sheet) {
  const maxRows = Math.max(sheet.getLastRow(), 40);
  const formulasE = sheet.getRange(1, 5, maxRows, 1).getFormulas();
  for (let i = 0; i < formulasE.length; i++) {
    const f = formulasE[i][0] || '';
    if (/^=SUM\(E\d+:E\d+\)$/i.test(f) || /^=SUM\(E2:E/i.test(f)) return i + 1;
  }

  const formulasI = sheet.getRange(1, 9, maxRows, 1).getFormulas();
  for (let i = 0; i < formulasI.length; i++) {
    const f = formulasI[i][0] || '';
    if (/^=SUM\(I\d+:I\d+\)$/i.test(f) || /^=SUM\(I2:I/i.test(f)) return i + 1;
  }

  const diffValues = sheet.getRange(1, 7, maxRows, 1).getDisplayValues();
  for (let i = 0; i < diffValues.length; i++) {
    if (String(diffValues[i][0]).trim() === '差額') return Math.max(3, i - 1);
  }

  return Math.max(sheet.getLastRow() + 1, 33);
}

function findOrCreateEmptyRow_(sheet, totalRow, map) {
  const startRow = 2;
  const count = Math.max(0, totalRow - startRow);
  if (count > 0) {
    const dates = sheet.getRange(startRow, map.dateCol, count, 1).getValues();
    const amounts = sheet.getRange(startRow, map.amountCol, count, 1).getValues();
    const descs = sheet.getRange(startRow, map.descCol, count, 1).getValues();

    for (let i = 0; i < count; i++) {
      if (dates[i][0] === '' && amounts[i][0] === '' && descs[i][0] === '') return startRow + i;
    }
  }

  // 入力欄が埋まった場合は合計行の直前に1行追加し、見た目だけ前行からコピー。
  sheet.insertRowBefore(totalRow);
  if (totalRow > 2) {
    sheet.getRange(totalRow - 1, 1, 1, 10)
      .copyTo(sheet.getRange(totalRow, 1, 1, 10), { formatOnly: true });
    sheet.getRange(totalRow, 1, 1, 10).clearContent();
  }
  return totalRow;
}

function getOrCreateRootFolder_() {
  const props = PropertiesService.getScriptProperties();
  const existingId = props.getProperty('RECEIPT_ROOT_FOLDER_ID');
  if (existingId) {
    try { return DriveApp.getFolderById(existingId); } catch (e) {}
  }

  const folders = DriveApp.getFoldersByName(RECEIPT_APP.ROOT_FOLDER);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(RECEIPT_APP.ROOT_FOLDER);
  props.setProperty('RECEIPT_ROOT_FOLDER_ID', folder.getId());
  return folder;
}

function getOrCreateMonthFolder_(date) {
  const root = getOrCreateRootFolder_();
  const name = Utilities.formatDate(date, RECEIPT_APP.TZ, 'yyyy-MM');
  const folders = root.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : root.createFolder(name);
}

/**
 * Google Drive API の画像→Googleドキュメント変換を利用したOCR。
 * Apps Script 左メニュー「サービス」から Google Drive API を追加しておく。
 */
function ocrBlob_(blob) {
  const resource = {
    name: 'OCR_' + Utilities.getUuid(),
    mimeType: 'application/vnd.google-apps.document'
  };

  const tempDoc = Drive.Files.create(resource, blob, {
    ocrLanguage: 'ja',
    fields: 'id'
  });

  try {
    let text = '';
    let lastError = null;
    for (let i = 0; i < 4; i++) {
      try {
        Utilities.sleep(700 + i * 400);
        text = DocumentApp.openById(tempDoc.id).getBody().getText();
        if (text && text.trim()) break;
      } catch (err) {
        lastError = err;
      }
    }
    if (!text.trim() && lastError) throw lastError;
    return text || '';
  } finally {
    try { Drive.Files.remove(tempDoc.id); } catch (e) { console.warn(e); }
  }
}

function parseReceipt_(rawText) {
  const text = normalizeText_(rawText || '');
  const lines = text.split(/\r?\n/).map(s => s.trim()).filter(Boolean);

  const date = extractDate_(text);
  const amount = extractTotalAmount_(lines);
  const description = extractStoreName_(lines);

  const warnings = [];
  if (!amount) warnings.push('合計金額を自動判定できませんでした。');
  if (!description) warnings.push('店名を自動判定できませんでした。');

  return {
    date: date,
    amount: amount || '',
    description: description || '',
    warning: warnings.join(' ')
  };
}

function normalizeText_(text) {
  return String(text)
    .normalize('NFKC')
    .replace(/[￥]/g, '¥')
    .replace(/[‐‑–—−]/g, '-')
    .replace(/[ \t]+/g, ' ');
}

function extractDate_(text) {
  let m;
  const patterns = [
    /(20\d{2})[\/\.\-年]\s*(\d{1,2})[\/\.\-月]\s*(\d{1,2})(?:日)?/,
    /(20\d{2})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日/,
    /(\d{2})[\/\.\-](\d{1,2})[\/\.\-](\d{1,2})/
  ];

  for (const p of patterns) {
    m = text.match(p);
    if (!m) continue;
    let y = Number(m[1]);
    if (y < 100) y += 2000;
    const mo = Number(m[2]);
    const d = Number(m[3]);
    if (isValidDateParts_(y, mo, d)) return toIsoDate_(y, mo, d);
  }

  // 年がない「9月14日」などは現在年として扱う。
  m = text.match(/(?:^|\D)(\d{1,2})\s*月\s*(\d{1,2})\s*日/);
  if (m) {
    const now = new Date();
    const y = Number(Utilities.formatDate(now, RECEIPT_APP.TZ, 'yyyy'));
    const mo = Number(m[1]);
    const d = Number(m[2]);
    if (isValidDateParts_(y, mo, d)) return toIsoDate_(y, mo, d);
  }

  return '';
}

function isValidDateParts_(y, m, d) {
  const dt = new Date(y, m - 1, d, 12, 0, 0);
  return dt.getFullYear() === y && dt.getMonth() === m - 1 && dt.getDate() === d;
}

function toIsoDate_(y, m, d) {
  return y + '-' + String(m).padStart(2, '0') + '-' + String(d).padStart(2, '0');
}

function extractTotalAmount_(lines) {
  const priority = [
    /税込\s*合計/i,
    /お支払(?:い)?(?:額)?/i,
    /支払\s*合計/i,
    /総\s*合計/i,
    /合\s*計/i,
    /現\s*計/i,
    /grand\s*total/i,
    /total/i
  ];
  const exclude = /小計|消費税|税率|外税|内税|値引|割引|お預|預り|預かり|釣銭|おつり|change|ポイント/i;

  for (const key of priority) {
    for (const line of lines) {
      if (!key.test(line) || exclude.test(line)) continue;
      const nums = extractMoneyNumbers_(line);
      if (nums.length) return Math.max.apply(null, nums);
    }
  }

  // キーワード行で取れなければ「円」「¥」がある行から最大値を拾う。
  const candidates = [];
  for (const line of lines) {
    if (exclude.test(line)) continue;
    if (!/[¥円]/.test(line)) continue;
    candidates.push.apply(candidates, extractMoneyNumbers_(line));
  }
  const filtered = candidates.filter(n => n > 0 && n < 10000000);
  return filtered.length ? Math.max.apply(null, filtered) : 0;
}

function extractMoneyNumbers_(line) {
  const out = [];
  const re = /(?:¥\s*)?(\d{1,3}(?:,\d{3})+|\d{1,7})(?:\s*円)?/g;
  let m;
  while ((m = re.exec(line)) !== null) {
    const n = Number(m[1].replace(/,/g, ''));
    if (Number.isFinite(n)) out.push(n);
  }
  return out;
}

function extractStoreName_(lines) {
  const reject = /領収書|レシート|receipt|tel|電話|fax|〒|住所|登録番号|適格請求書|インボイス|日時|日付|年月日|レジ|担当|取引|伝票|合計|小計|税|現金|クレジット|カード|お客様|毎度|ありがとう|営業時間|https?:\/\/|www\./i;

  for (const raw of lines.slice(0, 15)) {
    const line = raw.replace(/^[*\-=\s]+|[*\-=\s]+$/g, '').trim();
    if (!line || line.length < 2 || line.length > 36) continue;
    if (reject.test(line)) continue;
    if (/^\d[\d\s\-:/.]+$/.test(line)) continue;
    if (/^\d{2,4}-\d{2,4}-\d{3,4}$/.test(line)) continue;
    if (!/[A-Za-zぁ-んァ-ヶ一-龠々]/.test(line)) continue;
    return line;
  }
  return '';
}
