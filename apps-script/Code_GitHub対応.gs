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
    detectedSheet: getMonthSheetName_(parsed.date || today),
    ocrText: String(ocrText || '').slice(0, 3000)
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
  const created = getOrCreateMonthSheet_(ss, dateString);
  const sheet = created.sheet;
  const sheetName = sheet.getName();

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
    message: (created.isNew ? '月タブ「' + sheetName + '」を新しく作成しました。\n' : '') + sheetName + ' の ' + row + '行目に登録しました。'
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
  const fallback = yy + month; // 2026/9 => 269, 2026/10 => 2610

  // 実在するタブ名（26/9, 26-9 など表記ゆれ）に合わせる
  try {
    const key = Number(yy) * 100 + month;
    const hit = getHouseholdSpreadsheet_().getSheets()
      .map(s => s.getName()).find(n => monthKey_(n) === key);
    if (hit) return hit;
  } catch (e) {}
  return fallback;
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
  const totalKeys = [
    /税込\s*合計/i, /お支払(?:い)?(?:額|金額)?/i, /支払\s*合計/i, /総\s*合計/i,
    /合\s*計/i, /現\s*計/i, /grand\s*total/i, /\btotal\b/i
  ];
  const exclude = /小計|消費税|税率|外税|内税|対象|値引|割引|預|釣|つり|change|ポイント|残高|クーポン/i;
  const depositRe = /預/;           // お預かり
  const changeRe = /釣|つり/;        // お釣り

  // 「お預かり」「お釣り」の金額を先に押さえておく（合計 = 預かり - 釣り の検算用）
  let deposit = 0, change = -1;
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (depositRe.test(line) && !deposit) {
      const n = firstNumberFrom_(lines, i);
      if (n > 0) deposit = n;
    }
    if (changeRe.test(line) && change < 0) {
      const n = firstNumberFrom_(lines, i, true);
      if (n >= 0) change = n;
    }
  }
  const derived = (deposit > 0 && change >= 0 && deposit - change > 0) ? deposit - change : 0;

  // 合計キーワード行（同じ行に数字がなければ次の行を見る）
  for (const key of totalKeys) {
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      if (!key.test(line) || exclude.test(line)) continue;
      const n = firstNumberFrom_(lines, i);
      if (n > 0) {
        // 預かり-釣り と一致するならそれで確定。矛盾したら預かり額そのものは避ける
        if (derived && n === deposit) return derived;
        return n;
      }
    }
  }

  if (derived) return derived;

  // 列ごとに読まれた場合（項目名と金額が別行）：「合計・預かり・釣り」の並びを数字列から探す
  const seq = [];
  for (const line of lines) seq.push.apply(seq, extractMoneyNumbers_(line));
  for (let i = 0; i + 2 < seq.length; i++) {
    if (seq[i] > 0 && seq[i + 1] > seq[i] && seq[i] === seq[i + 1] - seq[i + 2]) return seq[i];
  }

  // 最後の手段：除外語のない「¥」「円」行の最大値
  const candidates = [];
  for (const line of lines) {
    if (exclude.test(line)) continue;
    if (!/[¥円]/.test(line)) continue;
    candidates.push.apply(candidates, extractMoneyNumbers_(line));
  }
  const filtered = candidates.filter(n => n > 0 && n < 10000000);
  return filtered.length ? Math.max.apply(null, filtered) : 0;
}

/** i行目のキーワード以降の数字。なければ次の1〜2行から拾う（次行が別キーワード行なら拾わない） */
function firstNumberFrom_(lines, i, allowZero) {
  const keyword = /合計|計|預|釣|つり|支払|total/i;
  const stripped = lines[i].replace(/^.*?(合\s*計|計|預かり|預り|釣り?|つり|支払い?|total)/i, '');
  let nums = extractMoneyNumbers_(stripped);
  if (!nums.length) nums = extractMoneyNumbers_(lines[i]);
  nums = nums.filter(n => allowZero ? n >= 0 : n > 0);
  if (nums.length) return nums[nums.length - 1];
  for (let k = 1; k <= 2 && i + k < lines.length; k++) {
    const next = lines[i + k];
    if (keyword.test(next)) break;
    const nn = extractMoneyNumbers_(next).filter(n => allowZero ? n >= 0 : n > 0);
    if (nn.length) return nn[0];
  }
  return allowZero ? -1 : 0;
}

function extractMoneyNumbers_(line) {
  const out = [];
  const cleaned = String(line)
    .replace(/\d{1,2}:\d{2}/g, ' ')                 // 時刻
    .replace(/\d{2,4}[-\/.]\d{1,2}[-\/.]\d{1,4}/g, ' ')  // 日付
    .replace(/\d{2,4}-\d{2,4}-\d{3,4}/g, ' ');       // 電話番号
  const re = /(?:¥\s*)?(\d{1,3}(?:,\d{3})+|\d{1,7})(?:\s*円)?/g;
  let m;
  while ((m = re.exec(cleaned)) !== null) {
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

/* =========================================================
 * 閲覧機能（月サマリー・登録一覧・余り金残高）
 * ========================================================= */

/** 月タブ一覧（新しい順）と、初期表示する月を返す */
function listMonthSheets() {
  const ss = getHouseholdSpreadsheet_();
  const all = ss.getSheets().map(s => s.getName());
  const names = all.filter(n => monthKey_(n) > 0).sort((a, b) => monthKey_(b) - monthKey_(a));
  if (!names.length) {
    throw new Error('月タブが見つかりません。タブ名: ' + all.join(', '));
  }
  const today = getMonthSheetName_(new Date());
  return { sheets: names, current: names.indexOf(today) >= 0 ? today : names[0] };
}

/**
 * タブ名 → 年月キー（例: 269 / 26/9 / 26-9 / 2026年9月 / ２６９ → 2609）
 * 月タブでなければ 0。
 */
function monthKey_(name) {
  let n = String(name).normalize('NFKC').trim();
  let m = n.match(/^(20)?(\d{2})\s*[\/\.\-_年]?\s*(\d{1,2})\s*月?$/);
  if (!m) return 0;
  const yy = Number(m[2]), mo = Number(m[3]);
  if (mo < 1 || mo > 12) return 0;
  return yy * 100 + mo;
}

/** 指定月タブのサマリーを返す */
function getMonthSummary(sheetName) {
  const ss = getHouseholdSpreadsheet_();
  const sheet = ss.getSheetByName(String(sheetName));
  if (!sheet) throw new Error('月タブ「' + sheetName + '」が見つかりません。');

  const totalRow = findTotalRow_(sheet);
  const lastRow = Math.max(sheet.getLastRow(), totalRow + 3);
  const values = sheet.getRange(1, 1, lastRow, 10).getValues();
  const displays = sheet.getRange(1, 1, lastRow, 10).getDisplayValues();

  // 左側（A/B列）：予算・固定費・余り・翔平受け取り
  const budget = toNum_(values[1][1]); // B2
  const fixed = [];
  let leftover = null, shoheiReceive = null;
  for (let r = 2; r < values.length; r++) {
    const label = String(values[r][0] || '').trim();
    const v = values[r][1];
    if (!label) continue;
    if (label === '余り' || label === '余り1') { if (leftover === null) leftover = toNum_(v); continue; }
    if (label === '翔平受け取り') { shoheiReceive = toNum_(v); continue; }
    if (r <= 9 && v !== '') fixed.push({ label: label, amount: toNum_(v) });
  }
  if (shoheiReceive === null && leftover !== null) shoheiReceive = budget - leftover;

  // 合計・差額
  const yumikoTotal = toNum_(values[totalRow - 1][4]); // E
  const shoheiTotal = toNum_(values[totalRow - 1][8]); // I
  let diff = null;
  for (let r = 0; r < values.length; r++) {
    if (String(values[r][6]).trim() === '差額') { diff = toNum_(values[r][8]); break; }
  }
  if (diff === null) diff = shoheiTotal - yumikoTotal;

  // 明細
  const entries = [];
  const pushEntry = (payer, dCol, aCol, tCol) => {
    for (let r = 1; r < totalRow - 1; r++) {
      const amt = values[r][aCol];
      if (amt === '' || amt === null) continue;
      entries.push({
        payer: payer,
        row: r + 1,
        date: displays[r][dCol] || '',
        sortKey: values[r][dCol] instanceof Date ? values[r][dCol].getTime() : 0,
        amount: toNum_(amt),
        description: String(values[r][tCol] || '')
      });
    }
  };
  pushEntry('ゆみこ', 3, 4, 5);
  pushEntry('翔平', 7, 8, 9);
  entries.sort((a, b) => b.sortKey - a.sortKey);

  return {
    sheetName: String(sheetName),
    budget: budget,
    fixed: fixed,
    leftover: leftover,          // 貯金に回せる額（余り）
    shoheiReceive: shoheiReceive,
    yumikoTotal: yumikoTotal,
    shoheiTotal: shoheiTotal,
    diff: diff,                  // +なら翔平が多く払っている
    entries: entries,
    savings: getSavingsBalance_(ss)
  };
}

/** 「余り金」タブの累計残金（最新月）を返す */
function getSavingsBalance_(ss) {
  const sheet = ss.getSheetByName('余り金');
  if (!sheet) return null;
  const last = sheet.getLastRow();
  // M〜P列：M=年, N=月, O=その月の余り, P=累計残金
  const vals = sheet.getRange(1, 13, last, 4).getValues();
  let year = '', result = null;
  for (let r = 0; r < vals.length; r++) {
    if (vals[r][0] !== '' && vals[r][0] !== null) year = String(vals[r][0]).replace(/\.0$/, '');
    if (vals[r][1] !== '' && vals[r][2] !== '' && vals[r][2] !== null) {
      result = { year: year, month: String(vals[r][1]), balance: toNum_(vals[r][3]) };
    }
  }
  return result;
}

function toNum_(v) {
  if (v === '' || v === null || v === undefined) return 0;
  const n = Number(String(v).replace(/,/g, ''));
  return Number.isFinite(n) ? n : 0;
}


/* =========================================================
 * 月タブの自動作成（前月タブをコピー）
 * ========================================================= */

/** 日付に対応する月タブを返す。無ければ直近の月タブをコピーして作る。 */
function getOrCreateMonthSheet_(ss, dateString) {
  const m = String(dateString).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) throw new Error('日付形式が不正です。');
  const year = Number(m[1]), month = Number(m[2]);
  const key = (year % 100) * 100 + month;

  const sheets = ss.getSheets();
  const existing = sheets.find(sh => monthKey_(sh.getName()) === key);
  if (existing) return { sheet: existing, isNew: false };

  // コピー元：対象月より前で一番新しい月タブ（無ければ全体で最新）
  const monthSheets = sheets.filter(sh => monthKey_(sh.getName()) > 0);
  if (!monthSheets.length) throw new Error('コピー元になる月タブがありません。');
  const before = monthSheets.filter(sh => monthKey_(sh.getName()) < key);
  const pool = before.length ? before : monthSheets;
  const src = pool.reduce((a, b) => monthKey_(a.getName()) >= monthKey_(b.getName()) ? a : b);

  const newName = formatMonthName_(src.getName(), year, month);
  const sheet = src.copyTo(ss).setName(newName);
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(src.getIndex() + 1);

  clearMonthEntries_(sheet);
  try { linkSavingsTab_(ss, sheet.getName(), year, month); } catch (e) { console.warn(e); }

  return { sheet: sheet, isNew: true };
}

/** コピー元タブ名の表記（26/9, 269, 2026年9月 …）に合わせて新しい名前を作る */
function formatMonthName_(srcName, year, month) {
  const n = String(srcName).normalize('NFKC').trim();
  const yy = String(year % 100).padStart(2, '0');
  let mm = n.match(/^(20)?(\d{2})(\s*[\/\.\-_年]\s*)(\d{1,2})(\s*月?)$/);
  if (mm) {
    const prefix = mm[1] ? '20' : '';
    const two = /^\d{2}$/.test(mm[4]);
    return prefix + yy + mm[3] + (two ? String(month).padStart(2, '0') : String(month)) + mm[5];
  }
  return yy + month; // 269 形式
}

/** 明細・固定費の入力欄を空にする（数式・書式は残す） */
function clearMonthEntries_(sheet) {
  const totalRow = findTotalRow_(sheet);
  const rows = Math.max(0, totalRow - 2);
  if (rows > 0) {
    sheet.getRange(2, 4, rows, 3).clearContent().clearNote(); // D:F ゆみこ
    sheet.getRange(2, 8, rows, 3).clearContent().clearNote(); // H:J 翔平
  }
  // B3:B12 の固定費（数式でないセルだけ空に。予算B2は残す）
  const rng = sheet.getRange(3, 2, 10, 1);
  const formulas = rng.getFormulas();
  const values = rng.getValues();
  const labels = sheet.getRange(3, 1, 10, 1).getValues();
  for (let i = 0; i < 10; i++) {
    const label = String(labels[i][0] || '').trim();
    if (!formulas[i][0] && values[i][0] !== '' && label && !/余り|受け取り/.test(label)) {
      sheet.getRange(3 + i, 2).clearContent();
    }
  }
  SpreadsheetApp.flush();
}

/** 「余り金」タブの該当月に ='タブ名'!B9 を入れる（未入力の場合のみ） */
function linkSavingsTab_(ss, sheetName, year, month) {
  const sv = ss.getSheetByName('余り金');
  if (!sv) return;
  const last = Math.max(sv.getLastRow(), 6);
  const vals = sv.getRange(1, 13, last, 4).getValues();     // M:P
  const forms = sv.getRange(1, 13, last, 4).getFormulas();
  const ref = "='" + sheetName.replace(/'/g, "''") + "'!B9";

  let curYear = null, targetRow = 0, lastRow = 0;
  for (let r = 0; r < vals.length; r++) {
    if (vals[r][0] !== '' && vals[r][0] !== null) curYear = Number(vals[r][0]);
    const monthCell = String(vals[r][1] || '').trim();
    if (monthCell) lastRow = r + 1;
    if (curYear === year && monthCell === month + '月') { targetRow = r + 1; break; }
  }

  if (targetRow) {
    if (!forms[targetRow - 1][2] && (vals[targetRow - 1][2] === '' || vals[targetRow - 1][2] === null)) {
      sv.getRange(targetRow, 15).setFormula(ref);
    }
    return;
  }

  // 行が無ければ末尾に追加（年が変わる場合はM列に年も入れる）
  const row = lastRow + 1;
  if (curYear !== year || month === 1) sv.getRange(row, 13).setValue(year);
  sv.getRange(row, 14).setValue(month + '月');
  sv.getRange(row, 15).setFormula(ref);
  sv.getRange(row, 16).setFormula('=P' + (row - 1) + '+O' + row);
}


/** 明細を1件削除（該当行の日付・金額・内容を空にする。行自体は残す） */
function deleteEntry(data) {
  if (!data || !data.sheetName || !data.payer || !data.row) throw new Error('削除対象が不正です。');
  const map = RECEIPT_APP.PAYERS[String(data.payer)];
  if (!map) throw new Error('支払者が不正です。');
  const row = Number(data.row);
  if (!Number.isFinite(row) || row < 2) throw new Error('行番号が不正です。');

  const ss = getHouseholdSpreadsheet_();
  const sheet = ss.getSheetByName(String(data.sheetName));
  if (!sheet) throw new Error('月タブ「' + data.sheetName + '」が見つかりません。');

  const totalRow = findTotalRow_(sheet);
  if (row >= totalRow) throw new Error('合計行より下は削除できません。');

  // 念のため内容が一致するか確認
  const cur = sheet.getRange(row, map.amountCol).getValue();
  if (data.amount !== undefined && Number(cur) !== Number(data.amount)) {
    throw new Error('データが変わっています。更新してからやり直してください。');
  }

  sheet.getRange(row, map.dateCol, 1, 3).clearContent().clearNote();
  SpreadsheetApp.flush();
  return { ok: true, message: '削除しました。' };
}
