const SHEET_NAME = '日報データ';
const CORE_HEADERS = [
  'reportId',
  'createdAt',
  'updatedAt',
  'confirmed',
  'confirmedBy',
  'confirmedAt',
  'staffName',
  'jobRole',
  'workPlaceType',
  'workDate',
  'storeName',
  'eventVenue',
  'eventOverallTarget',
  'eventVenueTarget',
  'photoCount',
  'photoUrls',
  'rawJson'
];
const PAYLOAD_PREFIX = 'payload.';

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || '{}');
    const scriptProps = PropertiesService.getScriptProperties();
    const token = scriptProps.getProperty('SYNC_TOKEN') || '';

    if (token && payload.token !== token) {
      return jsonOut({ ok: false, error: 'Unauthorized' });
    }

    if (payload.action === 'delete') {
      deleteReport(payload.reportId);
      return jsonOut({ ok: true, action: 'delete' });
    }

    if (payload.action === 'upsert' && payload.report) {
      upsertReport(payload.report);
      return jsonOut({ ok: true, action: 'upsert' });
    }

    if (payload.action === 'replaceAll' && Array.isArray(payload.reports)) {
      replaceAllReports(payload.reports);
      return jsonOut({ ok: true, action: 'replaceAll' });
    }

    if (payload.action === 'repairSheet') {
      const repaired = repairSheetFromCurrentRows();
      return jsonOut({ ok: true, action: 'repairSheet', repaired: repaired });
    }

    if (payload.action === 'uploadPhoto' && payload.dataUrl) {
      const uploaded = uploadPhotoToDrive(payload);
      return jsonOut({
        ok: true,
        action: 'uploadPhoto',
        photoFileId: uploaded.fileId,
        photoUrl: uploaded.photoUrl,
        webViewLink: uploaded.webViewLink,
        sharingEnabled: uploaded.sharingEnabled,
        warning: uploaded.warning
      });
    }

    return jsonOut({ ok: false, error: 'Invalid action' });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  try {
    const params = e.parameter || {};
    const scriptProps = PropertiesService.getScriptProperties();
    const token = scriptProps.getProperty('SYNC_TOKEN') || '';

    if (token && params.token !== token) {
      return jsonOut({ ok: false, error: 'Unauthorized' });
    }

    if (params.action === 'list') {
      return jsonOut({ ok: true, reports: listReports() });
    }

    return jsonOut({ ok: false, error: 'Invalid action' });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function jsonOut(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  ensureHeaderForReports_(sheet, []);
  return sheet;
}

function getHeader_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return [];
  const row = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  let end = row.length;
  while (end > 0 && String(row[end - 1] || '').trim() === '') end -= 1;
  return row.slice(0, end).map((v) => String(v || '').trim());
}

function getHeaderIndexMap_(header) {
  const map = {};
  for (let i = 0; i < header.length; i += 1) {
    if (!header[i]) continue;
    map[header[i]] = i;
  }
  return map;
}

function ensureHeaderForReports_(sheet, reports) {
  const current = getHeader_(sheet);
  const desired = composeDesiredHeader_(current, reports);
  if (desired.length === 0) return desired;

  if (sheet.getMaxColumns() < desired.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), desired.length - sheet.getMaxColumns());
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, desired.length).setValues([desired]);
    return desired;
  }

  if (!isSameHeader_(current, desired)) {
    sheet.getRange(1, 1, 1, desired.length).setValues([desired]);
  }
  return desired;
}

function composeDesiredHeader_(current, reports) {
  const dynamic = collectDynamicHeaders_(reports);

  if (!current || current.length === 0) {
    return CORE_HEADERS.concat(dynamic);
  }

  const out = current.slice();
  const seen = new Set(out);

  for (let i = 0; i < CORE_HEADERS.length; i += 1) {
    if (!seen.has(CORE_HEADERS[i])) {
      out.push(CORE_HEADERS[i]);
      seen.add(CORE_HEADERS[i]);
    }
  }
  for (let i = 0; i < dynamic.length; i += 1) {
    if (!seen.has(dynamic[i])) {
      out.push(dynamic[i]);
      seen.add(dynamic[i]);
    }
  }
  return out;
}

function collectDynamicHeaders_(reports) {
  const keys = new Set();
  const list = Array.isArray(reports) ? reports : [];
  for (let i = 0; i < list.length; i += 1) {
    const report = list[i] || {};
    const payload = report.payload || {};
    const flat = {};
    flattenPayload_(payload, PAYLOAD_PREFIX.slice(0, -1), flat);
    const names = Object.keys(flat);
    for (let j = 0; j < names.length; j += 1) {
      keys.add(names[j]);
    }
  }
  return Array.from(keys).sort();
}

function flattenPayload_(value, path, out) {
  if (value == null) return;

  if (Array.isArray(value)) {
    out[path] = JSON.stringify(value);
    return;
  }

  if (typeof value === 'object') {
    const keys = Object.keys(value);
    if (keys.length === 0) {
      out[path] = '{}';
      return;
    }
    for (let i = 0; i < keys.length; i += 1) {
      const key = keys[i];
      flattenPayload_(value[key], `${path}.${key}`, out);
    }
    return;
  }

  out[path] = value;
}

function buildCoreMap_(report) {
  const p = report.payload || {};
  const step1 = p.step1 || {};
  const photos = getPhotoList_(step1);
  const photoUrls = photos
    .map((item) => item.url || item.dataUrl || '')
    .filter(Boolean)
    .join('\n');

  return {
    reportId: report.id || '',
    createdAt: report.createdAt || '',
    updatedAt: report.updatedAt || '',
    confirmed: report.confirmed ? '確認済み' : '未確認',
    confirmedBy: report.confirmedBy || '',
    confirmedAt: report.confirmedAt || '',
    staffName: step1.staffName || '',
    jobRole: step1.jobRole || '',
    workPlaceType: step1.workPlaceType || '',
    workDate: step1.workDate || '',
    storeName: step1.storeName || '',
    eventVenue: step1.eventVenue || '',
    eventOverallTarget: step1.eventOverallTarget || '',
    eventVenueTarget: step1.eventVenueTarget || '',
    photoCount: photos.length,
    photoUrls: photoUrls,
    rawJson: JSON.stringify(report)
  };
}

function rowFromReportByHeader_(report, header) {
  const core = buildCoreMap_(report);
  const flat = {};
  flattenPayload_(report.payload || {}, PAYLOAD_PREFIX.slice(0, -1), flat);

  const row = new Array(header.length).fill('');
  for (let i = 0; i < header.length; i += 1) {
    const key = header[i];
    if (Object.prototype.hasOwnProperty.call(core, key)) {
      row[i] = core[key];
    } else if (Object.prototype.hasOwnProperty.call(flat, key)) {
      row[i] = flat[key];
    }
  }
  return row;
}

function findRowByReportId_(sheet, header, reportId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  const idCol = header.indexOf('reportId') + 1;
  const targetCol = idCol > 0 ? idCol : 1;
  const values = sheet.getRange(2, targetCol, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i += 1) {
    if (String(values[i][0] || '') === String(reportId || '')) {
      return i + 2;
    }
  }
  return -1;
}

function upsertReport(report) {
  if (!report || typeof report !== 'object') return;
  if (!report.id) {
    report.id = `rpt-${Date.now()}-${Math.floor(Math.random() * 100000)}`;
  }

  const sheet = getSheet_();
  const header = ensureHeaderForReports_(sheet, [report]);
  const rowData = rowFromReportByHeader_(report, header);
  const foundRow = findRowByReportId_(sheet, header, report.id);

  if (foundRow > 0) {
    sheet.getRange(foundRow, 1, 1, header.length).setValues([rowData]);
  } else {
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, header.length).setValues([rowData]);
  }
}

function deleteReport(reportId) {
  if (!reportId) return;
  const sheet = getSheet_();
  const header = getHeader_(sheet);
  const foundRow = findRowByReportId_(sheet, header, reportId);
  if (foundRow > 0) {
    sheet.deleteRow(foundRow);
  }
}

function replaceAllReports(reports) {
  const list = Array.isArray(reports) ? reports : [];
  const sheet = getSheet_();
  const header = composeDesiredHeader_(getHeader_(sheet), list);
  const rows = list.map((report) => rowFromReportByHeader_(report, header));

  sheet.clear();
  if (sheet.getMaxColumns() < header.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), header.length - sheet.getMaxColumns());
  }
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  }
}

function listReports() {
  const sheet = getSheet_();
  const header = getHeader_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2 || header.length === 0) return [];

  const rows = sheet.getRange(2, 1, lastRow - 1, header.length).getValues();
  const idx = getHeaderIndexMap_(header);
  const reports = [];

  for (let i = 0; i < rows.length; i += 1) {
    const row = rows[i];
    const raw = pickCell_(row, idx, ['rawJson', '生データJSON', '管理用:rawJson']);
    let report = {};

    if (raw) {
      try {
        report = JSON.parse(String(raw));
      } catch (err) {
        report = {};
      }
    }

    report = mergeCoreFromRow_(report, row, idx);
    report = mergePayloadFromDynamicColumns_(report, row, header, idx);
    if (report && report.id) reports.push(report);
  }
  return reports;
}

function pickCell_(row, indexMap, keys) {
  for (let i = 0; i < keys.length; i += 1) {
    const idx = indexMap[keys[i]];
    if (typeof idx === 'number') {
      const value = row[idx];
      if (value !== '' && value != null) return value;
    }
  }
  return '';
}

function mergeCoreFromRow_(report, row, idx) {
  const base = report && typeof report === 'object' ? report : {};
  const payload = base.payload && typeof base.payload === 'object' ? base.payload : {};
  const step1 = payload.step1 && typeof payload.step1 === 'object' ? payload.step1 : {};

  const id = String(pickCell_(row, idx, ['reportId', '日報ID']) || base.id || '').trim();
  if (!id) return null;

  const confirmedRaw = String(pickCell_(row, idx, ['confirmed', '確認状況']) || base.confirmed || '').trim();
  const confirmed = confirmedRaw === '確認済み' || confirmedRaw === 'true' || confirmedRaw === '1';
  const photoUrlsText = String(pickCell_(row, idx, ['photoUrls', '会場写真URL']) || '').trim();
  const photos = photoUrlsText
    ? photoUrlsText.split(/\n+/).map((url) => ({ name: '会場写真', type: 'image/jpeg', size: 0, url: String(url).trim(), dataUrl: '', fileId: '' })).filter((x) => x.url)
    : (Array.isArray(step1.photos) ? step1.photos : []);

  return {
    ...base,
    id: id,
    createdAt: formatDateTimeValue_(pickCell_(row, idx, ['createdAt']), base.createdAt || ''),
    updatedAt: formatDateTimeValue_(pickCell_(row, idx, ['updatedAt']), base.updatedAt || ''),
    confirmed: confirmed,
    confirmedBy: String(pickCell_(row, idx, ['confirmedBy']) || base.confirmedBy || ''),
    confirmedAt: formatDateTimeValue_(pickCell_(row, idx, ['confirmedAt']), base.confirmedAt || ''),
    payload: {
      ...payload,
      step1: {
        ...step1,
        staffName: String(pickCell_(row, idx, ['staffName']) || step1.staffName || ''),
        workPlaceType: String(pickCell_(row, idx, ['workPlaceType']) || step1.workPlaceType || ''),
        workDate: formatDateOnlyValue_(pickCell_(row, idx, ['workDate']), step1.workDate || ''),
        storeName: String(pickCell_(row, idx, ['storeName']) || step1.storeName || ''),
        eventVenue: String(pickCell_(row, idx, ['eventVenue']) || step1.eventVenue || ''),
        photos: photos
      }
    }
  };
}

function mergePayloadFromDynamicColumns_(report, row, header, idx) {
  if (!report) return report;
  const next = { ...report };
  const payload = next.payload && typeof next.payload === 'object' ? next.payload : {};
  next.payload = payload;

  for (let i = 0; i < header.length; i += 1) {
    const key = header[i];
    if (key.indexOf(PAYLOAD_PREFIX) !== 0) continue;
    const value = row[i];
    if (value === '' || value == null) continue;

    const path = key.slice(PAYLOAD_PREFIX.length);
    if (path === 'step1.workDate') {
      setByPath_(payload, path, formatDateOnlyValue_(value, ''));
      continue;
    }
    setByPath_(payload, path, parseCellValue_(value));
  }
  return next;
}

function setByPath_(obj, path, value) {
  if (!path) return;
  const keys = path.split('.');
  let cursor = obj;
  for (let i = 0; i < keys.length - 1; i += 1) {
    const k = keys[i];
    if (!cursor[k] || typeof cursor[k] !== 'object' || Array.isArray(cursor[k])) {
      cursor[k] = {};
    }
    cursor = cursor[k];
  }
  cursor[keys[keys.length - 1]] = value;
}

function parseCellValue_(value) {
  if (value instanceof Date) return formatDateTimeValue_(value, '');
  if (typeof value === 'number' || typeof value === 'boolean') return value;
  const text = String(value || '').trim();
  if (!text) return '';

  if (text === 'true') return true;
  if (text === 'false') return false;
  if (/^-?\d+(\.\d+)?$/.test(text)) return Number(text);
  if ((text[0] === '{' && text[text.length - 1] === '}') || (text[0] === '[' && text[text.length - 1] === ']')) {
    try {
      return JSON.parse(text);
    } catch (err) {
      return text;
    }
  }
  return text;
}

function formatDateOnlyValue_(value, fallback) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const text = String(value || '').trim();
  if (!text) return String(fallback || '');
  const matched = text.match(/^(\d{4}-\d{2}-\d{2})/);
  if (matched) return matched[1];
  return text;
}

function formatDateTimeValue_(value, fallback) {
  if (value instanceof Date) return value.toISOString();
  const text = String(value || '').trim();
  if (!text) return String(fallback || '');
  return text;
}

function repairSheetFromCurrentRows() {
  const reports = listReports();
  replaceAllReports(reports);
  return reports.length;
}

function isSameHeader_(a, b) {
  if (!a || !b || a.length !== b.length) return false;
  for (let i = 0; i < a.length; i += 1) {
    if (String(a[i] || '') !== String(b[i] || '')) return false;
  }
  return true;
}

function getPhotoList_(step1) {
  const photos = Array.isArray(step1.photos) ? step1.photos : [];
  if (photos.length > 0) return photos;
  const legacy = step1.photoUrl || step1.photoDataUrl || '';
  if (!legacy) return [];
  return [{ url: step1.photoUrl || '', dataUrl: step1.photoDataUrl || '' }];
}

function uploadPhotoToDrive(payload) {
  const parsed = parseDataUrl_(payload.dataUrl);
  const mimeType = parsed.mimeType || payload.mimeType || 'image/jpeg';
  const bytes = Utilities.base64Decode(parsed.base64Data);
  const ext = extensionFromMime_(mimeType);
  const fileName = sanitizeFileName_(payload.fileName || `photo-${Date.now()}.${ext}`);

  const blob = Utilities.newBlob(bytes, mimeType, fileName);
  const folder = getPhotoFolder_();
  const file = folder.createFile(blob);

  let sharingEnabled = false;
  let warning = '';
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    sharingEnabled = true;
  } catch (err) {
    warning = `setSharing failed: ${String(err)}`;
  }

  return {
    fileId: file.getId(),
    photoUrl: `https://drive.google.com/uc?export=view&id=${file.getId()}`,
    webViewLink: file.getUrl(),
    sharingEnabled: sharingEnabled,
    warning: warning
  };
}

function parseDataUrl_(dataUrl) {
  const source = String(dataUrl || '');
  const matched = source.match(/^data:([^;]+);base64,(.+)$/);
  if (!matched) throw new Error('Invalid dataUrl');
  return {
    mimeType: matched[1],
    base64Data: matched[2]
  };
}

function extensionFromMime_(mimeType) {
  if (!mimeType) return 'jpg';
  if (mimeType.indexOf('png') >= 0) return 'png';
  if (mimeType.indexOf('webp') >= 0) return 'webp';
  if (mimeType.indexOf('gif') >= 0) return 'gif';
  return 'jpg';
}

function sanitizeFileName_(name) {
  const cleaned = String(name || '')
    .replace(/[\\/:*?"<>|]/g, '_')
    .trim();
  if (!cleaned) return `photo-${Date.now()}.jpg`;
  return cleaned.length > 120 ? cleaned.slice(0, 120) : cleaned;
}

function getPhotoFolder_() {
  const scriptProps = PropertiesService.getScriptProperties();
  const folderId = scriptProps.getProperty('PHOTO_FOLDER_ID');
  if (folderId) {
    try {
      return DriveApp.getFolderById(folderId);
    } catch (err) {
      return DriveApp.getRootFolder();
    }
  }
  return DriveApp.getRootFolder();
}

// ===== Spreadsheet-only cleanup utilities (does not modify app behavior) =====
const READABLE_SHEET_NAME = '日報データ_見やすい';
const BACKUP_SHEET_PREFIX = '日報データ_backup_';
const STAFF_SHEET_NAME_SUFFIX = 'さん';
const STAFF_SUMMARY_SHEET_NAME_SUFFIX = 'さん_集計';
const STAFF_SYNC_CURSOR_KEY = 'STAFF_SYNC_CURSOR';
const STAFF_SYNC_BATCH_SIZE = 4;

const READABLE_HEADERS = [
  '日報ID',
  '作成日時',
  '更新日時',
  '確認状況',
  '確認者',
  '確認日時',
  'スタッフ名',
  '担当業務',
  '稼働日',
  '区分',
  '店舗名',
  'イベント会場',
  'イベント全体目標（店舗とイベント）',
  'イベント会場目標',
  '会場写真枚数',
  '会場写真URL',
  '来店数',
  'キャッチ数（反応数）',
  '着座数',
  '見込み',
  '着座内訳 au/UQ既存',
  '着座内訳 SB／ワイモバイル',
  '着座内訳 docomo／ahamo',
  '着座内訳 楽天',
  '着座内訳 その他',
  'au MNP SIM単',
  'au MNP HS',
  'au純新規 SIM単',
  'au純新規 HS',
  'UQ MNP SIM単',
  'UQ MNP HS',
  'UQ純新規 SIM単',
  'UQ純新規 HS',
  'セルアップ',
  'auでんき',
  'ゴールドカード',
  'シルバーカード',
  'ランクアップ',
  'じぶん銀行',
  'ノートン',
  'auひかり 新規',
  'auひかり ドコモ光から切替',
  'auひかり ソフトバンク光から切替',
  'auひかり その他から切替',
  'BLひかり 新規',
  'BLひかり ドコモ光から切替',
  'BLひかり ソフトバンク光から切替',
  'BLひかり その他から切替',
  'コミュファ光 新規',
  'コミュファ光 ドコモ光から切替',
  'コミュファ光 ソフトバンク光から切替',
  'コミュファ光 その他から切替',
  '成約事例 件数',
  '成約事例1 来店理由',
  '成約事例1 客層',
  '成約事例1 決め手トーク（タグ）',
  '成約事例1 決め手トーク（具体）',
  '成約事例1 成約要因',
  '成約事例1 その他',
  '改善事例 件数',
  '改善事例1 改善ポイント',
  '改善事例1 理由（具体）',
  '改善事例1 その他',
  'イベント会場の評価',
  'その他（会場評価）',
  '所感（短文）',
  'その他備考',
  '管理者コメント',
  '生データJSON'
];

/**
 * 実行手順:
 * 1) 元シート「日報データ」をそのままコピーしてバックアップ作成
 * 2) 「日報データ_見やすい」に中学生でも読みやすい列名で整形出力
 * 3) スタッフごとのシート（○○さん）を整形出力
 * この関数は日報アプリの同期シート「日報データ」の内容を変更しません。
 */
function runSpreadsheetCleanupForReadableView() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(SHEET_NAME);
  if (!rawSheet) throw new Error(`シートが見つかりません: ${SHEET_NAME}`);

  backupRawSheet_(ss, rawSheet);
  buildReadableSheet_(ss, rawSheet);
  buildStaffReadableSheets_(ss, rawSheet);
}

/**
 * 元データは変更せず、見やすいシートとスタッフ別シートだけを更新する。
 * 自動更新トリガーからはこの関数を呼ぶ。
 */
function refreshReadableSheetOnly() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(SHEET_NAME);
  if (!rawSheet) throw new Error(`シートが見つかりません: ${SHEET_NAME}`);
  buildReadableSheet_(ss, rawSheet);
  buildStaffReadableSheetsChunked_(ss, rawSheet, STAFF_SYNC_BATCH_SIZE);
}

/**
 * 10分ごとに「見やすいシート」だけ自動更新するトリガーを作成。
 * 既に同じトリガーがある場合は重複作成しない。
 */
function setupReadableAutoRefreshTrigger() {
  const fn = 'refreshReadableSheetOnly';
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some((t) => t.getHandlerFunction() === fn);
  if (exists) return;

  ScriptApp.newTrigger(fn)
    .timeBased()
    .everyMinutes(10)
    .create();
}

/**
 * 自動更新トリガーを削除（停止）する。
 */
function removeReadableAutoRefreshTrigger() {
  const fn = 'refreshReadableSheetOnly';
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((t) => {
    if (t.getHandlerFunction() === fn) {
      ScriptApp.deleteTrigger(t);
    }
  });
}

function backupRawSheet_(ss, rawSheet) {
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const backupName = `${BACKUP_SHEET_PREFIX}${stamp}`;
  const copied = rawSheet.copyTo(ss);
  copied.setName(backupName);
  ss.setActiveSheet(rawSheet);
}

function buildReadableSheet_(ss, rawSheet) {
  let readable = ss.getSheetByName(READABLE_SHEET_NAME);
  if (!readable) readable = ss.insertSheet(READABLE_SHEET_NAME);
  readable.clear();

  const lastRow = rawSheet.getLastRow();
  const lastCol = rawSheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) {
    readable.getRange(1, 1, 1, READABLE_HEADERS.length).setValues([READABLE_HEADERS]);
    return;
  }

  const header = rawSheet.getRange(1, 1, 1, lastCol).getValues()[0].map((v) => String(v || '').trim());
  const idx = {};
  for (let i = 0; i < header.length; i += 1) idx[header[i]] = i;

  const rows = lastRow >= 2 ? rawSheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  const outRows = rows.map((row) => readableRowFromSource_(row, idx));

  readable.getRange(1, 1, 1, READABLE_HEADERS.length).setValues([READABLE_HEADERS]);
  if (outRows.length > 0) {
    readable.getRange(2, 1, outRows.length, READABLE_HEADERS.length).setValues(outRows);
  }

  readable.setFrozenRows(1);
  readable.autoResizeColumns(1, READABLE_HEADERS.length);
}

function buildStaffReadableSheetsChunked_(ss, rawSheet, batchSize) {
  const props = PropertiesService.getScriptProperties();
  const start = Number(props.getProperty(STAFF_SYNC_CURSOR_KEY) || 0);
  const size = Math.max(1, Number(batchSize || STAFF_SYNC_BATCH_SIZE));
  const result = buildStaffReadableSheets_(ss, rawSheet, start, size);
  if (!result || result.total <= 0) {
    props.setProperty(STAFF_SYNC_CURSOR_KEY, '0');
    return;
  }
  const next = result.end >= result.total ? 0 : result.end;
  props.setProperty(STAFF_SYNC_CURSOR_KEY, String(next));
}

function refreshAllStaffSheetsNow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(SHEET_NAME);
  if (!rawSheet) throw new Error(`シートが見つかりません: ${SHEET_NAME}`);
  buildReadableSheet_(ss, rawSheet);
  buildStaffReadableSheets_(ss, rawSheet);
  PropertiesService.getScriptProperties().setProperty(STAFF_SYNC_CURSOR_KEY, '0');
}

function buildStaffReadableSheets_(ss, rawSheet, startIndex, batchSize) {
  const lastRow = rawSheet.getLastRow();
  const lastCol = rawSheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return { total: 0, start: 0, end: 0 };

  const header = rawSheet.getRange(1, 1, 1, lastCol).getValues()[0].map((v) => String(v || '').trim());
  const idx = {};
  for (let i = 0; i < header.length; i += 1) idx[header[i]] = i;

  const rows = rawSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const grouped = {};
  const groupedReports = {};

  for (let i = 0; i < rows.length; i += 1) {
    const row = rows[i];
    const report = reportFromSourceRow_(row, idx);
    const staffName = String((((report || {}).payload || {}).step1 || {}).staffName || '').trim();
    if (!staffName) continue;

    if (!grouped[staffName]) grouped[staffName] = [];
    if (!groupedReports[staffName]) groupedReports[staffName] = [];
    grouped[staffName].push(readableRowFromReport_(report));
    groupedReports[staffName].push(report);
  }

  const names = Object.keys(grouped).sort();
  const total = names.length;
  if (total === 0) return { total: 0, start: 0, end: 0 };
  const safeStart = Math.max(0, Math.floor(Number(startIndex || 0))) % total;
  const safeSize = batchSize == null ? total : Math.max(1, Math.floor(Number(batchSize)));
  const end = Math.min(safeStart + safeSize, total);

  for (let i = safeStart; i < end; i += 1) {
    const staffName = names[i];
    const sheetName = buildStaffSheetName_(staffName);
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.insertSheet(sheetName);
    sheet.clear();
    sheet.getRange(1, 1, 1, READABLE_HEADERS.length).setValues([READABLE_HEADERS]);
    const outRows = grouped[staffName];
    if (outRows.length > 0) {
      sheet.getRange(2, 1, outRows.length, READABLE_HEADERS.length).setValues(outRows);
    }
    sheet.setFrozenRows(1);

    buildStaffSummarySheet_(ss, staffName, groupedReports[staffName] || []);
  }
  return { total: total, start: safeStart, end: end };
}

function buildStaffSheetName_(staffName) {
  const base = `${String(staffName || '').trim()}${STAFF_SHEET_NAME_SUFFIX}`;
  const sanitized = String(base || '')
    .replace(/[\\/?*[\]:]/g, '_')
    .trim();
  return sanitized.length > 90 ? sanitized.slice(0, 90) : sanitized;
}

function buildStaffSummarySheet_(ss, staffName, reports) {
  const sheetName = buildStaffSummarySheetName_(staffName);
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  sheet.clear();

  const list = Array.isArray(reports) ? reports : [];
  const sorted = list.slice().sort((a, b) => {
    const aDate = normalizeDateOnly_(((((a || {}).payload || {}).step1 || {}).workDate) || '');
    const bDate = normalizeDateOnly_(((((b || {}).payload || {}).step1 || {}).workDate) || '');
    return aDate.localeCompare(bDate);
  });

  const totals = summarizeReports_(sorted);
  const periods = buildPeriodSummaries_(sorted);
  const items = getSummaryItems_();

  sheet.getRange(1, 1).setValue('全体集計');
  sheet.getRange(2, 1, 1, 6).setValues([['スタッフ名', 'キャッチ数', '着座数', '成約台数合計', '着座率', '成約率']]);
  sheet.getRange(3, 1, 1, 6).setValues([[
    staffName,
    totals.catchCount,
    totals.seatedCount,
    totals.contractCount,
    totals.seatedRate,
    totals.contractRate
  ]]);

  sheet.getRange(5, 1).setValue('期間集計');
  sheet.getRange(6, 1).setValue(staffName);
  sheet.getRange(7, 1, 1, 6).setValues([['期間', 'キャッチ数', '着座数', '成約台数合計', '着座率', '成約率']]);

  const periodRows = periods.map((p) => [
    p.label,
    p.catchCount,
    p.seatedCount,
    p.contractCount,
    p.seatedRate,
    p.contractRate
  ]);
  if (periodRows.length > 0) {
    sheet.getRange(8, 1, periodRows.length, 6).setValues(periodRows);
  }

  const detailTop = 10 + Math.max(periodRows.length, 1);
  sheet.getRange(detailTop, 1).setValue('期間内訳');
  const periodLabels = periods.map((p) => p.label);
  sheet.getRange(detailTop + 1, 1, 1, periodLabels.length + 1).setValues([['項目名'].concat(periodLabels)]);

  const detailRows = items.map((item) => {
    const row = [item.label];
    for (let i = 0; i < periods.length; i += 1) {
      const val = item.getter(periods[i].totals);
      row.push(toInt_(val));
    }
    return row;
  });
  if (detailRows.length > 0) {
    sheet.getRange(detailTop + 2, 1, detailRows.length, periodLabels.length + 1).setValues(detailRows);
  }

  sheet.getRange(3, 5, 1, 2).setNumberFormat('0.0%');
  if (periodRows.length > 0) {
    sheet.getRange(8, 5, periodRows.length, 2).setNumberFormat('0.0%');
  }
  sheet.setFrozenRows(2);
}

function buildStaffSummarySheetName_(staffName) {
  const base = `${String(staffName || '').trim()}${STAFF_SUMMARY_SHEET_NAME_SUFFIX}`;
  const sanitized = String(base || '')
    .replace(/[\\/?*[\]:]/g, '_')
    .trim();
  return sanitized.length > 90 ? sanitized.slice(0, 90) : sanitized;
}

function summarizeReports_(reports) {
  const totals = createEmptyTotals_();
  const list = Array.isArray(reports) ? reports : [];
  for (let i = 0; i < list.length; i += 1) {
    addReportToTotals_(totals, list[i]);
  }
  totals.seatedRate = totals.catchCount > 0 ? totals.seatedCount / totals.catchCount : 0;
  totals.contractRate = totals.seatedCount > 0 ? totals.contractCount / totals.seatedCount : 0;
  return totals;
}

function buildPeriodSummaries_(reports) {
  const map = {};
  const list = Array.isArray(reports) ? reports : [];
  for (let i = 0; i < list.length; i += 1) {
    const report = list[i] || {};
    const workDate = normalizeDateOnly_(((((report || {}).payload || {}).step1 || {}).workDate) || '');
    const dateObj = toDateFromYmd_(workDate);
    if (!dateObj) continue;
    const key = getWeekStartKey_(dateObj);
    if (!map[key]) {
      const start = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate() - ((dateObj.getDay() + 6) % 7));
      const end = new Date(start.getFullYear(), start.getMonth(), start.getDate() + 6);
      map[key] = {
        key: key,
        label: `${formatMd_(start)} ～ ${formatMd_(end)}`,
        totals: createEmptyTotals_()
      };
    }
    addReportToTotals_(map[key].totals, report);
  }

  const periods = Object.keys(map)
    .sort()
    .map((key) => {
      const p = map[key];
      p.catchCount = p.totals.catchCount;
      p.seatedCount = p.totals.seatedCount;
      p.contractCount = p.totals.contractCount;
      p.seatedRate = p.catchCount > 0 ? p.seatedCount / p.catchCount : 0;
      p.contractRate = p.seatedCount > 0 ? p.contractCount / p.seatedCount : 0;
      return p;
    });
  return periods;
}

function createEmptyTotals_() {
  return {
    catchCount: 0,
    seatedCount: 0,
    contractCount: 0,
    step3: {
      auMnpSim: 0,
      auMnpHs: 0,
      auNewSim: 0,
      auNewHs: 0,
      uqMnpSim: 0,
      uqMnpHs: 0,
      uqNewSim: 0,
      uqNewHs: 0,
      cellUp: 0
    },
    ltv: {
      auDenki: 0,
      goldCard: 0,
      silverCard: 0,
      rankUp: 0,
      jibunBank: 0,
      norton: 0,
      auHikari_new: 0,
      auHikari_fromDocomo: 0,
      auHikari_fromSoftbank: 0,
      auHikari_fromOther: 0,
      blHikari_new: 0,
      blHikari_fromDocomo: 0,
      blHikari_fromSoftbank: 0,
      blHikari_fromOther: 0,
      commufaHikari_new: 0,
      commufaHikari_fromDocomo: 0,
      commufaHikari_fromSoftbank: 0,
      commufaHikari_fromOther: 0
    }
  };
}

function addReportToTotals_(totals, report) {
  const p = (report || {}).payload || {};
  const step2 = p.step2 || {};
  const step3 = p.step3 || {};
  const newA = step3.newAcquisitions || {};
  const ltv = step3.ltv || {};
  const auH = ltv.auHikariBreakdown || {};
  const blH = ltv.blHikariBreakdown || {};
  const cmH = ltv.commufaHikariBreakdown || {};

  totals.catchCount += toInt_(step2.catchCount);
  totals.seatedCount += toInt_(step2.seated);

  totals.step3.auMnpSim += toInt_(newA.auMnpSim);
  totals.step3.auMnpHs += toInt_(newA.auMnpHs);
  totals.step3.auNewSim += toInt_(newA.auNewSim);
  totals.step3.auNewHs += toInt_(newA.auNewHs);
  totals.step3.uqMnpSim += toInt_(newA.uqMnpSim);
  totals.step3.uqMnpHs += toInt_(newA.uqMnpHs);
  totals.step3.uqNewSim += toInt_(newA.uqNewSim);
  totals.step3.uqNewHs += toInt_(newA.uqNewHs);
  totals.step3.cellUp += toInt_(newA.cellUp);

  totals.contractCount +=
    toInt_(newA.auMnpSim) +
    toInt_(newA.auMnpHs) +
    toInt_(newA.auNewSim) +
    toInt_(newA.auNewHs) +
    toInt_(newA.uqMnpSim) +
    toInt_(newA.uqMnpHs) +
    toInt_(newA.uqNewSim) +
    toInt_(newA.uqNewHs);

  totals.ltv.auDenki += toInt_(ltv.auDenki);
  totals.ltv.goldCard += toInt_(ltv.goldCard);
  totals.ltv.silverCard += toInt_(ltv.silverCard);
  totals.ltv.rankUp += toInt_(ltv.rankUp);
  totals.ltv.jibunBank += toInt_(ltv.jibunBank);
  totals.ltv.norton += toInt_(ltv.norton);
  totals.ltv.auHikari_new += toInt_(auH.new);
  totals.ltv.auHikari_fromDocomo += toInt_(auH.fromDocomo);
  totals.ltv.auHikari_fromSoftbank += toInt_(auH.fromSoftbank);
  totals.ltv.auHikari_fromOther += toInt_(auH.fromOther);
  totals.ltv.blHikari_new += toInt_(blH.new);
  totals.ltv.blHikari_fromDocomo += toInt_(blH.fromDocomo);
  totals.ltv.blHikari_fromSoftbank += toInt_(blH.fromSoftbank);
  totals.ltv.blHikari_fromOther += toInt_(blH.fromOther);
  totals.ltv.commufaHikari_new += toInt_(cmH.new);
  totals.ltv.commufaHikari_fromDocomo += toInt_(cmH.fromDocomo);
  totals.ltv.commufaHikari_fromSoftbank += toInt_(cmH.fromSoftbank);
  totals.ltv.commufaHikari_fromOther += toInt_(cmH.fromOther);
}

function getSummaryItems_() {
  return [
    { label: 'au MNP SIM単', getter: (t) => t.step3.auMnpSim },
    { label: 'au MNP HS', getter: (t) => t.step3.auMnpHs },
    { label: 'au純新規 SIM単', getter: (t) => t.step3.auNewSim },
    { label: 'au純新規 HS', getter: (t) => t.step3.auNewHs },
    { label: 'UQ MNP SIM単', getter: (t) => t.step3.uqMnpSim },
    { label: 'UQ MNP HS', getter: (t) => t.step3.uqMnpHs },
    { label: 'UQ純新規 SIM単', getter: (t) => t.step3.uqNewSim },
    { label: 'UQ純新規 HS', getter: (t) => t.step3.uqNewHs },
    { label: 'セルアップ', getter: (t) => t.step3.cellUp },
    { label: 'auでんき', getter: (t) => t.ltv.auDenki },
    { label: 'ゴールドカード', getter: (t) => t.ltv.goldCard },
    { label: 'シルバーカード', getter: (t) => t.ltv.silverCard },
    { label: 'ランクアップ', getter: (t) => t.ltv.rankUp },
    { label: 'じぶん銀行', getter: (t) => t.ltv.jibunBank },
    { label: 'ノートン', getter: (t) => t.ltv.norton },
    { label: 'auひかり 新規', getter: (t) => t.ltv.auHikari_new },
    { label: 'auひかり ドコモ光から切替', getter: (t) => t.ltv.auHikari_fromDocomo },
    { label: 'auひかり ソフトバンク光から切替', getter: (t) => t.ltv.auHikari_fromSoftbank },
    { label: 'auひかり その他から切替', getter: (t) => t.ltv.auHikari_fromOther },
    { label: 'BLひかり 新規', getter: (t) => t.ltv.blHikari_new },
    { label: 'BLひかり ドコモ光から切替', getter: (t) => t.ltv.blHikari_fromDocomo },
    { label: 'BLひかり ソフトバンク光から切替', getter: (t) => t.ltv.blHikari_fromSoftbank },
    { label: 'BLひかり その他から切替', getter: (t) => t.ltv.blHikari_fromOther },
    { label: 'コミュファ光 新規', getter: (t) => t.ltv.commufaHikari_new },
    { label: 'コミュファ光 ドコモ光から切替', getter: (t) => t.ltv.commufaHikari_fromDocomo },
    { label: 'コミュファ光 ソフトバンク光から切替', getter: (t) => t.ltv.commufaHikari_fromSoftbank },
    { label: 'コミュファ光 その他から切替', getter: (t) => t.ltv.commufaHikari_fromOther }
  ];
}

function toDateFromYmd_(ymd) {
  const text = String(ymd || '').trim();
  const matched = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!matched) return null;
  const y = Number(matched[1]);
  const m = Number(matched[2]) - 1;
  const d = Number(matched[3]);
  const date = new Date(y, m, d);
  if (!isFinite(date.getTime())) return null;
  return date;
}

function getWeekStartKey_(date) {
  const d = new Date(date.getFullYear(), date.getMonth(), date.getDate() - ((date.getDay() + 6) % 7));
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function formatMd_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM/dd');
}

function readableRowFromSource_(row, idx) {
  const report = reportFromSourceRow_(row, idx);
  return readableRowFromReport_(report);
}

function readableRowFromReport_(report) {
  const p = report.payload || {};
  const step1 = p.step1 || {};
  const step2 = p.step2 || {};
  const seated = step2.seatedBreakdown || {};
  const step3 = p.step3 || {};
  const newA = step3.newAcquisitions || {};
  const ltv = step3.ltv || {};
  const auH = ltv.auHikariBreakdown || {};
  const blH = ltv.blHikariBreakdown || {};
  const cmH = ltv.commufaHikariBreakdown || {};
  const step4Cases = Array.isArray((p.step4 || {}).cases) ? p.step4.cases : [];
  const step5Cases = Array.isArray((p.step5 || {}).cases) ? p.step5.cases : [];
  const first4 = step4Cases.find((x) => x && (x.visitReason || x.customerType || x.talkTag || x.talkDetail || x.contractFactor || x.other)) || {};
  const first5 = step5Cases.find((x) => x && (x.improvePoint || x.reason || x.other)) || {};
  const step5_5 = p.step5_5 || {};
  const step6 = p.step6 || {};
  const photos = Array.isArray(step1.photos) ? step1.photos : [];
  const photoUrls = photos.map((x) => x.url || x.dataUrl || '').filter(Boolean).join('\n');

  return [
    report.id || '',
    formatDateTimeWithWeekday_(report.createdAt || ''),
    formatDateTimeWithWeekday_(report.updatedAt || ''),
    report.confirmed ? '確認済み' : '未確認',
    report.confirmedBy || '',
    formatDateTimeWithWeekday_(report.confirmedAt || ''),
    step1.staffName || '',
    step1.jobRole || '',
    formatDateOnlyWithWeekday_(step1.workDate || ''),
    step1.workPlaceType || '',
    step1.storeName || '',
    step1.eventVenue || '',
    step1.eventOverallTarget || '',
    step1.eventVenueTarget || '',
    photos.length,
    photoUrls,
    toInt_(step2.visitors),
    toInt_(step2.catchCount),
    toInt_(step2.seated),
    toInt_(step2.prospects),
    toInt_(seated.auUqExisting),
    toInt_(seated.sbYmobile),
    toInt_(seated.docomoAhamo),
    toInt_(seated.rakuten),
    toInt_(seated.other),
    toInt_(newA.auMnpSim),
    toInt_(newA.auMnpHs),
    toInt_(newA.auNewSim),
    toInt_(newA.auNewHs),
    toInt_(newA.uqMnpSim),
    toInt_(newA.uqMnpHs),
    toInt_(newA.uqNewSim),
    toInt_(newA.uqNewHs),
    toInt_(newA.cellUp),
    toInt_(ltv.auDenki),
    toInt_(ltv.goldCard),
    toInt_(ltv.silverCard),
    toInt_(ltv.rankUp),
    toInt_(ltv.jibunBank),
    toInt_(ltv.norton),
    toInt_(auH.new),
    toInt_(auH.fromDocomo),
    toInt_(auH.fromSoftbank),
    toInt_(auH.fromOther),
    toInt_(blH.new),
    toInt_(blH.fromDocomo),
    toInt_(blH.fromSoftbank),
    toInt_(blH.fromOther),
    toInt_(cmH.new),
    toInt_(cmH.fromDocomo),
    toInt_(cmH.fromSoftbank),
    toInt_(cmH.fromOther),
    step4Cases.length,
    first4.visitReason || '',
    first4.customerType || '',
    first4.talkTag || '',
    first4.talkDetail || '',
    first4.contractFactor || '',
    first4.other || '',
    step5Cases.length,
    first5.improvePoint || '',
    first5.reason || '',
    first5.other || '',
    step5_5.venueEvaluation || '',
    step5_5.other || '',
    step6.impression || '',
    step6.notes || '',
    step6.adminSummary || '',
    JSON.stringify(report)
  ];
}

function reportFromSourceRow_(row, idx) {
  const rawJson = pickFromRow_(row, idx, ['rawJson', '生データJSON', '管理用:rawJson']);
  if (rawJson) {
    try {
      const parsed = JSON.parse(String(rawJson));
      if (parsed && typeof parsed === 'object') return parsed;
    } catch (err) {
      // fallback below
    }
  }

  const step1 = {
    staffName: String(pickFromRow_(row, idx, ['staffName', 'payload.step1.staffName']) || ''),
    jobRole: String(pickFromRow_(row, idx, ['jobRole', 'payload.step1.jobRole']) || ''),
    workDate: normalizeDateOnly_(pickFromRow_(row, idx, ['workDate', 'payload.step1.workDate']) || ''),
    workPlaceType: String(pickFromRow_(row, idx, ['workPlaceType', 'payload.step1.workPlaceType']) || ''),
    storeName: String(pickFromRow_(row, idx, ['storeName', 'payload.step1.storeName']) || ''),
    eventVenue: String(pickFromRow_(row, idx, ['eventVenue', 'payload.step1.eventVenue']) || ''),
    eventOverallTarget: String(pickFromRow_(row, idx, ['eventOverallTarget', 'payload.step1.eventOverallTarget']) || ''),
    eventVenueTarget: String(pickFromRow_(row, idx, ['eventVenueTarget', 'payload.step1.eventVenueTarget']) || ''),
    photos: parsePhotoUrls_(pickFromRow_(row, idx, ['photoUrls']))
  };

  return {
    id: String(pickFromRow_(row, idx, ['reportId']) || ''),
    createdAt: String(pickFromRow_(row, idx, ['createdAt']) || ''),
    updatedAt: String(pickFromRow_(row, idx, ['updatedAt']) || ''),
    confirmed: String(pickFromRow_(row, idx, ['confirmed']) || '') === '確認済み',
    confirmedBy: String(pickFromRow_(row, idx, ['confirmedBy']) || ''),
    confirmedAt: String(pickFromRow_(row, idx, ['confirmedAt']) || ''),
    payload: {
      step1: step1,
      step2: {
        visitors: toInt_(pickFromRow_(row, idx, ['step2_visitors', 'payload.step2.visitors'])),
        catchCount: toInt_(pickFromRow_(row, idx, ['step2_catchCount', 'payload.step2.catchCount'])),
        seated: toInt_(pickFromRow_(row, idx, ['step2_seated', 'payload.step2.seated'])),
        prospects: toInt_(pickFromRow_(row, idx, ['step2_prospects', 'payload.step2.prospects'])),
        seatedBreakdown: {
          auUqExisting: toInt_(pickFromRow_(row, idx, ['step2_seated_auUqExisting', 'payload.step2.seatedBreakdown.auUqExisting'])),
          sbYmobile: toInt_(pickFromRow_(row, idx, ['step2_seated_sbYmobile', 'payload.step2.seatedBreakdown.sbYmobile'])),
          docomoAhamo: toInt_(pickFromRow_(row, idx, ['step2_seated_docomoAhamo', 'payload.step2.seatedBreakdown.docomoAhamo'])),
          rakuten: toInt_(pickFromRow_(row, idx, ['step2_seated_rakuten', 'payload.step2.seatedBreakdown.rakuten'])),
          other: toInt_(pickFromRow_(row, idx, ['step2_seated_other', 'payload.step2.seatedBreakdown.other']))
        }
      },
      step3: {
        newAcquisitions: {
          auMnpSim: toInt_(pickFromRow_(row, idx, ['step3_new_auMnpSim', 'payload.step3.newAcquisitions.auMnpSim'])),
          auMnpHs: toInt_(pickFromRow_(row, idx, ['step3_new_auMnpHs', 'payload.step3.newAcquisitions.auMnpHs'])),
          auNewSim: toInt_(pickFromRow_(row, idx, ['step3_new_auNewSim', 'payload.step3.newAcquisitions.auNewSim'])),
          auNewHs: toInt_(pickFromRow_(row, idx, ['step3_new_auNewHs', 'payload.step3.newAcquisitions.auNewHs'])),
          uqMnpSim: toInt_(pickFromRow_(row, idx, ['step3_new_uqMnpSim', 'payload.step3.newAcquisitions.uqMnpSim'])),
          uqMnpHs: toInt_(pickFromRow_(row, idx, ['step3_new_uqMnpHs', 'payload.step3.newAcquisitions.uqMnpHs'])),
          uqNewSim: toInt_(pickFromRow_(row, idx, ['step3_new_uqNewSim', 'payload.step3.newAcquisitions.uqNewSim'])),
          uqNewHs: toInt_(pickFromRow_(row, idx, ['step3_new_uqNewHs', 'payload.step3.newAcquisitions.uqNewHs'])),
          cellUp: toInt_(pickFromRow_(row, idx, ['step3_new_cellUp', 'payload.step3.newAcquisitions.cellUp']))
        },
        ltv: {
          auDenki: toInt_(pickFromRow_(row, idx, ['step3_ltv_auDenki', 'payload.step3.ltv.auDenki'])),
          goldCard: toInt_(pickFromRow_(row, idx, ['step3_ltv_goldCard', 'payload.step3.ltv.goldCard'])),
          silverCard: toInt_(pickFromRow_(row, idx, ['step3_ltv_silverCard', 'payload.step3.ltv.silverCard'])),
          rankUp: toInt_(pickFromRow_(row, idx, ['step3_ltv_rankUp', 'payload.step3.ltv.rankUp'])),
          jibunBank: toInt_(pickFromRow_(row, idx, ['step3_ltv_jibunBank', 'payload.step3.ltv.jibunBank'])),
          norton: toInt_(pickFromRow_(row, idx, ['step3_ltv_norton', 'payload.step3.ltv.norton'])),
          auHikariBreakdown: {
            new: toInt_(pickFromRow_(row, idx, ['step3_ltv_auHikari_new', 'payload.step3.ltv.auHikariBreakdown.new'])),
            fromDocomo: toInt_(pickFromRow_(row, idx, ['step3_ltv_auHikari_fromDocomo', 'payload.step3.ltv.auHikariBreakdown.fromDocomo'])),
            fromSoftbank: toInt_(pickFromRow_(row, idx, ['step3_ltv_auHikari_fromSoftbank', 'payload.step3.ltv.auHikariBreakdown.fromSoftbank'])),
            fromOther: toInt_(pickFromRow_(row, idx, ['step3_ltv_auHikari_fromOther', 'payload.step3.ltv.auHikariBreakdown.fromOther']))
          },
          blHikariBreakdown: {
            new: toInt_(pickFromRow_(row, idx, ['step3_ltv_blHikari_new', 'payload.step3.ltv.blHikariBreakdown.new'])),
            fromDocomo: toInt_(pickFromRow_(row, idx, ['step3_ltv_blHikari_fromDocomo', 'payload.step3.ltv.blHikariBreakdown.fromDocomo'])),
            fromSoftbank: toInt_(pickFromRow_(row, idx, ['step3_ltv_blHikari_fromSoftbank', 'payload.step3.ltv.blHikariBreakdown.fromSoftbank'])),
            fromOther: toInt_(pickFromRow_(row, idx, ['step3_ltv_blHikari_fromOther', 'payload.step3.ltv.blHikariBreakdown.fromOther']))
          },
          commufaHikariBreakdown: {
            new: toInt_(pickFromRow_(row, idx, ['step3_ltv_commufaHikari_new', 'payload.step3.ltv.commufaHikariBreakdown.new'])),
            fromDocomo: toInt_(pickFromRow_(row, idx, ['step3_ltv_commufaHikari_fromDocomo', 'payload.step3.ltv.commufaHikariBreakdown.fromDocomo'])),
            fromSoftbank: toInt_(pickFromRow_(row, idx, ['step3_ltv_commufaHikari_fromSoftbank', 'payload.step3.ltv.commufaHikariBreakdown.fromSoftbank'])),
            fromOther: toInt_(pickFromRow_(row, idx, ['step3_ltv_commufaHikari_fromOther', 'payload.step3.ltv.commufaHikariBreakdown.fromOther']))
          }
        }
      },
      step4: {
        cases: parseCaseJson_(pickFromRow_(row, idx, ['step4_casesJson', 'payload.step4.cases']))
      },
      step5: {
        cases: parseCaseJson_(pickFromRow_(row, idx, ['step5_casesJson', 'payload.step5.cases']))
      },
      step5_5: {
        venueEvaluation: String(pickFromRow_(row, idx, ['step5_5_venueEvaluation', 'payload.step5_5.venueEvaluation']) || ''),
        other: String(pickFromRow_(row, idx, ['step5_5_other', 'payload.step5_5.other']) || '')
      },
      step6: {
        impression: String(pickFromRow_(row, idx, ['impression', 'payload.step6.impression']) || ''),
        notes: String(pickFromRow_(row, idx, ['notes', 'payload.step6.notes']) || ''),
        adminSummary: String(pickFromRow_(row, idx, ['adminSummary', 'payload.step6.adminSummary']) || '')
      }
    }
  };
}

function pickFromRow_(row, idx, keys) {
  for (let i = 0; i < keys.length; i += 1) {
    const p = idx[keys[i]];
    if (typeof p === 'number') {
      const v = row[p];
      if (v !== '' && v != null) return v;
    }
  }
  return '';
}

function parseCaseJson_(value) {
  if (!value) return [];
  if (Array.isArray(value)) return value;
  const text = String(value || '').trim();
  if (!text) return [];
  try {
    const parsed = JSON.parse(text);
    return Array.isArray(parsed) ? parsed : [];
  } catch (err) {
    return [];
  }
}

function parsePhotoUrls_(value) {
  const text = String(value || '').trim();
  if (!text) return [];
  return text
    .split(/\n+/)
    .map((url) => String(url || '').trim())
    .filter(Boolean)
    .map((url) => ({ name: '会場写真', type: 'image/jpeg', size: 0, url: url, dataUrl: '', fileId: '' }));
}

function normalizeDateOnly_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const text = String(value || '').trim();
  const matched = text.match(/^(\d{4}-\d{2}-\d{2})/);
  return matched ? matched[1] : text;
}

function formatDateOnlyWithWeekday_(value) {
  const normalized = normalizeDateOnly_(value);
  if (!normalized) return '';
  const matched = normalized.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!matched) return normalized;
  const y = Number(matched[1]);
  const m = Number(matched[2]) - 1;
  const d = Number(matched[3]);
  const date = new Date(y, m, d);
  if (!isFinite(date.getTime())) return normalized;
  const days = ['日', '月', '火', '水', '木', '金', '土'];
  return `${normalized}(${days[date.getDay()]})`;
}

function formatDateTimeWithWeekday_(value) {
  const text = String(value || '').trim();
  if (!text) return '';
  const normalized = normalizeDateOnly_(text);
  if (!normalized) return text;
  const withWeekday = formatDateOnlyWithWeekday_(normalized);
  if (!withWeekday || withWeekday === normalized) return text;
  return text.replace(normalized, withWeekday);
}

function toInt_(value) {
  const n = Number(value);
  if (!isFinite(n) || n < 0) return 0;
  return Math.floor(n);
}
