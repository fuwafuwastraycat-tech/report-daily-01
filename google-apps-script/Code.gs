const SHEET_NAME = '日報データ';
const CORE_HEADERS = [
  'reportId',
  'createdAt',
  'updatedAt',
  'confirmed',
  'confirmedBy',
  'confirmedAt',
  'staffName',
  'workPlaceType',
  'workDate',
  'storeName',
  'eventVenue',
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
    workPlaceType: step1.workPlaceType || '',
    workDate: step1.workDate || '',
    storeName: step1.storeName || '',
    eventVenue: step1.eventVenue || '',
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
    createdAt: String(pickCell_(row, idx, ['createdAt']) || base.createdAt || ''),
    updatedAt: String(pickCell_(row, idx, ['updatedAt']) || base.updatedAt || ''),
    confirmed: confirmed,
    confirmedBy: String(pickCell_(row, idx, ['confirmedBy']) || base.confirmedBy || ''),
    confirmedAt: String(pickCell_(row, idx, ['confirmedAt']) || base.confirmedAt || ''),
    payload: {
      ...payload,
      step1: {
        ...step1,
        staffName: String(pickCell_(row, idx, ['staffName']) || step1.staffName || ''),
        workPlaceType: String(pickCell_(row, idx, ['workPlaceType']) || step1.workPlaceType || ''),
        workDate: String(pickCell_(row, idx, ['workDate']) || step1.workDate || ''),
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
