const SHEET_NAME = '日報データ';
const HEADER = [
  'reportId',
  'updatedAt',
  'staffName',
  'workDate',
  'storeName',
  'eventVenue',
  'confirmed',
  'folder',
  'successTitle',
  'failureTitle',
  'impression',
  'adminSummary',
  'rawJson'
];

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
      const reports = listReports();
      return jsonOut({ ok: true, reports: reports });
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
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]);
  }
  return sheet;
}

function rowFromReport_(report) {
  const p = report.payload || {};
  const step1 = p.step1 || {};
  const step4 = p.step4 || {};
  const step5 = p.step5 || {};
  const step6 = p.step6 || {};

  return [
    report.id || '',
    report.updatedAt || '',
    step1.staffName || '',
    step1.workDate || '',
    step1.storeName || '',
    step1.eventVenue || '',
    report.confirmed ? '確認済み' : '未確認',
    report.folder || '未分類',
    step4.title || '',
    step5.title || '',
    step6.impression || '',
    step6.adminSummary || '',
    JSON.stringify(report)
  ];
}

function findRowByReportId_(sheet, reportId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i += 1) {
    if (values[i][0] === reportId) {
      return i + 2;
    }
  }
  return -1;
}

function upsertReport(report) {
  const sheet = getSheet_();
  const rowData = rowFromReport_(report);
  const foundRow = findRowByReportId_(sheet, report.id);
  if (foundRow > 0) {
    sheet.getRange(foundRow, 1, 1, HEADER.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
}

function deleteReport(reportId) {
  if (!reportId) return;
  const sheet = getSheet_();
  const foundRow = findRowByReportId_(sheet, reportId);
  if (foundRow > 0) {
    sheet.deleteRow(foundRow);
  }
}

function replaceAllReports(reports) {
  const sheet = getSheet_();
  const rows = reports.map((report) => rowFromReport_(report));
  sheet.clear();
  sheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, HEADER.length).setValues(rows);
  }
}

function listReports() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const rawJsonCol = HEADER.indexOf('rawJson') + 1;
  if (rawJsonCol <= 0) return [];
  const values = sheet.getRange(2, rawJsonCol, lastRow - 1, 1).getValues();

  const reports = [];
  for (let i = 0; i < values.length; i += 1) {
    const raw = values[i][0];
    if (!raw) continue;
    try {
      reports.push(JSON.parse(raw));
    } catch (err) {
      // skip broken rows
    }
  }
  return reports;
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
  // 組織ポリシーで失敗する場合があるため、共有設定エラーは握りつぶして継続
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
  if (!matched) {
    throw new Error('Invalid dataUrl');
  }
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
    return DriveApp.getFolderById(folderId);
  }
  return DriveApp.getRootFolder();
}
