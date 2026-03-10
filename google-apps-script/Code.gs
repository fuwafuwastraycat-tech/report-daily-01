const SHEET_NAME = '日報データ';
const HEADER = [
  '日報ID',
  '作成日時',
  '更新日時',
  '確認状況',
  '確認者',
  '確認日時',
  'スタッフ名',
  '区分',
  '稼働日',
  '店舗名',
  'イベント会場',
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
  '成約事例件数',
  '成約事例1 来店理由',
  '成約事例1 客層',
  '成約事例1 決め手トーク（タグ）',
  '成約事例1 決め手トーク（具体）',
  '成約事例1 成約要因',
  '成約事例1 その他',
  '成約事例JSON',
  '改善事例件数',
  '改善事例1 改善ポイント',
  '改善事例1 理由（具体）',
  '改善事例1 その他',
  '改善事例JSON',
  'イベント会場の評価',
  'その他（会場評価）',
  '所感（短文）',
  'その他備考',
  '管理者専用 総括コメント',
  '生データJSON'
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
  ensureHeader_(sheet);
  return sheet;
}

function rowFromReport_(report) {
  const p = report.payload || {};
  const step1 = p.step1 || {};
  const step2 = p.step2 || {};
  const seated = step2.seatedBreakdown || {};
  const step3 = p.step3 || {};
  const newA = step3.newAcquisitions || {};
  const ltv = step3.ltv || {};
  const auHikari = ltv.auHikariBreakdown || {};
  const blHikari = ltv.blHikariBreakdown || {};
  const commufaHikari = ltv.commufaHikariBreakdown || {};
  const step4 = p.step4 || {};
  const step5 = p.step5 || {};
  const step5_5 = p.step5_5 || {};
  const step6 = p.step6 || {};
  const photos = getPhotoList_(step1);
  const photoUrls = photos
    .map((item) => item.url || item.dataUrl || '')
    .filter(Boolean)
    .join('\n');
  const step4Cases = getStep4Cases_(step4);
  const step5Cases = getStep5Cases_(step5);
  const firstStep4Case = pickFirstFilledStep4Case_(step4);
  const firstStep5Case = pickFirstFilledStep5Case_(step5);

  return [
    report.id || '',
    report.createdAt || '',
    report.updatedAt || '',
    report.confirmed ? '確認済み' : '未確認',
    report.confirmedBy || '',
    report.confirmedAt || '',
    step1.staffName || '',
    step1.workPlaceType || '',
    step1.workDate || '',
    step1.storeName || '',
    step1.eventVenue || '',
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
    toInt_(auHikari.new),
    toInt_(auHikari.fromDocomo),
    toInt_(auHikari.fromSoftbank),
    toInt_(auHikari.fromOther),
    toInt_(blHikari.new),
    toInt_(blHikari.fromDocomo),
    toInt_(blHikari.fromSoftbank),
    toInt_(blHikari.fromOther),
    toInt_(commufaHikari.new),
    toInt_(commufaHikari.fromDocomo),
    toInt_(commufaHikari.fromSoftbank),
    toInt_(commufaHikari.fromOther),
    step4Cases.length,
    firstStep4Case ? (firstStep4Case.visitReason || '') : '',
    firstStep4Case ? (firstStep4Case.customerType || '') : '',
    firstStep4Case ? (firstStep4Case.talkTag || '') : '',
    firstStep4Case ? (firstStep4Case.talkDetail || '') : '',
    firstStep4Case ? (firstStep4Case.contractFactor || '') : '',
    firstStep4Case ? (firstStep4Case.other || '') : '',
    JSON.stringify(step4Cases),
    step5Cases.length,
    firstStep5Case ? (firstStep5Case.improvePoint || '') : '',
    firstStep5Case ? (firstStep5Case.reason || '') : '',
    firstStep5Case ? (firstStep5Case.other || '') : '',
    JSON.stringify(step5Cases),
    step5_5.venueEvaluation || '',
    step5_5.other || '',
    step6.impression || '',
    step6.notes || '',
    step6.adminSummary || '',
    JSON.stringify(report)
  ];
}

function ensureHeader_(sheet) {
  if (sheet.getMaxColumns() < HEADER.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), HEADER.length - sheet.getMaxColumns());
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]);
    return;
  }
  const current = sheet.getRange(1, 1, 1, HEADER.length).getValues()[0];
  if (!isSameHeader_(current, HEADER)) {
    sheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]);
  }
}

function isSameHeader_(a, b) {
  if (!a || !b || a.length !== b.length) return false;
  for (let i = 0; i < a.length; i += 1) {
    if (String(a[i] || '') !== String(b[i] || '')) return false;
  }
  return true;
}

function toInt_(value) {
  const n = Number(value);
  if (!isFinite(n) || n < 0) return 0;
  return Math.floor(n);
}

function getPhotoList_(step1) {
  const photos = Array.isArray(step1.photos) ? step1.photos : [];
  if (photos.length > 0) return photos;
  const legacy = step1.photoUrl || step1.photoDataUrl || '';
  if (!legacy) return [];
  return [{ url: step1.photoUrl || '', dataUrl: step1.photoDataUrl || '' }];
}

function getStep4Cases_(step4) {
  return Array.isArray(step4.cases) ? step4.cases : [];
}

function getStep5Cases_(step5) {
  return Array.isArray(step5.cases) ? step5.cases : [];
}

function pickFirstFilledStep4Case_(step4) {
  const list = Array.isArray(step4.cases) ? step4.cases : [];
  for (let i = 0; i < list.length; i += 1) {
    const item = list[i] || {};
    if (item.visitReason || item.customerType || item.talkTag || item.talkDetail || item.contractFactor || item.other) {
      return item;
    }
  }
  return null;
}

function pickFirstFilledStep5Case_(step5) {
  const list = Array.isArray(step5.cases) ? step5.cases : [];
  for (let i = 0; i < list.length; i += 1) {
    const item = list[i] || {};
    if (item.improvePoint || item.reason || item.other) {
      return item;
    }
  }
  return null;
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

  const rawJsonCol = HEADER.length;
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
    try {
      return DriveApp.getFolderById(folderId);
    } catch (err) {
      // 指定フォルダにアクセスできない場合はルートへフォールバック
      return DriveApp.getRootFolder();
    }
  }
  return DriveApp.getRootFolder();
}
