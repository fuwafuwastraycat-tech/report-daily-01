const STORAGE_KEY = 'daily-report-app-v1';
const ADMIN_SESSION_KEY = 'daily-report-admin-session-v1';
const SYNC_CONFIG_KEY = 'daily-report-sync-config-v1';
const ACHIEVEMENTS_COMMENT_HIDDEN_KEY = 'daily-report-achievements-hidden-comments-v1';
const SYNC_POLL_INTERVAL_MS = 5000;
const PHOTO_MAX_EDGE_PX = 1280;
const PHOTO_JPEG_QUALITY = 0.72;
const PHOTO_MAX_DATAURL_CHARS = 500000;
const PHOTO_MAX_COUNT = 5;
const LIST_PAGE_SIZE = 50;
const ACHIEVEMENT_PERIOD_PRIMARY_COUNT = 4;
const ACHIEVEMENTS_ALL_OPTION = '__ALL__';
const STAFF_NAME_OPTIONS = ['', '今村優志', '天野竜毅', '桑原佑太', '柘植稜平', '中西琥太朗'];
const JOB_ROLE_OPTIONS = ['', '統括ディレクター', 'ディレクター', '特販', 'クローザー', 'キャッチャー'];

// 全スタッフ端末で共通利用する既定の連携先。
// ここを設定しておくと、管理者以外でも自動同期されます。
const DEFAULT_SYNC_CONFIG = {
  endpoint: 'https://script.google.com/macros/s/AKfycbyQt0jv8wwTX2gJB3DTtPeAC-hLfksq7-TVjthaR2be1BGdJ5A-rAHy_5_y-59W0Dbw/exec',
  token: 'daily-report-token'
};

const ADMIN_USERS = [
  { id: 'admin01', name: '五十嵐', password: 'admin1234' },
  { id: 'sv01', name: '米澤', password: 'sv1234' },
  { id: 'admin02', name: '小澤', password: 'admin5678' }
];

const state = {
  reports: [],
  mode: 'staff-list',
  editingId: null,
  returnView: 'staff-list',
  detailReturnView: 'staff-list',
  detailReportId: '',
  adminFocusReportId: '',
  adminReportReturnView: 'admin',
  achievementsSelectedStaff: '',
  achievementsSelectedPeriodKey: '',
  achievementsTab: 'summary',
  achievementsRankingPeriodKey: '',
  achievementsReportDrafts: {},
  hiddenAchievementCommentIds: {},
  adminConfirmedSelectedStaff: '',
  adminConfirmedSelectedReportId: '',
  form: createEmptyForm(),
  currentStep: 1,
  selectedStaffName: '',
  errors: {},
  photoPreview: null,
  photoUploading: false,
  toastTimer: null,
  adminUser: null,
  syncTimer: null,
  syncConfig: {
    endpoint: '',
    token: ''
  },
  viewLimits: createDefaultViewLimits()
};

const views = {
  staffList: document.getElementById('view-list'),
  admin: document.getElementById('view-admin'),
  achievements: document.getElementById('view-achievements'),
  adminReport: document.getElementById('view-admin-report'),
  detail: document.getElementById('view-detail'),
  form: document.getElementById('view-form')
};

const elements = {
  reportList: document.getElementById('report-list'),
  listEmpty: document.getElementById('list-empty'),
  cardTemplate: document.getElementById('card-template'),
  adminCardTemplate: document.getElementById('admin-card-template'),
  adminUncheckedList: document.getElementById('admin-unchecked-list'),
  adminUncheckedEmpty: document.getElementById('admin-unchecked-empty'),
  adminConfirmedStaffList: document.getElementById('admin-confirmed-staff-list'),
  adminConfirmedStaffEmpty: document.getElementById('admin-confirmed-staff-empty'),
  adminConfirmedDateWrap: document.getElementById('admin-confirmed-date-wrap'),
  adminConfirmedStaffTitle: document.getElementById('admin-confirmed-staff-title'),
  adminConfirmedBackButton: document.getElementById('admin-confirmed-back-button'),
  adminConfirmedDateList: document.getElementById('admin-confirmed-date-list'),
  adminConfirmedDetailWrap: document.getElementById('admin-confirmed-detail-wrap'),
  adminConfirmedDateTitle: document.getElementById('admin-confirmed-date-title'),
  adminConfirmedDetailContent: document.getElementById('admin-confirmed-detail-content'),
  adminConfirmedSummaryToggleButton: document.getElementById('admin-confirmed-summary-toggle-button'),
  adminConfirmedSummaryEditor: document.getElementById('admin-confirmed-summary-editor'),
  adminConfirmedSummaryInput: document.getElementById('admin-confirmed-summary-input'),
  adminConfirmedSummarySaveButton: document.getElementById('admin-confirmed-summary-save-button'),
  adminConfirmedDetailEditButton: document.getElementById('admin-confirmed-detail-edit-button'),
  createButton: document.getElementById('create-report-button'),
  backToListButton: document.getElementById('back-to-list-button'),
  formTitle: document.getElementById('form-title'),
  stepText: document.getElementById('step-text'),
  stepProgress: document.getElementById('step-progress'),
  stepProgressBar: document.getElementById('step-progress-bar'),
  stepContainer: document.getElementById('step-container'),
  prevStepButton: document.getElementById('prev-step-button'),
  nextStepButton: document.getElementById('next-step-button'),
  detailContainer: document.getElementById('detail-container'),
  detailBackButton: document.getElementById('detail-back-button'),
  toast: document.getElementById('toast'),
  switchStaffButton: document.getElementById('switch-staff-button'),
  switchAdminButton: document.getElementById('switch-admin-button'),
  switchAchievementsButton: document.getElementById('switch-achievements-button'),
  staffListTitle: document.getElementById('staff-list-title'),
  staffBackButton: document.getElementById('staff-back-button'),
  achievementsBackButton: document.getElementById('achievements-back-button'),
  achievementsExportPdfButton: document.getElementById('achievements-export-pdf-button'),
  achievementsTabSummaryButton: document.getElementById('achievements-tab-summary'),
  achievementsTabRankingButton: document.getElementById('achievements-tab-ranking'),
  achievementsTabReportButton: document.getElementById('achievements-tab-report'),
  achievementsStaffSelect: document.getElementById('achievements-staff-select'),
  achievementsContainer: document.getElementById('achievements-container'),
  adminAuthPanel: document.getElementById('admin-auth-panel'),
  adminContent: document.getElementById('admin-content'),
  adminLoginId: document.getElementById('admin-login-id'),
  adminLoginPassword: document.getElementById('admin-login-password'),
  adminLoginButton: document.getElementById('admin-login-button'),
  adminLoginError: document.getElementById('admin-login-error'),
  adminUserLabel: document.getElementById('admin-user-label'),
  adminLogoutButton: document.getElementById('admin-logout-button'),
  adminReportBackButton: document.getElementById('admin-report-back-button'),
  adminReportStaff: document.getElementById('admin-report-staff'),
  adminReportDate: document.getElementById('admin-report-date'),
  adminReportConfirmButton: document.getElementById('admin-report-confirm-button'),
  adminReportPreviewButton: document.getElementById('admin-report-preview-button'),
  adminReportSummaryToggleButton: document.getElementById('admin-report-summary-toggle-button'),
  adminReportSummaryEditor: document.getElementById('admin-report-summary-editor'),
  adminReportSummaryInput: document.getElementById('admin-report-summary-input'),
  adminReportSummarySaveButton: document.getElementById('admin-report-summary-save-button'),
  adminReportEditButton: document.getElementById('admin-report-edit-button'),
  syncPanel: document.getElementById('sync-panel'),
  syncEndpointInput: document.getElementById('sync-endpoint-input'),
  syncTokenInput: document.getElementById('sync-token-input'),
  syncSaveButton: document.getElementById('sync-save-button'),
  syncAllButton: document.getElementById('sync-all-button'),
  syncPullButton: document.getElementById('sync-pull-button'),
  syncStatusText: document.getElementById('sync-status-text')
};

const options = {
  jobRoles: JOB_ROLE_OPTIONS,
  workPlaceTypes: ['', '店頭SV', 'イベント'],
  successVisitReasons: ['', '料金見直し', 'MNP検討', '新規契約', '機種変更', '故障相談', 'キャッチ獲得', 'POPアイキャッチ', 'アトラクション参加', 'その他'],
  successCustomerTypes: ['', 'ご家族', '単身者', 'ご高齢者', 'その他'],
  successTalkTags: ['', '料金訴求', '特典訴求', '安心感', '限定感', '端末訴求', 'その他'],
  improvePoints: ['', '在庫不足', 'ヒアリング不足', 'キャッチ', 'クロージング', '知識不足', '特典', '時間がない', 'その他']
};

function createEmptySuccessCase() {
  return {
    visitReason: '',
    customerType: '',
    talkTag: '',
    talkDetail: '',
    contractFactor: '',
    other: ''
  };
}

function createEmptyImproveCase() {
  return {
    improvePoint: '',
    reason: '',
    other: ''
  };
}

function createDefaultViewLimits() {
  return {
    staffGroups: LIST_PAGE_SIZE,
    staffReports: LIST_PAGE_SIZE,
    adminUnchecked: LIST_PAGE_SIZE,
    adminConfirmedStaff: LIST_PAGE_SIZE,
    adminConfirmedDates: LIST_PAGE_SIZE
  };
}

init();

function init() {
  state.reports = loadReports();
  state.adminUser = loadAdminSession();
  state.hiddenAchievementCommentIds = loadHiddenAchievementCommentIds();
  state.syncConfig = loadSyncConfig();
  bindEvents();
  openStaffListView();
  renderAdminView();
  startSyncPolling();
  void pullReportsFromSheet(false);
}

function bindEvents() {
  elements.createButton.addEventListener('click', openCreateView);
  elements.backToListButton.addEventListener('click', backToReturnView);
  elements.prevStepButton.addEventListener('click', goPrevStep);
  elements.nextStepButton.addEventListener('click', goNextStepOrSubmit);
  elements.reportList.addEventListener('click', onStaffListClick);
  elements.adminUncheckedList.addEventListener('click', onAdminListClick);
  elements.adminConfirmedStaffList.addEventListener('click', onAdminConfirmedClick);
  elements.adminConfirmedDateList.addEventListener('click', onAdminConfirmedClick);

  elements.switchStaffButton.addEventListener('click', openStaffListView);
  elements.switchAdminButton.addEventListener('click', openAdminView);
  elements.switchAchievementsButton.addEventListener('click', openAchievementsView);
  elements.staffBackButton.addEventListener('click', backToStaffGroupList);
  elements.achievementsBackButton.addEventListener('click', openAdminView);
  elements.achievementsExportPdfButton.addEventListener('click', handleExportAchievementsPdf);
  elements.achievementsTabSummaryButton.addEventListener('click', () => {
    state.achievementsTab = 'summary';
    renderAchievementsView();
  });
  elements.achievementsTabRankingButton.addEventListener('click', () => {
    state.achievementsTab = 'ranking';
    renderAchievementsView();
  });
  elements.achievementsTabReportButton.addEventListener('click', () => {
    state.achievementsTab = 'report';
    renderAchievementsView();
  });
  elements.achievementsStaffSelect.addEventListener('change', onAchievementsStaffChange);
  elements.achievementsContainer.addEventListener('click', onAchievementsContainerClick);
  elements.achievementsContainer.addEventListener('input', onAchievementsContainerInput);
  elements.detailBackButton.addEventListener('click', backFromDetailView);

  elements.adminLoginButton.addEventListener('click', handleAdminLogin);
  elements.adminLogoutButton.addEventListener('click', handleAdminLogout);
  elements.adminReportBackButton.addEventListener('click', handleAdminReportBack);
  elements.adminReportConfirmButton.addEventListener('click', handleAdminReportConfirm);
  elements.adminReportPreviewButton.addEventListener('click', handleAdminReportPreview);
  elements.adminReportSummaryToggleButton.addEventListener('click', toggleAdminReportSummaryEditor);
  elements.adminReportSummarySaveButton.addEventListener('click', handleAdminReportSummarySave);
  elements.adminReportEditButton.addEventListener('click', handleAdminReportEdit);
  elements.adminConfirmedBackButton.addEventListener('click', handleAdminConfirmedBack);
  elements.adminConfirmedSummaryToggleButton.addEventListener('click', toggleAdminConfirmedSummaryEditor);
  elements.adminConfirmedSummarySaveButton.addEventListener('click', handleAdminConfirmedSummarySave);
  elements.adminConfirmedDetailEditButton.addEventListener('click', handleAdminConfirmedDetailEdit);
  elements.syncSaveButton.addEventListener('click', handleSaveSyncConfig);
  elements.syncAllButton.addEventListener('click', handleSyncAllReports);
  elements.syncPullButton.addEventListener('click', handleSyncPullReports);
  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'visible') {
      void pullReportsFromSheet(false);
    }
  });
}

function createEmptyForm() {
  return {
    step1: {
      workDate: '',
      staffName: '',
      jobRole: '',
      workPlaceType: '',
      eventCompany: '',
      storeName: '',
      eventVenue: '',
      eventOverallTarget: '',
      eventVenueTarget: '',
      photoMeta: null,
      photoDataUrl: '',
      photoUrl: '',
      photoFileId: '',
      photos: []
    },
    step2: {
      visitors: 0,
      catchCount: 0,
      seated: 0,
      prospects: 0,
      seatedBreakdown: {
        auUqExisting: 0,
        sbYmobile: 0,
        docomoAhamo: 0,
        rakuten: 0,
        other: 0
      }
    },
    step3: {
      newAcquisitions: {
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
        auHikariBreakdown: {
          new: 0,
          fromDocomo: 0,
          fromSoftbank: 0,
          fromOther: 0
        },
        blHikariBreakdown: {
          new: 0,
          fromDocomo: 0,
          fromSoftbank: 0,
          fromOther: 0
        },
        commufaHikariBreakdown: {
          new: 0,
          fromDocomo: 0,
          fromSoftbank: 0,
          fromOther: 0
        },
        auDenki: 0,
        goldCard: 0,
        silverCard: 0,
        rankUp: 0,
        jibunBank: 0,
        norton: 0
      }
    },
    step4: {
      cases: [createEmptySuccessCase()]
    },
    step5: {
      cases: [createEmptyImproveCase()]
    },
    step5_5: {
      venueEvaluation: '',
      other: ''
    },
    step6: {
      impression: '',
      notes: '',
      adminSummary: ''
    }
  };
}

function makeId() {
  if (typeof crypto !== 'undefined' && crypto.randomUUID) return crypto.randomUUID();
  return `rpt-${Date.now()}-${Math.floor(Math.random() * 10000)}`;
}

function deepCopy(obj) {
  return JSON.parse(JSON.stringify(obj));
}

function loadReports() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    return parsed.map(normalizeReport).filter(Boolean);
  } catch {
    return [];
  }
}

function normalizeReport(report) {
  if (!report || typeof report !== 'object') return null;
  const payload = mergeForm(createEmptyForm(), report.payload || {});
  normalizeCaseSections(payload);
  payload.step1.staffName = normalizeStaffName(payload.step1.staffName);
  return {
    id: report.id || makeId(),
    createdAt: report.createdAt || new Date().toISOString(),
    updatedAt: report.updatedAt || new Date().toISOString(),
    folder: typeof report.folder === 'string' && report.folder.trim() ? report.folder.trim() : '未分類',
    confirmed: Boolean(report.confirmed),
    confirmedBy: normalizeConfirmedByName(report.confirmedBy),
    confirmedAt: typeof report.confirmedAt === 'string' ? report.confirmedAt : '',
    payload
  };
}

function normalizeConfirmedByName(value) {
  if (typeof value !== 'string') return '';
  const normalized = value.trim();
  if (!normalized) return '';
  const legacyMap = {
    管理者A: '五十嵐',
    '管理者A（admin01）': '五十嵐',
    SV管理者: '米澤',
    'SV管理者（sv01）': '米澤',
    管理者B: '小澤',
    '管理者B（admin02）': '小澤'
  };
  return legacyMap[normalized] || normalized;
}

function normalizeStaffName(value) {
  const normalized = String(value || '').trim();
  if (!normalized) return '';
  if (STAFF_NAME_OPTIONS.includes(normalized)) return normalized;
  const legacyMap = {
    桑原: '桑原佑太',
    今桑原: '桑原佑太'
  };
  return legacyMap[normalized] || normalized;
}

function normalizeCaseSections(payload) {
  const step4 = payload.step4 || {};
  const step5 = payload.step5 || {};

  if (!Array.isArray(step4.cases) || step4.cases.length === 0) {
    const hasLegacyStep4 = Boolean(
      (step4.title && String(step4.title).trim()) ||
      (step4.customerSegment && String(step4.customerSegment).trim()) ||
      (step4.visitPurpose && String(step4.visitPurpose).trim()) ||
      (step4.keyTalk && String(step4.keyTalk).trim()) ||
      (step4.reason && String(step4.reason).trim()) ||
      (step4.other && String(step4.other).trim())
    );
    step4.cases = hasLegacyStep4
      ? [{
          visitReason: String(step4.visitPurpose || ''),
          customerType: mapLegacyCustomerSegment(String(step4.customerSegment || '')),
          talkTag: 'その他',
          talkDetail: String(step4.keyTalk || ''),
          contractFactor: String(step4.reason || ''),
          other: String(step4.other || '')
        }]
      : [createEmptySuccessCase()];
  }

  if (!Array.isArray(step5.cases) || step5.cases.length === 0) {
    const hasLegacyStep5 = Boolean(
      (step5.title && String(step5.title).trim()) ||
      (step5.cause && String(step5.cause).trim()) ||
      (step5.nextAction && String(step5.nextAction).trim()) ||
      (step5.other && String(step5.other).trim())
    );
    step5.cases = hasLegacyStep5
      ? [{
          improvePoint: String(step5.cause || ''),
          reason: String(step5.nextAction || step5.title || ''),
          other: String(step5.other || '')
        }]
      : [createEmptyImproveCase()];
  }

  step4.cases = step4.cases.map((item) => ({ ...createEmptySuccessCase(), ...item }));
  step5.cases = step5.cases.map((item) => ({ ...createEmptyImproveCase(), ...item }));
}

function mapLegacyCustomerSegment(value) {
  const text = String(value || '').trim();
  if (!text) return '';
  if (text === 'ファミリー') return 'ご家族';
  if (text === '単身') return '単身者';
  if (text === 'シニア') return 'ご高齢者';
  if (['ご家族', '単身者', 'ご高齢者', 'その他'].includes(text)) return text;
  return 'その他';
}

function hasFilledSuccessCase(item) {
  if (!item) return false;
  return Boolean(
    String(item.visitReason || '').trim() ||
    String(item.customerType || '').trim() ||
    String(item.talkTag || '').trim() ||
    String(item.talkDetail || '').trim() ||
    String(item.contractFactor || '').trim() ||
    String(item.other || '').trim()
  );
}

function hasFilledImproveCase(item) {
  if (!item) return false;
  return Boolean(
    String(item.improvePoint || '').trim() ||
    String(item.reason || '').trim() ||
    String(item.other || '').trim()
  );
}

function mergeForm(base, incoming) {
  if (typeof incoming !== 'object' || incoming === null) return deepCopy(base);
  const output = deepCopy(base);
  fillObject(output, incoming);
  return output;
}

function fillObject(target, source) {
  Object.keys(target).forEach((key) => {
    if (!(key in source)) return;
    const targetValue = target[key];
    const sourceValue = source[key];

    if (Array.isArray(targetValue)) {
      target[key] = Array.isArray(sourceValue) ? sourceValue : targetValue;
      return;
    }

    if (targetValue && typeof targetValue === 'object' && !Array.isArray(targetValue)) {
      if (sourceValue && typeof sourceValue === 'object' && !Array.isArray(sourceValue)) {
        fillObject(targetValue, sourceValue);
      }
      return;
    }

    if (typeof targetValue === 'number') {
      target[key] = toInt(sourceValue);
      return;
    }

    target[key] = sourceValue == null ? targetValue : String(sourceValue);
  });
}

function saveReports(options = {}) {
  const { silent = false } = options;
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state.reports));
    return true;
  } catch {
    if (!silent) {
      showToast('保存容量を超えました。写真サイズを小さくしてください');
    }
    return false;
  }
}

function loadAdminSession() {
  const id = localStorage.getItem(ADMIN_SESSION_KEY);
  if (!id) return null;
  return ADMIN_USERS.find((user) => user.id === id) || null;
}

function saveAdminSession(user) {
  if (!user) {
    localStorage.removeItem(ADMIN_SESSION_KEY);
    return;
  }
  localStorage.setItem(ADMIN_SESSION_KEY, user.id);
}

function loadSyncConfig() {
  const raw = localStorage.getItem(SYNC_CONFIG_KEY);
  if (!raw) return { ...DEFAULT_SYNC_CONFIG };
  try {
    const parsed = JSON.parse(raw);
    return {
      endpoint: typeof parsed.endpoint === 'string' && parsed.endpoint.trim() ? parsed.endpoint : DEFAULT_SYNC_CONFIG.endpoint,
      token: typeof parsed.token === 'string' && parsed.token.trim() ? parsed.token : DEFAULT_SYNC_CONFIG.token
    };
  } catch {
    return { ...DEFAULT_SYNC_CONFIG };
  }
}

function saveSyncConfig() {
  localStorage.setItem(SYNC_CONFIG_KEY, JSON.stringify(state.syncConfig));
}

function renderSyncConfig() {
  elements.syncEndpointInput.value = state.syncConfig.endpoint;
  elements.syncTokenInput.value = state.syncConfig.token;
  elements.syncStatusText.textContent = state.syncConfig.endpoint
    ? '連携設定済み: 保存・更新時に自動同期します'
    : '未設定: URLを保存するとスプレッドシート同期が有効になります';
}

async function postSync(payload) {
  const endpoint = state.syncConfig.endpoint.trim();
  if (!endpoint) return { skipped: true };

  const response = await fetch(endpoint, {
    method: 'POST',
    // Apps Script Webアプリ向け: JSONヘッダを外してCORS preflightを回避
    body: JSON.stringify({
      ...payload,
      token: state.syncConfig.token || ''
    })
  });

  if (!response.ok) {
    throw new Error(`sync failed: ${response.status}`);
  }
  const result = await response.json().catch(() => ({ ok: true }));
  if (result && result.ok === false) {
    throw new Error(result.error || 'sync rejected');
  }
  return result;
}

async function fetchSyncReports() {
  const endpoint = state.syncConfig.endpoint.trim();
  if (!endpoint) return [];

  const url = new URL(endpoint);
  url.searchParams.set('action', 'list');
  if (state.syncConfig.token) {
    url.searchParams.set('token', state.syncConfig.token);
  }

  const response = await fetch(url.toString(), { method: 'GET' });
  if (!response.ok) {
    throw new Error(`sync fetch failed: ${response.status}`);
  }

  const json = await response.json();
  if (!json || !json.ok || !Array.isArray(json.reports)) return [];
  return json.reports.map(normalizeReport).filter(Boolean);
}

async function initialSyncFromSheet() {
  if (!state.syncConfig.endpoint.trim()) return;
  try {
    const reports = await fetchSyncReports();
    if (reports.length > 0) {
      state.reports = reports;
      saveReports({ silent: true });
    }
  } catch {
    // 起動時は静かにローカルを優先
  }
}

function startSyncPolling() {
  if (state.syncTimer) clearInterval(state.syncTimer);
  if (!state.syncConfig.endpoint.trim()) return;
  state.syncTimer = setInterval(() => {
    void pullReportsFromSheet(false);
  }, SYNC_POLL_INTERVAL_MS);
}

async function pullReportsFromSheet(showToastOnSuccess) {
  if (!state.syncConfig.endpoint.trim()) return;
  try {
    const reports = await fetchSyncReports();
    state.reports = reports;
    saveReports({ silent: true });
    renderStaffList();
    if (state.mode === 'admin') renderAdminView();
    if (state.mode === 'achievements') renderAchievementsView();
    if (showToastOnSuccess) showToast('シートから取得しました');
  } catch {
    if (showToastOnSuccess) showToast('シート取得に失敗しました');
  }
}

function getReportsVersion(reports) {
  if (!Array.isArray(reports) || reports.length === 0) return '0';
  const stamps = reports
    .map((item) => {
      const step1 = item && item.payload && item.payload.step1 ? item.payload.step1 : {};
      return [
        item && item.id ? item.id : '',
        item && item.updatedAt ? item.updatedAt : '',
        step1 && step1.workDate ? step1.workDate : '',
        step1 && step1.staffName ? step1.staffName : '',
        step1 && step1.storeName ? step1.storeName : '',
        step1 && step1.eventVenue ? step1.eventVenue : '',
        item && item.confirmed ? '1' : '0',
        item && item.confirmedBy ? item.confirmedBy : ''
      ].join('|');
    })
    .sort();
  return `${reports.length}:${stamps.join('||')}`;
}

function syncUpsert(report) {
  void postSync({ action: 'upsert', report }).catch(() => {
    showToast('シート同期に失敗しました');
  });
}

function syncDelete(reportId) {
  void postSync({ action: 'delete', reportId }).catch(() => {
    showToast('シート同期に失敗しました');
  });
}

function setHeaderActiveRole(role) {
  elements.switchStaffButton.classList.toggle('is-active', role === 'staff');
  elements.switchAdminButton.classList.toggle('is-active', role === 'admin');
  elements.switchAchievementsButton.classList.toggle('is-active', role === 'achievements');
}

function openStaffListView() {
  state.mode = 'staff-list';
  state.editingId = null;
  state.currentStep = 1;
  state.selectedStaffName = '';
  state.viewLimits.staffGroups = LIST_PAGE_SIZE;
  state.viewLimits.staffReports = LIST_PAGE_SIZE;
  state.errors = {};
  state.photoPreview = null;
  state.achievementsSelectedStaff = '';
  state.achievementsSelectedPeriodKey = '';
  setHeaderActiveRole('staff');
  switchView('staffList');
  renderStaffList();
}

function backToStaffGroupList() {
  state.selectedStaffName = '';
  state.viewLimits.staffReports = LIST_PAGE_SIZE;
  renderStaffList();
}

function openAdminView() {
  state.mode = 'admin';
  state.editingId = null;
  state.currentStep = 1;
  state.adminConfirmedSelectedStaff = '';
  state.adminConfirmedSelectedReportId = '';
  state.viewLimits.adminUnchecked = LIST_PAGE_SIZE;
  state.viewLimits.adminConfirmedStaff = LIST_PAGE_SIZE;
  state.viewLimits.adminConfirmedDates = LIST_PAGE_SIZE;
  state.errors = {};
  state.photoPreview = null;
  setHeaderActiveRole('admin');
  switchView('admin');
  renderAdminView();
}

function openAdminViewKeepSelection() {
  state.mode = 'admin';
  state.editingId = null;
  state.currentStep = 1;
  state.viewLimits.adminUnchecked = LIST_PAGE_SIZE;
  state.errors = {};
  state.photoPreview = null;
  setHeaderActiveRole('admin');
  switchView('admin');
  renderAdminView();
}

function openAchievementsView() {
  if (!state.adminUser) {
    showToast('実績一覧は管理者ログイン後に表示されます');
    openAdminView();
    return;
  }
  state.mode = 'achievements';
  state.editingId = null;
  state.currentStep = 1;
  state.errors = {};
  state.photoPreview = null;
  setHeaderActiveRole('achievements');
  switchView('achievements');
  renderAchievementsView();
}

function renderAdminView() {
  const loggedIn = Boolean(state.adminUser);
  const canManageSync = loggedIn && state.adminUser && state.adminUser.id === 'admin02';
  if (elements.switchAchievementsButton) {
    elements.switchAchievementsButton.style.display = loggedIn ? 'inline-flex' : 'none';
  }

  elements.adminAuthPanel.style.display = loggedIn ? 'none' : 'block';
  elements.adminContent.style.display = loggedIn ? 'block' : 'none';
  if (elements.syncPanel) {
    elements.syncPanel.style.display = canManageSync ? 'block' : 'none';
  }

  if (!loggedIn) return;

  elements.adminUserLabel.textContent = `${state.adminUser.name}でログイン中`;
  renderSyncConfig();
  renderAdminLists();
}

function onAchievementsStaffChange(event) {
  state.achievementsSelectedStaff = String(event.target.value || '');
  state.achievementsSelectedPeriodKey = '';
  renderAchievementsView();
}

function onAchievementsContainerClick(event) {
  const rankingButton = event.target.closest('[data-action="select-achievement-ranking-period"]');
  if (rankingButton) {
    state.achievementsRankingPeriodKey = String(rankingButton.dataset.periodKey || '');
    renderAchievementsView();
    return;
  }

  const hideButton = event.target.closest('[data-action="hide-achievement-comment"]');
  if (hideButton) {
    const reportId = String(hideButton.dataset.reportId || '');
    if (reportId) {
      state.hiddenAchievementCommentIds[reportId] = true;
      saveHiddenAchievementCommentIds();
      renderAchievementsView();
    }
    return;
  }

  const restoreButton = event.target.closest('[data-action="restore-achievement-comments"]');
  if (restoreButton) {
    state.hiddenAchievementCommentIds = {};
    saveHiddenAchievementCommentIds();
    renderAchievementsView();
    return;
  }

  const button = event.target.closest('[data-action="select-achievement-period"]');
  if (!button) return;
  state.achievementsSelectedPeriodKey = String(button.dataset.periodKey || '');
  renderAchievementsView();
}

function onAchievementsContainerInput(event) {
  const editable = event.target.closest('[data-action="edit-report-row-comment"]');
  if (!editable) return;
  const draftKey = String(editable.dataset.draftKey || '');
  const rowId = String(editable.dataset.rowId || '');
  const field = String(editable.dataset.field || '');
  if (!draftKey || !rowId || !field) return;
  updateAchievementReportRowDraft(draftKey, rowId, field, String(editable.textContent || '').trim());
}

function renderAchievementsView() {
  if (!state.adminUser) {
    openAdminView();
    return;
  }

  elements.achievementsTabSummaryButton.classList.toggle('is-active', state.achievementsTab === 'summary');
  elements.achievementsTabRankingButton.classList.toggle('is-active', state.achievementsTab === 'ranking');
  elements.achievementsTabReportButton.classList.toggle('is-active', state.achievementsTab === 'report');
  if (elements.achievementsExportPdfButton) {
    elements.achievementsExportPdfButton.style.display = state.achievementsTab === 'ranking' ? 'none' : 'inline-flex';
  }

  if (state.achievementsTab === 'ranking') {
    elements.achievementsStaffSelect.closest('.panel').style.display = 'none';
    const ranking = summarizeAchievementRanking(state.reports);
    if (!state.achievementsRankingPeriodKey || !ranking.periods.some((p) => p.key === state.achievementsRankingPeriodKey)) {
      state.achievementsRankingPeriodKey = ranking.periods[0] ? ranking.periods[0].key : '';
    }
    elements.achievementsContainer.innerHTML = buildAchievementsRankingHtml(ranking);
    return;
  }

  const canRender = renderAchievementsStaffSelector();
  if (!canRender) {
    elements.achievementsContainer.innerHTML = '<p class="hint">集計できるスタッフがいません。</p>';
    return;
  }

  if (state.achievementsSelectedStaff === ACHIEVEMENTS_ALL_OPTION) {
    if (state.achievementsTab === 'report') {
      elements.achievementsContainer.innerHTML = '<p class="hint">稼働報告はスタッフを選択して表示してください。</p>';
      return;
    }
    const summaryRows = summarizeAchievementAllStaff(state.reports);
    elements.achievementsContainer.innerHTML = buildAchievementsAllStaffHtml(summaryRows);
    return;
  }

  const reports = getAchievementReportsByStaff(state.achievementsSelectedStaff);
  const summary = summarizeAchievementReports(reports);
  if (!state.achievementsSelectedPeriodKey || !summary.periods.some((p) => p.key === state.achievementsSelectedPeriodKey)) {
    state.achievementsSelectedPeriodKey = summary.periods[0] ? summary.periods[0].key : '';
  }

  if (state.achievementsTab === 'report') {
    elements.achievementsContainer.innerHTML = buildAchievementsReportHtml(state.achievementsSelectedStaff, summary);
    return;
  }

  elements.achievementsContainer.innerHTML = buildAchievementsHtml(state.achievementsSelectedStaff, summary);
}

function renderAchievementsStaffSelector() {
  elements.achievementsStaffSelect.closest('.panel').style.display = 'block';
  const staffNames = getAchievementStaffNames();
  const selectValues = [ACHIEVEMENTS_ALL_OPTION].concat(staffNames);
  if (!state.achievementsSelectedStaff || !selectValues.includes(state.achievementsSelectedStaff)) {
    state.achievementsSelectedStaff = selectValues[0] || '';
  }
  if (state.achievementsTab === 'report' && state.achievementsSelectedStaff === ACHIEVEMENTS_ALL_OPTION && staffNames.length > 0) {
    state.achievementsSelectedStaff = staffNames[0];
  }

  const optionsHtml = selectValues
    .map((value) => {
      const selected = value === state.achievementsSelectedStaff ? 'selected' : '';
      const label = value === ACHIEVEMENTS_ALL_OPTION ? 'ALL（全スタッフ）' : value;
      return `<option value="${escapeHtml(value)}" ${selected}>${escapeHtml(label)}</option>`;
    })
    .join('');
  elements.achievementsStaffSelect.innerHTML = optionsHtml;

  if (!state.achievementsSelectedStaff) {
    return false;
  }
  return true;
}

function handleExportAchievementsPdf() {
  if (!state.adminUser) {
    showToast('管理者ログイン後に利用できます');
    return;
  }
  if (state.achievementsTab === 'ranking') {
    showToast('ランキングはPDF保存対象外です');
    return;
  }
  const contentHtml = buildPrintableAchievementsHtml();
  if (!contentHtml) {
    showToast('PDF出力対象がありません');
    return;
  }

  const staffLabel = state.achievementsSelectedStaff === ACHIEVEMENTS_ALL_OPTION
    ? 'ALL（全スタッフ）'
    : (state.achievementsSelectedStaff || '-');
  const documentTitle = state.achievementsTab === 'report' ? '稼働方向 PDF' : 'スタッフ実績 PDF';
  const now = new Date();
  const issuedAt = `${now.getFullYear()}/${String(now.getMonth() + 1).padStart(2, '0')}/${String(now.getDate()).padStart(2, '0')} ${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
  const popup = window.open('', '_blank');
  if (!popup) {
    showToast('ポップアップを許可するとPDF保存できます');
    return;
  }

  popup.document.open();
  popup.document.write(`<!doctype html>
<html lang="ja">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>${escapeHtml(documentTitle)}</title>
  <style>
    :root { --line: #d6deea; --head: #e9f0ff; --text: #1f2937; --hi: #fff6d6; }
    * { box-sizing: border-box; }
    body { margin: 18px; font-family: "Hiragino Kaku Gothic ProN", "Yu Gothic", "Meiryo", sans-serif; color: var(--text); }
    h1 { margin: 0 0 6px; font-size: 18px; }
    .meta { margin: 0 0 12px; font-size: 12px; color: #475569; }
    h3, h4 { margin: 14px 0 6px; }
    .table-wrap { overflow: visible; margin-top: 6px; }
    .summary-table { width: 100%; border-collapse: collapse; min-width: 0; table-layout: fixed; }
    .summary-table th, .summary-table td { border: 1px solid var(--line); padding: 6px 7px; font-size: 11px; background: #fff; word-break: break-word; }
    .summary-table th { background: var(--head); font-weight: 700; text-align: left; }
    .summary-table td.num { text-align: right; }
    .summary-table td.highlight { background: var(--hi); font-weight: 700; }
    .summary-total-row td { background: var(--head); font-weight: 700; }
    .summary-table-comments td { white-space: normal; }
    .report-sheet { border: 1px solid var(--line); border-radius: 10px; padding: 10px; }
    .report-sheet-head { display: flex; justify-content: space-between; align-items: flex-start; gap: 8px; margin-bottom: 10px; }
    .report-sheet-title { margin: 0; font-size: 18px; font-weight: 700; }
    .report-sheet-sub, .report-sheet-date { margin: 4px 0 0; font-size: 12px; color: #475569; }
    .report-grid-2 { display: grid; grid-template-columns: 2fr 1fr; gap: 10px; }
    .report-summary p { margin: 4px 0; line-height: 1.45; }
    @page { size: A4 landscape; margin: 10mm; }
    @media print {
      body { margin: 0; }
      .page-break-avoid { break-inside: avoid; page-break-inside: avoid; }
    }
  </style>
</head>
<body>
  <h1>${escapeHtml(state.achievementsTab === 'report' ? '稼働方向' : 'スタッフ実績')}</h1>
  <p class="meta">スタッフ: ${escapeHtml(staffLabel)} / 出力日時: ${escapeHtml(issuedAt)}</p>
  <div class="page-break-avoid">${contentHtml}</div>
  <script>
    window.addEventListener('load', function () {
      setTimeout(function () { window.print(); }, 250);
    });
  <\/script>
</body>
</html>`);
  popup.document.close();
  popup.focus();
}

function buildPrintableAchievementsHtml() {
  if (!elements.achievementsContainer) return '';
  const source = elements.achievementsContainer;
  const wrapper = document.createElement('div');
  wrapper.innerHTML = source.innerHTML;

  wrapper.querySelectorAll('button, .card-actions, .period-button-row, .period-older-wrap, [data-print-exclude="true"]').forEach((el) => el.remove());

  wrapper.querySelectorAll('table').forEach((table) => {
    const firstHeadRow = table.querySelector('thead tr');
    if (!firstHeadRow) return;
    const ths = Array.from(firstHeadRow.children || []);
    const removeIndexes = [];
    ths.forEach((th, idx) => {
      if (String(th.textContent || '').trim() === '操作') removeIndexes.push(idx);
    });
    if (removeIndexes.length === 0) return;

    const rows = table.querySelectorAll('tr');
    rows.forEach((row) => {
      const cells = Array.from(row.children || []);
      removeIndexes.slice().reverse().forEach((cellIdx) => {
        if (cells[cellIdx]) cells[cellIdx].remove();
      });
    });
  });

  return wrapper.innerHTML.trim();
}

function getAchievementStaffNames() {
  const set = new Set();
  state.reports.forEach((report) => {
    const name = normalizeStaffName(report && report.payload && report.payload.step1 ? report.payload.step1.staffName : '');
    if (name) set.add(name);
  });
  return Array.from(set).sort((a, b) => a.localeCompare(b, 'ja'));
}

function getAchievementReportsByStaff(staffName) {
  return state.reports
    .filter((report) => normalizeStaffName(report && report.payload && report.payload.step1 ? report.payload.step1.staffName : '') === staffName)
    .sort((a, b) => String((a.payload.step1 && a.payload.step1.workDate) || '').localeCompare(String((b.payload.step1 && b.payload.step1.workDate) || '')));
}

function summarizeAchievementReports(reports) {
  const totals = createAchievementTotals();
  const periodMap = new Map();

  reports.forEach((report) => {
    addReportToAchievementTotals(totals, report);
    const workDate = String((report.payload.step1 && report.payload.step1.workDate) || '').trim();
    const parsed = parseYmd(workDate);
    if (!parsed) return;
    const weekStart = getWeekStartWed(parsed);
    const key = formatYmd(weekStart);
    if (!periodMap.has(key)) {
      periodMap.set(key, {
        key,
        dates: new Set(),
        venues: new Set(),
        reports: [],
        totals: createAchievementTotals()
      });
    }
    const p = periodMap.get(key);
    p.dates.add(workDate);
    const venue = String((report.payload.step1 && report.payload.step1.eventVenue) || '').trim();
    if (venue) p.venues.add(venue);
    p.reports.push(report);
    addReportToAchievementTotals(p.totals, report);
  });

  const periods = Array.from(periodMap.values())
    .sort((a, b) => b.key.localeCompare(a.key))
    .map((p) => {
      const catches = p.totals.catchCount;
      const seated = p.totals.seatedCount;
      return {
        key: p.key,
        periodLabel: buildExistingDatesLabel(Array.from(p.dates)),
        venueLabel: Array.from(p.venues).sort((a, b) => a.localeCompare(b, 'ja')).join('、'),
        catches,
        seated,
        contracts: p.totals.contractCount,
        seatedRate: catches > 0 ? seated / catches : 0,
        contractRate: seated > 0 ? p.totals.contractCount / seated : 0,
        commentRows: buildAchievementCommentRows(p.reports),
        reports: p.reports.slice(),
        totals: p.totals
      };
    });

  const catches = totals.catchCount;
  const seated = totals.seatedCount;
  return {
    totals: {
      catches,
      seated,
      contracts: totals.contractCount,
      seatedRate: catches > 0 ? seated / catches : 0,
      contractRate: seated > 0 ? totals.contractCount / seated : 0
    },
    periods,
    items: getAchievementItems()
  };
}

function summarizeAchievementAllStaff(reports) {
  const map = new Map();
  (Array.isArray(reports) ? reports : []).forEach((report) => {
    const name = normalizeStaffName(report && report.payload && report.payload.step1 ? report.payload.step1.staffName : '');
    if (!name) return;
    if (!map.has(name)) {
      map.set(name, { staffName: name, totals: createAchievementTotals(), latestWorkDate: '' });
    }
    const row = map.get(name);
    addReportToAchievementTotals(row.totals, report);
    const workDate = String((report.payload && report.payload.step1 && report.payload.step1.workDate) || '');
    if (workDate && (!row.latestWorkDate || workDate > row.latestWorkDate)) {
      row.latestWorkDate = workDate;
    }
  });

  return Array.from(map.values())
    .map((row) => {
      const catches = row.totals.catchCount;
      const seated = row.totals.seatedCount;
      return {
        staffName: row.staffName,
        latestWorkDate: row.latestWorkDate,
        catches,
        seated,
        contracts: row.totals.contractCount,
        seatedRate: catches > 0 ? seated / catches : 0,
        contractRate: seated > 0 ? row.totals.contractCount / seated : 0
      };
    })
    .sort((a, b) => b.contracts - a.contracts || b.catches - a.catches || a.staffName.localeCompare(b.staffName, 'ja'));
}

function buildAchievementsAllStaffHtml(rows) {
  const body = (Array.isArray(rows) ? rows : [])
    .map((row) => `
      <tr>
        <td>${escapeHtml(row.staffName)}</td>
        <td>${escapeHtml(row.latestWorkDate || '-')}</td>
        <td class="num">${row.catches}</td>
        <td class="num">${row.seated}</td>
        <td class="num">${row.contracts}</td>
        <td class="num">${formatPercent(row.seatedRate)}</td>
        <td class="num">${formatPercent(row.contractRate)}</td>
      </tr>
    `)
    .join('');

  return `
    <h3>全スタッフ全体集計</h3>
    <p class="hint">ALL選択時は、期間詳細を表示せずスタッフ別の全体集計のみを表示します。</p>
    <div class="table-wrap">
      <table class="summary-table">
        <thead>
          <tr>
            <th>スタッフ名</th>
            <th>最新稼働日</th>
            <th>キャッチ数</th>
            <th>着座数</th>
            <th>成約台数合計</th>
            <th>着座率</th>
            <th>成約率</th>
          </tr>
        </thead>
        <tbody>${body || '<tr><td colspan="7">データなし</td></tr>'}</tbody>
      </table>
    </div>
  `;
}

function createAchievementTotals() {
  return {
    catchCount: 0,
    seatedCount: 0,
    contractCount: 0,
    newA: {
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
      auHikariNew: 0,
      auHikariFromDocomo: 0,
      auHikariFromSoftbank: 0,
      auHikariFromOther: 0,
      blHikariNew: 0,
      blHikariFromDocomo: 0,
      blHikariFromSoftbank: 0,
      blHikariFromOther: 0,
      commufaHikariNew: 0,
      commufaHikariFromDocomo: 0,
      commufaHikariFromSoftbank: 0,
      commufaHikariFromOther: 0
    }
  };
}

function addReportToAchievementTotals(totals, report) {
  const step2 = (report.payload && report.payload.step2) || {};
  const step3 = (report.payload && report.payload.step3) || {};
  const newA = step3.newAcquisitions || {};
  const ltv = step3.ltv || {};
  const auH = ltv.auHikariBreakdown || {};
  const blH = ltv.blHikariBreakdown || {};
  const cmH = ltv.commufaHikariBreakdown || {};

  totals.catchCount += toInt(step2.catchCount);
  totals.seatedCount += toInt(step2.seated);

  totals.newA.auMnpSim += toInt(newA.auMnpSim);
  totals.newA.auMnpHs += toInt(newA.auMnpHs);
  totals.newA.auNewSim += toInt(newA.auNewSim);
  totals.newA.auNewHs += toInt(newA.auNewHs);
  totals.newA.uqMnpSim += toInt(newA.uqMnpSim);
  totals.newA.uqMnpHs += toInt(newA.uqMnpHs);
  totals.newA.uqNewSim += toInt(newA.uqNewSim);
  totals.newA.uqNewHs += toInt(newA.uqNewHs);
  totals.newA.cellUp += toInt(newA.cellUp);

  totals.contractCount +=
    toInt(newA.auMnpSim) +
    toInt(newA.auMnpHs) +
    toInt(newA.auNewSim) +
    toInt(newA.auNewHs) +
    toInt(newA.uqMnpSim) +
    toInt(newA.uqMnpHs) +
    toInt(newA.uqNewSim) +
    toInt(newA.uqNewHs);

  totals.ltv.auDenki += toInt(ltv.auDenki);
  totals.ltv.goldCard += toInt(ltv.goldCard);
  totals.ltv.silverCard += toInt(ltv.silverCard);
  totals.ltv.rankUp += toInt(ltv.rankUp);
  totals.ltv.jibunBank += toInt(ltv.jibunBank);
  totals.ltv.norton += toInt(ltv.norton);
  totals.ltv.auHikariNew += toInt(auH.new);
  totals.ltv.auHikariFromDocomo += toInt(auH.fromDocomo);
  totals.ltv.auHikariFromSoftbank += toInt(auH.fromSoftbank);
  totals.ltv.auHikariFromOther += toInt(auH.fromOther);
  totals.ltv.blHikariNew += toInt(blH.new);
  totals.ltv.blHikariFromDocomo += toInt(blH.fromDocomo);
  totals.ltv.blHikariFromSoftbank += toInt(blH.fromSoftbank);
  totals.ltv.blHikariFromOther += toInt(blH.fromOther);
  totals.ltv.commufaHikariNew += toInt(cmH.new);
  totals.ltv.commufaHikariFromDocomo += toInt(cmH.fromDocomo);
  totals.ltv.commufaHikariFromSoftbank += toInt(cmH.fromSoftbank);
  totals.ltv.commufaHikariFromOther += toInt(cmH.fromOther);
}

function getAchievementItems() {
  return [
    { label: 'au MNP SIM単', get: (t) => t.newA.auMnpSim },
    { label: 'au MNP HS', get: (t) => t.newA.auMnpHs },
    { label: 'au純新規 SIM単', get: (t) => t.newA.auNewSim },
    { label: 'au純新規 HS', get: (t) => t.newA.auNewHs },
    { label: 'UQ MNP SIM単', get: (t) => t.newA.uqMnpSim },
    { label: 'UQ MNP HS', get: (t) => t.newA.uqMnpHs },
    { label: 'UQ純新規 SIM単', get: (t) => t.newA.uqNewSim },
    { label: 'UQ純新規 HS', get: (t) => t.newA.uqNewHs },
    { label: 'セルアップ', get: (t) => t.newA.cellUp },
    { label: 'auでんき', get: (t) => t.ltv.auDenki },
    { label: 'ゴールドカード', get: (t) => t.ltv.goldCard },
    { label: 'シルバーカード', get: (t) => t.ltv.silverCard },
    { label: 'ランクアップ', get: (t) => t.ltv.rankUp },
    { label: 'じぶん銀行', get: (t) => t.ltv.jibunBank },
    { label: 'ノートン', get: (t) => t.ltv.norton },
    { label: 'auひかり 新規', get: (t) => t.ltv.auHikariNew },
    { label: 'auひかり ドコモ光から切替', get: (t) => t.ltv.auHikariFromDocomo },
    { label: 'auひかり ソフトバンク光から切替', get: (t) => t.ltv.auHikariFromSoftbank },
    { label: 'auひかり その他から切替', get: (t) => t.ltv.auHikariFromOther },
    { label: 'BLひかり 新規', get: (t) => t.ltv.blHikariNew },
    { label: 'BLひかり ドコモ光から切替', get: (t) => t.ltv.blHikariFromDocomo },
    { label: 'BLひかり ソフトバンク光から切替', get: (t) => t.ltv.blHikariFromSoftbank },
    { label: 'BLひかり その他から切替', get: (t) => t.ltv.blHikariFromOther },
    { label: 'コミュファ光 新規', get: (t) => t.ltv.commufaHikariNew },
    { label: 'コミュファ光 ドコモ光から切替', get: (t) => t.ltv.commufaHikariFromDocomo },
    { label: 'コミュファ光 ソフトバンク光から切替', get: (t) => t.ltv.commufaHikariFromSoftbank },
    { label: 'コミュファ光 その他から切替', get: (t) => t.ltv.commufaHikariFromOther }
  ];
}

function getAchievementDailyNewItems() {
  return [
    { label: 'au MNP SIM単', get: (t) => t.newA.auMnpSim },
    { label: 'au MNP HS', get: (t) => t.newA.auMnpHs },
    { label: 'au純新規 SIM単', get: (t) => t.newA.auNewSim },
    { label: 'au純新規 HS', get: (t) => t.newA.auNewHs },
    { label: 'UQ MNP SIM単', get: (t) => t.newA.uqMnpSim },
    { label: 'UQ MNP HS', get: (t) => t.newA.uqMnpHs },
    { label: 'UQ純新規 SIM単', get: (t) => t.newA.uqNewSim },
    { label: 'UQ純新規 HS', get: (t) => t.newA.uqNewHs },
    { label: 'セルアップ', get: (t) => t.newA.cellUp }
  ];
}

function getAchievementDailyLtvItems() {
  return [
    { label: 'auでんき', get: (t) => t.ltv.auDenki },
    { label: 'ゴールドカード', get: (t) => t.ltv.goldCard },
    { label: 'シルバーカード', get: (t) => t.ltv.silverCard },
    { label: 'ランクアップ', get: (t) => t.ltv.rankUp },
    { label: 'じぶん銀行', get: (t) => t.ltv.jibunBank },
    { label: 'ノートン', get: (t) => t.ltv.norton },
    { label: 'auひかり 新規', get: (t) => t.ltv.auHikariNew },
    { label: 'auひかり ドコモ光から切替', get: (t) => t.ltv.auHikariFromDocomo },
    { label: 'auひかり ソフトバンク光から切替', get: (t) => t.ltv.auHikariFromSoftbank },
    { label: 'auひかり その他から切替', get: (t) => t.ltv.auHikariFromOther },
    { label: 'BLひかり 新規', get: (t) => t.ltv.blHikariNew },
    { label: 'BLひかり ドコモ光から切替', get: (t) => t.ltv.blHikariFromDocomo },
    { label: 'BLひかり ソフトバンク光から切替', get: (t) => t.ltv.blHikariFromSoftbank },
    { label: 'BLひかり その他から切替', get: (t) => t.ltv.blHikariFromOther },
    { label: 'コミュファ光 新規', get: (t) => t.ltv.commufaHikariNew },
    { label: 'コミュファ光 ドコモ光から切替', get: (t) => t.ltv.commufaHikariFromDocomo },
    { label: 'コミュファ光 ソフトバンク光から切替', get: (t) => t.ltv.commufaHikariFromSoftbank },
    { label: 'コミュファ光 その他から切替', get: (t) => t.ltv.commufaHikariFromOther }
  ];
}

function buildAchievementDailyBreakdown(reports) {
  const byDate = new Map();
  (Array.isArray(reports) ? reports : []).forEach((report) => {
    const step1 = (report.payload && report.payload.step1) || {};
    const workDate = String(step1.workDate || '').trim();
    if (!parseYmd(workDate)) return;
    if (!byDate.has(workDate)) byDate.set(workDate, createAchievementTotals());
    addReportToAchievementTotals(byDate.get(workDate), report);
  });

  return {
    dates: Array.from(byDate.keys()).sort(),
    byDate
  };
}

function formatMd(ymd) {
  const date = parseYmd(ymd);
  if (!date) return ymd;
  return `${String(date.getMonth() + 1).padStart(2, '0')}/${String(date.getDate()).padStart(2, '0')}`;
}

function buildAchievementDailyTableHtml(title, dailyDates, byDate, items, totalTotals) {
  const headHtml = ['日付'].concat(items.map((item) => escapeHtml(item.label))).map((v) => `<th>${v}</th>`).join('');
  const bodyRows = dailyDates
    .map((ymd) => {
      const dayTotals = byDate.get(ymd) || createAchievementTotals();
      const valueCells = items
        .map((item) => {
          const v = toInt(item.get(dayTotals));
          const cls = v > 0 ? 'num highlight' : 'num';
          return `<td class="${cls}">${v}</td>`;
        })
        .join('');
      return `<tr><td>${escapeHtml(formatMd(ymd))}</td>${valueCells}</tr>`;
    })
    .join('');

  const totalCells = items
    .map((item) => {
      const v = toInt(item.get(totalTotals || createAchievementTotals()));
      const cls = v > 0 ? 'num highlight' : 'num';
      return `<td class="${cls}">${v}</td>`;
    })
    .join('');

  return `
    <h4>${escapeHtml(title)}</h4>
    <div class="table-wrap">
      <table class="summary-table">
        <thead><tr>${headHtml}</tr></thead>
        <tbody>
          ${bodyRows || `<tr><td colspan="${items.length + 1}">データなし</td></tr>`}
          <tr class="summary-total-row"><td>合計</td>${totalCells}</tr>
        </tbody>
      </table>
    </div>
  `;
}

function buildAchievementsHtml(staffName, summary) {
  const overall = summary.totals;
  const selectedPeriod = summary.periods.find((p) => p.key === state.achievementsSelectedPeriodKey) || null;
  const periodSelectorHtml = buildPeriodSelectorHtml(summary.periods, selectedPeriod);
  const dailyNewItems = getAchievementDailyNewItems();
  const dailyLtvItems = getAchievementDailyLtvItems();
  const dailyBreakdown = selectedPeriod ? buildAchievementDailyBreakdown(selectedPeriod.reports) : { dates: [], byDate: new Map() };
  const detailHtml = selectedPeriod
    ? [
      buildAchievementDailyTableHtml('【日別　新規成約実績】', dailyBreakdown.dates, dailyBreakdown.byDate, dailyNewItems, selectedPeriod.totals),
      buildAchievementDailyTableHtml('【日別　LTV成約実績】', dailyBreakdown.dates, dailyBreakdown.byDate, dailyLtvItems, selectedPeriod.totals)
    ].join('')
    : '<p class="hint">データなし</p>';

  const visibleCommentRows = selectedPeriod
    ? selectedPeriod.commentRows.filter((row) => !state.hiddenAchievementCommentIds[row.reportId])
    : [];
  const hiddenCommentCount = selectedPeriod ? Math.max(0, selectedPeriod.commentRows.length - visibleCommentRows.length) : 0;
  const commentRows = selectedPeriod
    ? visibleCommentRows
      .map((row) => `
        <tr>
          <td>${escapeHtml(row.dateLabel)}</td>
          <td>${escapeHtml(row.successComment)}</td>
          <td>${escapeHtml(row.improveComment)}</td>
          <td>${escapeHtml(row.reflectionComment)}</td>
          <td><button type="button" class="btn btn-ghost btn-small" data-action="hide-achievement-comment" data-report-id="${escapeHtml(row.reportId)}">非表示</button></td>
        </tr>
      `)
      .join('')
    : '';

  return `
    <h3>全体集計</h3>
    <div class="table-wrap">
      <table class="summary-table">
        <thead>
          <tr>
            <th>スタッフ名</th><th>キャッチ数</th><th>着座数</th><th>成約台数合計</th><th>着座率</th><th>成約率</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td>${escapeHtml(staffName)}</td>
            <td class="num">${overall.catches}</td>
            <td class="num">${overall.seated}</td>
            <td class="num">${overall.contracts}</td>
            <td class="num">${formatPercent(overall.seatedRate)}</td>
            <td class="num">${formatPercent(overall.contractRate)}</td>
          </tr>
        </tbody>
      </table>
    </div>
    <h3>期間集計</h3>
    ${periodSelectorHtml}
    <div class="table-wrap">
      <table class="summary-table">
        <thead>
          <tr>
            <th>期間</th><th>イベント会場</th><th>キャッチ数</th><th>着座数</th><th>成約台数合計</th><th>着座率</th><th>成約率</th>
          </tr>
        </thead>
        <tbody>${selectedPeriod ? `
          <tr>
            <td>${escapeHtml(selectedPeriod.periodLabel || '-')}</td>
            <td>${escapeHtml(selectedPeriod.venueLabel || '-')}</td>
            <td class="num">${selectedPeriod.catches}</td>
            <td class="num">${selectedPeriod.seated}</td>
            <td class="num">${selectedPeriod.contracts}</td>
            <td class="num">${formatPercent(selectedPeriod.seatedRate)}</td>
            <td class="num">${formatPercent(selectedPeriod.contractRate)}</td>
          </tr>
        ` : '<tr><td colspan="7">データなし</td></tr>'}</tbody>
      </table>
    </div>
    <h3>期間内訳</h3>
    ${detailHtml}
    <h3>期間コメント（成功・改善・振り返り）</h3>
    <div class="card-actions">
      <button type="button" class="btn btn-ghost" data-action="restore-achievement-comments">非表示コメントをすべて戻す</button>
      ${hiddenCommentCount > 0 ? `<p class="hint">非表示中: ${hiddenCommentCount}件</p>` : ''}
    </div>
    <div class="table-wrap">
      <table class="summary-table summary-table-comments">
        <thead>
          <tr>
            <th>入力日付（曜日）</th>
            <th>成功コメント</th>
            <th>改善コメント</th>
            <th>振り返りコメント</th>
            <th>操作</th>
          </tr>
        </thead>
        <tbody>${commentRows || '<tr><td colspan="5">データなし</td></tr>'}</tbody>
      </table>
    </div>
  `;
}

function buildAchievementsReportHtml(staffName, summary) {
  const selectedPeriod = summary.periods.find((p) => p.key === state.achievementsSelectedPeriodKey) || null;
  const periodSelectorHtml = buildPeriodSelectorHtml(summary.periods, selectedPeriod);
  if (!selectedPeriod) {
    return `
      ${periodSelectorHtml}
      <p class="hint">期間データがありません。</p>
    `;
  }

  const dailyBreakdown = buildAchievementDailyBreakdown(selectedPeriod.reports || []);
  const dailyNewItems = getAchievementDailyNewItems();
  const dailyLtvItems = getAchievementDailyLtvItems();
  const achievedUnits = toInt(selectedPeriod.contracts);
  const reportDate = formatYmd(new Date());
  const allCommentRows = selectedPeriod.commentRows || [];
  const draftKey = buildAchievementReportDraftKey(staffName, selectedPeriod.key);
  const draftRows = getOrInitAchievementReportRowDraft(draftKey, allCommentRows);

  return `
    ${periodSelectorHtml}
    <div class="report-sheet" data-print-keep="true">
      <div class="report-sheet-head">
        <div>
          <p class="report-sheet-title">稼働方向</p>
          <p class="report-sheet-sub">スタッフ: ${escapeHtml(staffName)}</p>
        </div>
        <p class="report-sheet-date">${escapeHtml(reportDate)}</p>
      </div>

      <div class="report-basic">
        <h4>【基本情報】</h4>
        <div class="table-wrap">
          <table class="summary-table report-table-compact">
            <tbody>
              <tr><th>稼働期間</th><td>${escapeHtml(selectedPeriod.periodLabel || '-')}</td></tr>
              <tr><th>稼働場所</th><td>${escapeHtml(selectedPeriod.venueLabel || '-')}</td></tr>
              <tr><th>実績</th><td>新規 ${achievedUnits} 台</td></tr>
            </tbody>
          </table>
        </div>
      </div>

      ${buildAchievementDailyTableHtml('【日別　新規実績】', dailyBreakdown.dates, dailyBreakdown.byDate, dailyNewItems, selectedPeriod.totals)}

      ${buildAchievementDailyTableHtml('【日別　LTV実績】', dailyBreakdown.dates, dailyBreakdown.byDate, dailyLtvItems, selectedPeriod.totals)}

      <div class="report-summary">
        <h4>【総括】</h4>
        <p class="hint" data-print-exclude="true">コメントセルを直接編集できます（元の日報データは変更されません）。</p>
        ${buildReportCommentTableHtml('成功コメント', draftRows, 'successComment', draftKey)}
        ${buildReportCommentTableHtml('改善コメント', draftRows, 'improveComment', draftKey)}
        ${buildReportCommentTableHtml('振り返りコメント', draftRows, 'reflectionComment', draftKey)}
      </div>
    </div>
  `;
}

function buildAchievementReportDraftKey(staffName, periodKey) {
  return `${String(staffName || '').trim()}::${String(periodKey || '').trim()}`;
}

function getOrInitAchievementReportRowDraft(draftKey, commentRows) {
  const existing = state.achievementsReportDrafts[draftKey];
  if (existing && Array.isArray(existing.rows)) return existing.rows;
  const rows = Array.isArray(commentRows) ? commentRows : [];
  const draftRows = rows.map((row, idx) => ({
    rowId: String(row.reportId || `row-${idx + 1}`),
    dateLabel: String(row.dateLabel || '-'),
    successComment: String(row.successComment || '-'),
    improveComment: String(row.improveComment || '-'),
    reflectionComment: String(row.reflectionComment || '-')
  }));
  state.achievementsReportDrafts[draftKey] = { rows: draftRows };
  return draftRows;
}

function updateAchievementReportRowDraft(draftKey, rowId, field, value) {
  if (!state.achievementsReportDrafts[draftKey] || !Array.isArray(state.achievementsReportDrafts[draftKey].rows)) return;
  const rows = state.achievementsReportDrafts[draftKey].rows;
  const row = rows.find((r) => r.rowId === rowId);
  if (!row) return;
  if (!['successComment', 'improveComment', 'reflectionComment'].includes(field)) return;
  row[field] = value || '-';
}

function buildReportCommentTableHtml(title, rows, field, draftKey) {
  const body = (Array.isArray(rows) ? rows : [])
    .map((row) => `
      <tr>
        <td>${escapeHtml(row.dateLabel || '-')}</td>
        <td>
          <div
            class="report-editable-cell"
            contenteditable="true"
            data-action="edit-report-row-comment"
            data-draft-key="${escapeHtml(draftKey)}"
            data-row-id="${escapeHtml(String(row.rowId || ''))}"
            data-field="${escapeHtml(field)}"
          >${escapeHtml(String(row[field] || '-'))}</div>
        </td>
      </tr>
    `)
    .join('');

  return `
    <h5>${escapeHtml(title)}</h5>
    <div class="table-wrap">
      <table class="summary-table summary-table-comments">
        <thead><tr><th>入力日付（曜日）</th><th>${escapeHtml(title)}</th></tr></thead>
        <tbody>${body || '<tr><td colspan="2">データなし</td></tr>'}</tbody>
      </table>
    </div>
  `;
}

function buildPeriodSelectorHtml(periods, selectedPeriod) {
  if (!Array.isArray(periods) || periods.length === 0) {
    return '<p class="hint">期間データなし</p>';
  }

  const primary = periods.slice(0, ACHIEVEMENT_PERIOD_PRIMARY_COUNT);
  const olders = periods.slice(ACHIEVEMENT_PERIOD_PRIMARY_COUNT);
  const renderButtons = (list) =>
    list
      .map((p) => {
        const active = selectedPeriod && selectedPeriod.key === p.key ? 'is-active' : '';
        return `<button type="button" class="btn btn-outline btn-period ${active}" data-action="select-achievement-period" data-period-key="${escapeHtml(p.key)}">${escapeHtml(p.periodLabel || '-')}</button>`;
      })
      .join('');

  const primaryHtml = `<div class="period-button-row period-button-row-primary">${renderButtons(primary)}</div>`;
  if (olders.length === 0) return primaryHtml;

  return `
    ${primaryHtml}
    <details class="period-older-wrap">
      <summary>過去期間を表示（${olders.length}件）</summary>
      <div class="period-button-row period-button-row-older">${renderButtons(olders)}</div>
    </details>
  `;
}

function parseYmd(value) {
  const text = String(value || '').trim();
  const m = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  const date = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  if (!isFinite(date.getTime())) return null;
  return date;
}

function getWeekStartWed(date) {
  const offset = (date.getDay() - 3 + 7) % 7;
  return new Date(date.getFullYear(), date.getMonth(), date.getDate() - offset);
}

function formatYmd(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

function buildExistingDatesLabel(dateList) {
  const sorted = (dateList || []).slice().sort();
  return sorted.map((d) => formatMdWithWeekday(d)).join(' / ');
}

function formatMdWithWeekday(ymd) {
  const date = parseYmd(ymd);
  if (!date) return ymd;
  const md = `${String(date.getMonth() + 1).padStart(2, '0')}/${String(date.getDate()).padStart(2, '0')}`;
  const days = ['日', '月', '火', '水', '木', '金', '土'];
  return `${md}(${days[date.getDay()]})`;
}

function formatYmdWithWeekday(ymd) {
  const date = parseYmd(ymd);
  if (!date) return ymd;
  const days = ['日', '月', '火', '水', '木', '金', '土'];
  return `${formatYmd(date)}(${days[date.getDay()]})`;
}

function buildAchievementCommentRows(reports) {
  return (Array.isArray(reports) ? reports : [])
    .slice()
    .sort((a, b) => String((a.payload.step1 && a.payload.step1.workDate) || '').localeCompare(String((b.payload.step1 && b.payload.step1.workDate) || '')))
    .map((report) => {
      const step1 = (report.payload && report.payload.step1) || {};
      const step4 = (report.payload && report.payload.step4) || {};
      const step5 = (report.payload && report.payload.step5) || {};
      const step6 = (report.payload && report.payload.step6) || {};

      const successCase = (Array.isArray(step4.cases) ? step4.cases : []).find(hasFilledSuccessCase) || null;
      const improveCase = (Array.isArray(step5.cases) ? step5.cases : []).find(hasFilledImproveCase) || null;

      const successComment = successCase
        ? [successCase.visitReason, successCase.contractFactor, successCase.talkDetail].map((v) => String(v || '').trim()).filter(Boolean).join(' / ') || '-'
        : '-';
      const improveComment = improveCase
        ? [improveCase.improvePoint, improveCase.reason, improveCase.other].map((v) => String(v || '').trim()).filter(Boolean).join(' / ') || '-'
        : '-';
      const reflectionComment = [step6.impression, step6.notes]
        .map((v) => String(v || '').trim())
        .filter(Boolean)
        .join(' / ') || '-';

      return {
        reportId: String(report.id || ''),
        dateLabel: formatYmdWithWeekday(String(step1.workDate || '')),
        successComment,
        improveComment,
        reflectionComment
      };
    });
}

function loadHiddenAchievementCommentIds() {
  const raw = localStorage.getItem(ACHIEVEMENTS_COMMENT_HIDDEN_KEY);
  if (!raw) return {};
  try {
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== 'object') return {};
    return parsed;
  } catch {
    return {};
  }
}

function saveHiddenAchievementCommentIds() {
  localStorage.setItem(ACHIEVEMENTS_COMMENT_HIDDEN_KEY, JSON.stringify(state.hiddenAchievementCommentIds || {}));
}

function formatPercent(value) {
  const n = Number(value || 0);
  return `${(n * 100).toFixed(1)}%`;
}

function summarizeAchievementRanking(reports) {
  const weekMap = new Map();
  const list = Array.isArray(reports) ? reports : [];

  list.forEach((report) => {
    const step1 = (report.payload && report.payload.step1) || {};
    const step3 = (report.payload && report.payload.step3) || {};
    const newA = step3.newAcquisitions || {};
    const ltv = step3.ltv || {};
    const workDate = String(step1.workDate || '').trim();
    const parsed = parseYmd(workDate);
    if (!parsed) return;
    const weekStart = getWeekStartWed(parsed);
    const key = formatYmd(weekStart);
    if (!weekMap.has(key)) {
      weekMap.set(key, { key, dates: new Set(), staff: {} });
    }
    const bucket = weekMap.get(key);
    bucket.dates.add(workDate);

    const staffName = normalizeStaffName(step1.staffName || '');
    if (!staffName) return;
    if (!bucket.staff[staffName]) {
      bucket.staff[staffName] = {
        newTotal: 0,
        ltv: {
          auDenki: 0, goldCard: 0, silverCard: 0, rankUp: 0, jibunBank: 0, norton: 0,
          auHikariNew: 0, auHikariFromDocomo: 0, auHikariFromSoftbank: 0, auHikariFromOther: 0,
          blHikariNew: 0, blHikariFromDocomo: 0, blHikariFromSoftbank: 0, blHikariFromOther: 0,
          commufaHikariNew: 0, commufaHikariFromDocomo: 0, commufaHikariFromSoftbank: 0, commufaHikariFromOther: 0
        }
      };
    }
    const row = bucket.staff[staffName];
    const auH = ltv.auHikariBreakdown || {};
    const blH = ltv.blHikariBreakdown || {};
    const cmH = ltv.commufaHikariBreakdown || {};
    row.newTotal +=
      toInt(newA.auMnpSim) + toInt(newA.auMnpHs) + toInt(newA.auNewSim) + toInt(newA.auNewHs) +
      toInt(newA.uqMnpSim) + toInt(newA.uqMnpHs) + toInt(newA.uqNewSim) + toInt(newA.uqNewHs) + toInt(newA.cellUp);

    row.ltv.auDenki += toInt(ltv.auDenki);
    row.ltv.goldCard += toInt(ltv.goldCard);
    row.ltv.silverCard += toInt(ltv.silverCard);
    row.ltv.rankUp += toInt(ltv.rankUp);
    row.ltv.jibunBank += toInt(ltv.jibunBank);
    row.ltv.norton += toInt(ltv.norton);
    row.ltv.auHikariNew += toInt(auH.new);
    row.ltv.auHikariFromDocomo += toInt(auH.fromDocomo);
    row.ltv.auHikariFromSoftbank += toInt(auH.fromSoftbank);
    row.ltv.auHikariFromOther += toInt(auH.fromOther);
    row.ltv.blHikariNew += toInt(blH.new);
    row.ltv.blHikariFromDocomo += toInt(blH.fromDocomo);
    row.ltv.blHikariFromSoftbank += toInt(blH.fromSoftbank);
    row.ltv.blHikariFromOther += toInt(blH.fromOther);
    row.ltv.commufaHikariNew += toInt(cmH.new);
    row.ltv.commufaHikariFromDocomo += toInt(cmH.fromDocomo);
    row.ltv.commufaHikariFromSoftbank += toInt(cmH.fromSoftbank);
    row.ltv.commufaHikariFromOther += toInt(cmH.fromOther);
  });

  const periods = Array.from(weekMap.values())
    .sort((a, b) => b.key.localeCompare(a.key))
    .map((bucket) => ({
      key: bucket.key,
      label: buildExistingDatesLabel(Array.from(bucket.dates)),
      staffRows: Object.entries(bucket.staff).map(([staffName, v]) => ({ staffName, ...v }))
    }));

  return { periods };
}

function buildAchievementsRankingHtml(ranking) {
  const periods = ranking.periods || [];
  if (periods.length === 0) return '<p class="hint">ランキング対象データがありません。</p>';
  const selected = periods.find((p) => p.key === state.achievementsRankingPeriodKey) || periods[0];

  const periodButtons = periods
    .map((p) => {
      const active = p.key === selected.key ? 'is-active' : '';
      return `<button type="button" class="btn btn-outline btn-period ${active}" data-action="select-achievement-ranking-period" data-period-key="${escapeHtml(p.key)}">${escapeHtml(p.label)}</button>`;
    })
    .join('');

  const newRows = selected.staffRows
    .slice()
    .sort((a, b) => b.newTotal - a.newTotal || a.staffName.localeCompare(b.staffName, 'ja'))
    .map((row, idx) => `<tr><td>${idx + 1}</td><td>${escapeHtml(row.staffName)}</td><td class="num">${row.newTotal}</td></tr>`)
    .join('');

  const ltvItems = [
    ['auでんき', 'auDenki'], ['ゴールドカード', 'goldCard'], ['シルバーカード', 'silverCard'], ['ランクアップ', 'rankUp'],
    ['じぶん銀行', 'jibunBank'], ['ノートン', 'norton'],
    ['auひかり 新規', 'auHikariNew'], ['auひかり ドコモ光から切替', 'auHikariFromDocomo'], ['auひかり ソフトバンク光から切替', 'auHikariFromSoftbank'], ['auひかり その他から切替', 'auHikariFromOther'],
    ['BLひかり 新規', 'blHikariNew'], ['BLひかり ドコモ光から切替', 'blHikariFromDocomo'], ['BLひかり ソフトバンク光から切替', 'blHikariFromSoftbank'], ['BLひかり その他から切替', 'blHikariFromOther'],
    ['コミュファ光 新規', 'commufaHikariNew'], ['コミュファ光 ドコモ光から切替', 'commufaHikariFromDocomo'], ['コミュファ光 ソフトバンク光から切替', 'commufaHikariFromSoftbank'], ['コミュファ光 その他から切替', 'commufaHikariFromOther']
  ];

  const ltvHtml = ltvItems
    .map(([label, key]) => {
      const rows = selected.staffRows
        .map((row) => ({ staffName: row.staffName, value: toInt(row.ltv[key]) }))
        .filter((r) => r.value > 0)
        .sort((a, b) => b.value - a.value || a.staffName.localeCompare(b.staffName, 'ja'));
      const body = rows.length > 0
        ? rows.map((r, idx) => `<tr><td>${idx + 1}</td><td>${escapeHtml(r.staffName)}</td><td class="num">${r.value}</td></tr>`).join('')
        : '<tr><td colspan="3">データなし</td></tr>';
      return `
        <details class="ranking-item">
          <summary>${escapeHtml(label)} ランキング</summary>
          <div class="table-wrap">
            <table class="summary-table">
              <thead><tr><th>順位</th><th>スタッフ</th><th>件数</th></tr></thead>
              <tbody>${body}</tbody>
            </table>
          </div>
        </details>
      `;
    })
    .join('');

  return `
    <h3>対象期間</h3>
    <div class="period-button-row">${periodButtons}</div>
    <h3>新規獲得（全項目合計）ランキング</h3>
    <div class="table-wrap">
      <table class="summary-table">
        <thead><tr><th>順位</th><th>スタッフ</th><th>台数</th></tr></thead>
        <tbody>${newRows || '<tr><td colspan="3">データなし</td></tr>'}</tbody>
      </table>
    </div>
    <h3>LTV項目ランキング</h3>
    <div class="ranking-stack">${ltvHtml}</div>
  `;
}

function handleSaveSyncConfig() {
  if (!state.adminUser || state.adminUser.id !== 'admin02') {
    showToast('連携設定は小澤アカウントのみ変更できます');
    return;
  }
  state.syncConfig.endpoint = elements.syncEndpointInput.value.trim();
  state.syncConfig.token = elements.syncTokenInput.value.trim();
  saveSyncConfig();
  renderSyncConfig();
  startSyncPolling();
  showToast('連携設定を保存しました');
}

async function handleSyncAllReports() {
  if (!state.adminUser || state.adminUser.id !== 'admin02') {
    showToast('シート同期は小澤アカウントのみ実行できます');
    return;
  }
  if (!state.syncConfig.endpoint.trim()) {
    showToast('先にApps Script URLを保存してください');
    return;
  }
  try {
    await postSync({ action: 'replaceAll', reports: state.reports });
    showToast('全件をシートへ同期しました');
  } catch {
    showToast('全件同期に失敗しました');
  }
}

async function handleSyncPullReports() {
  if (!state.adminUser || state.adminUser.id !== 'admin02') {
    showToast('シート取得は小澤アカウントのみ実行できます');
    return;
  }
  if (!state.syncConfig.endpoint.trim()) {
    showToast('先にApps Script URLを保存してください');
    return;
  }
  await pullReportsFromSheet(true);
}

function handleAdminLogin() {
  const id = elements.adminLoginId.value.trim();
  const password = elements.adminLoginPassword.value;
  const found = ADMIN_USERS.find((user) => user.id === id && user.password === password);

  if (!found) {
    elements.adminLoginError.textContent = '管理者IDまたはパスワードが正しくありません。';
    elements.adminLoginError.style.display = 'block';
    return;
  }

  state.adminUser = found;
  saveAdminSession(found);
  elements.adminLoginId.value = '';
  elements.adminLoginPassword.value = '';
  elements.adminLoginError.style.display = 'none';
  showToast('管理者ログインしました');
  renderAdminView();
}

function handleAdminLogout() {
  state.adminUser = null;
  saveAdminSession(null);
  state.achievementsSelectedStaff = '';
  state.achievementsSelectedPeriodKey = '';
  showToast('管理者ログアウトしました');
  if (state.mode === 'achievements' || state.mode === 'admin' || state.mode === 'admin-report') {
    openStaffListView();
    renderAdminView();
    return;
  }
  renderAdminView();
}

function openCreateView() {
  state.mode = 'create';
  state.returnView = 'staff-list';
  state.editingId = null;
  state.form = createEmptyForm();
  state.currentStep = 1;
  state.errors = {};
  state.photoPreview = null;
  switchView('form');
  renderFormView();
}

function openEditView(reportId, sourceView) {
  const target = state.reports.find((item) => item.id === reportId);
  if (!target) {
    showToast('対象の日報が見つかりませんでした');
    return;
  }
  state.mode = 'edit';
  state.returnView = sourceView;
  state.editingId = reportId;
  state.form = mergeForm(createEmptyForm(), target.payload);
  state.currentStep = 1;
  state.errors = {};
  state.photoPreview = null;
  switchView('form');
  renderFormView();
}

function backToReturnView() {
  if (state.returnView === 'admin') {
    openAdminView();
    return;
  }
  openStaffListView();
}

function switchView(key) {
  Object.values(views).forEach((view) => view.classList.remove('is-active'));
  views[key].classList.add('is-active');
}

function getReportById(reportId) {
  return state.reports.find((item) => item.id === reportId);
}

function openDetailView(reportId, fromView) {
  const report = getReportById(reportId);
  if (!report) {
    showToast('対象の日報が見つかりませんでした');
    return;
  }
  state.detailReportId = reportId;
  state.detailReturnView = fromView;
  switchView('detail');
  renderDetailView(report);
}

function backFromDetailView() {
  if (state.detailReturnView === 'admin-report') {
    openAdminReportView(state.adminFocusReportId);
    return;
  }
  if (state.detailReturnView === 'admin-confirmed') {
    openAdminViewKeepSelection();
    return;
  }
  if (state.detailReturnView === 'admin') {
    openAdminView();
    return;
  }
  setHeaderActiveRole('staff');
  switchView('staffList');
  renderStaffList();
}

function renderDetailView(report) {
  elements.detailContainer.innerHTML = buildDetailHtml(report);
}

function getStep1Photos(step1) {
  const list = Array.isArray(step1.photos) ? step1.photos.filter((item) => item && (item.url || item.dataUrl)) : [];
  if (list.length > 0) return list;

  const legacySrc = step1.photoUrl || step1.photoDataUrl || '';
  if (!legacySrc) return [];
  return [
    {
      name: step1.photoMeta && step1.photoMeta.name ? step1.photoMeta.name : '会場写真',
      type: step1.photoMeta && step1.photoMeta.type ? step1.photoMeta.type : 'image/jpeg',
      size: step1.photoMeta && step1.photoMeta.size ? step1.photoMeta.size : 0,
      url: step1.photoUrl || '',
      dataUrl: step1.photoDataUrl || '',
      fileId: step1.photoFileId || ''
    }
  ];
}

function syncLegacyPhotoFieldsFromPhotos(step1) {
  const photos = getStep1Photos(step1);
  if (photos.length === 0) {
    step1.photoMeta = null;
    step1.photoDataUrl = '';
    step1.photoUrl = '';
    step1.photoFileId = '';
    step1.photos = [];
    return;
  }

  const first = photos[0];
  step1.photoMeta = {
    name: first.name || '会場写真',
    size: Number(first.size || 0),
    type: first.type || 'image/jpeg'
  };
  step1.photoDataUrl = first.dataUrl || '';
  step1.photoUrl = first.url || '';
  step1.photoFileId = first.fileId || '';
  step1.photos = photos.slice(0, PHOTO_MAX_COUNT);
}

function buildDetailHtml(report) {
  const f = report.payload;
  const step2 = f.step2 || {};
  const seated = step2.seatedBreakdown || {};
  const step3 = f.step3 || {};
  const newA = step3.newAcquisitions || {};
  const ltv = step3.ltv || {};
  const auH = ltv.auHikariBreakdown || {};
  const blH = ltv.blHikariBreakdown || {};
  const cmH = ltv.commufaHikariBreakdown || {};
  const photos = getStep1Photos(f.step1);
  const hasPhoto = photos.length > 0;
  const photoHtml = photos
    .map((photo) => {
      if (!photo.url) return '';
      const link = photo.url
        ? `<p><a href="${escapeHtml(photo.url)}" target="_blank" rel="noopener noreferrer">写真を開く</a></p>`
        : '';
      return link;
    })
    .join('');

  const successCases = (Array.isArray(f.step4.cases) ? f.step4.cases : []).filter(hasFilledSuccessCase);
  const improveCases = (Array.isArray(f.step5.cases) ? f.step5.cases : []).filter(hasFilledImproveCase);
  const successHtml = successCases
    .map((item, index) => `
      <div class="panel">
        <p>事例${index + 1}</p>
        <p>来店理由: ${escapeHtml(item.visitReason || '-')}</p>
        <p>客層: ${escapeHtml(item.customerType || '-')}</p>
        <p>決め手トーク（タグ）: ${escapeHtml(item.talkTag || '-')}</p>
        <p>具体トーク: ${escapeHtml(item.talkDetail || '-')}</p>
        <p>成約要因: ${escapeHtml(item.contractFactor || '-')}</p>
        <p>その他: ${escapeHtml(item.other || '-')}</p>
      </div>
    `)
    .join('');
  const improveHtml = improveCases
    .map((item, index) => `
      <div class="panel">
        <p>改善${index + 1}</p>
        <p>改善ポイント: ${escapeHtml(item.improvePoint || '-')}</p>
        <p>理由: ${escapeHtml(item.reason || '-')}</p>
        <p>その他: ${escapeHtml(item.other || '-')}</p>
      </div>
    `)
    .join('');
  return `
    <h3>基本情報</h3>
    <p>日付: ${escapeHtml(f.step1.workDate || '-')}</p>
    <p>スタッフ: ${escapeHtml(f.step1.staffName || '-')}</p>
    <p>担当業務: ${escapeHtml(f.step1.jobRole || '-')}</p>
    <p>区分: ${escapeHtml(f.step1.workPlaceType || '-')}</p>
    <p>店舗名: ${escapeHtml(f.step1.storeName || '-')}</p>
    <p>イベント会場: ${escapeHtml(f.step1.eventVenue || '-')}</p>
    <p>イベント全体目標（店舗とイベント）: ${escapeHtml(f.step1.eventOverallTarget || '-')}</p>
    <p>イベント会場目標: ${escapeHtml(f.step1.eventVenueTarget || '-')}</p>
    <p>会場写真: ${hasPhoto ? `${photos.length}枚` : 'なし'}</p>
    ${photoHtml}

    <h3>アプローチ状況</h3>
    <p>来店数: ${step2.visitors ?? 0}</p>
    <p>キャッチ数: ${step2.catchCount ?? 0}</p>
    <p>着座数: ${step2.seated ?? 0}</p>
    <p>見込み数: ${step2.prospects ?? 0}</p>
    <p>着座内訳 au/UQ既存: ${seated.auUqExisting ?? 0}</p>
    <p>着座内訳 SB／ワイモバイル: ${seated.sbYmobile ?? 0}</p>
    <p>着座内訳 docomo／ahamo: ${seated.docomoAhamo ?? 0}</p>
    <p>着座内訳 楽天: ${seated.rakuten ?? 0}</p>
    <p>着座内訳 その他: ${seated.other ?? 0}</p>

    <h3>獲得実績（新規）</h3>
    <p>au MNP SIM単: ${newA.auMnpSim ?? 0}</p>
    <p>au MNP HS: ${newA.auMnpHs ?? 0}</p>
    <p>au純新規 SIM単: ${newA.auNewSim ?? 0}</p>
    <p>au純新規 HS: ${newA.auNewHs ?? 0}</p>
    <p>UQ MNP SIM単: ${newA.uqMnpSim ?? 0}</p>
    <p>UQ MNP HS: ${newA.uqMnpHs ?? 0}</p>
    <p>UQ純新規 SIM単: ${newA.uqNewSim ?? 0}</p>
    <p>UQ純新規 HS: ${newA.uqNewHs ?? 0}</p>
    <p>セルアップ: ${newA.cellUp ?? 0}</p>

    <h3>LTV</h3>
    <p>auでんき: ${ltv.auDenki ?? 0}</p>
    <p>ゴールドカード: ${ltv.goldCard ?? 0}</p>
    <p>シルバーカード: ${ltv.silverCard ?? 0}</p>
    <p>ランクアップ: ${ltv.rankUp ?? 0}</p>
    <p>じぶん銀行: ${ltv.jibunBank ?? 0}</p>
    <p>ノートン: ${ltv.norton ?? 0}</p>
    <p>auひかり 新規: ${auH.new ?? 0}</p>
    <p>auひかり ドコモ光から切替: ${auH.fromDocomo ?? 0}</p>
    <p>auひかり ソフトバンク光から切替: ${auH.fromSoftbank ?? 0}</p>
    <p>auひかり その他から切替: ${auH.fromOther ?? 0}</p>
    <p>BLひかり 新規: ${blH.new ?? 0}</p>
    <p>BLひかり ドコモ光から切替: ${blH.fromDocomo ?? 0}</p>
    <p>BLひかり ソフトバンク光から切替: ${blH.fromSoftbank ?? 0}</p>
    <p>BLひかり その他から切替: ${blH.fromOther ?? 0}</p>
    <p>コミュファ光 新規: ${cmH.new ?? 0}</p>
    <p>コミュファ光 ドコモ光から切替: ${cmH.fromDocomo ?? 0}</p>
    <p>コミュファ光 ソフトバンク光から切替: ${cmH.fromSoftbank ?? 0}</p>
    <p>コミュファ光 その他から切替: ${cmH.fromOther ?? 0}</p>

    <h3>成約事例</h3>
    ${successHtml || '<p>-</p>'}

    <h3>改善事例</h3>
    ${improveHtml || '<p>-</p>'}

    <h3>イベント会場の評価</h3>
    <p>評価: ${escapeHtml((f.step5_5 && f.step5_5.venueEvaluation) || '-')}</p>
    <p>その他: ${escapeHtml((f.step5_5 && f.step5_5.other) || '-')}</p>

    <h3>振り返り</h3>
    <p>所感: ${escapeHtml(f.step6.impression || '-')}</p>
    <p>備考: ${escapeHtml(f.step6.notes || '-')}</p>
    <p>管理者総括: ${escapeHtml(f.step6.adminSummary || '-')}</p>
    <p>確認状況: ${report.confirmed ? `確認済み（${escapeHtml(report.confirmedBy || '確認者不明')}）` : '未確認'}</p>
  `;
}

function openAdminReportView(reportId, returnView = 'admin') {
  const report = getReportById(reportId);
  if (!report) {
    showToast('対象の日報が見つかりませんでした');
    return;
  }
  state.adminFocusReportId = reportId;
  state.adminReportReturnView = returnView;
  switchView('adminReport');
  elements.adminReportStaff.textContent = `スタッフ: ${report.payload.step1.staffName || '-'}`;
  const confirmedInfo = report.confirmed ? ` / 確認者: ${report.confirmedBy || '不明'}` : '';
  elements.adminReportDate.textContent = `稼働日: ${report.payload.step1.workDate || '-'} / 会場: ${report.payload.step1.eventVenue || '-'}${confirmedInfo}`;
  elements.adminReportConfirmButton.disabled = false;
  elements.adminReportConfirmButton.textContent = report.confirmed ? '未確認に戻す' : '確認済みにする';
  elements.adminReportSummaryInput.value = report.payload.step6.adminSummary || '';
  elements.adminReportSummaryEditor.style.display = 'none';
}

function handleAdminReportConfirm() {
  if (!state.adminFocusReportId) return;
  const index = state.reports.findIndex((item) => item.id === state.adminFocusReportId);
  if (index < 0) return;
  const previous = deepCopy(state.reports[index]);
  state.reports[index].confirmed = !state.reports[index].confirmed;
  if (state.reports[index].confirmed) {
    state.reports[index].confirmedBy = state.adminUser ? state.adminUser.name : '管理者';
    state.reports[index].confirmedAt = new Date().toISOString();
  } else {
    state.reports[index].confirmedBy = '';
    state.reports[index].confirmedAt = '';
  }
  state.reports[index].updatedAt = new Date().toISOString();
  if (!saveReports()) {
    state.reports[index] = previous;
    return;
  }
  syncUpsert(state.reports[index]);
  showToast(state.reports[index].confirmed ? '確認済みにしました' : '未確認に戻しました');
  openAdminReportView(state.adminFocusReportId, state.adminReportReturnView || 'admin');
}

function handleAdminReportPreview() {
  if (!state.adminFocusReportId) return;
  openDetailView(state.adminFocusReportId, 'admin-report');
}

function handleAdminReportEdit() {
  if (!state.adminFocusReportId) return;
  openEditView(state.adminFocusReportId, 'admin');
}

function toggleAdminReportSummaryEditor() {
  if (!state.adminFocusReportId) return;
  const report = getReportById(state.adminFocusReportId);
  if (!report) return;
  elements.adminReportSummaryInput.value = report.payload.step6.adminSummary || '';
  const current = elements.adminReportSummaryEditor.style.display;
  elements.adminReportSummaryEditor.style.display = current === 'none' ? 'block' : 'none';
}

function handleAdminReportSummarySave() {
  if (!state.adminFocusReportId) return;
  const summary = elements.adminReportSummaryInput.value.trim();
  const ok = updateAdminSummary(state.adminFocusReportId, summary);
  if (!ok) return;
  elements.adminReportSummaryEditor.style.display = 'none';
  openAdminReportView(state.adminFocusReportId, state.adminReportReturnView || 'admin');
}

function handleAdminReportBack() {
  if (state.adminReportReturnView === 'admin-confirmed') {
    openAdminViewKeepSelection();
    return;
  }
  openAdminView();
}

function getSortedReports(reports) {
  return [...reports].sort((a, b) => {
    const dateDiff = b.payload.step1.workDate.localeCompare(a.payload.step1.workDate);
    if (dateDiff !== 0) return dateDiff;
    return b.updatedAt.localeCompare(a.updatedAt);
  });
}

function groupReportsByStaff(reports) {
  const map = new Map();

  reports.forEach((report) => {
    const name = resolveStaffGroupName(report.payload.step1.staffName);
    if (!map.has(name)) map.set(name, []);
    map.get(name).push(report);
  });

  const groups = [...map.entries()].map(([staffName, items]) => ({
    staffName,
    items: getSortedReports(items)
  }));

  groups.sort((a, b) => {
    const aDate = a.items[0] ? a.items[0].payload.step1.workDate : '';
    const bDate = b.items[0] ? b.items[0].payload.step1.workDate : '';
    return bDate.localeCompare(aDate);
  });

  return groups;
}

function resolveStaffGroupName(rawName) {
  const normalized = normalizeStaffName(rawName);
  if (normalized) return normalized;
  return '未設定スタッフ';
}

function getStaffFilteredReports() {
  return getSortedReports(state.reports);
}

function getConfirmText(report) {
  if (!report.confirmed) return '未確認';
  return report.confirmedBy ? `確認済み（${report.confirmedBy}）` : '確認済み';
}

function createSimpleReportCard(report, sourceView) {
  const fragment = elements.cardTemplate.content.cloneNode(true);
  const map = {
    primaryLabel: '日付',
    primaryValue: report.payload.step1.workDate || '-',
    secondaryValue: `会場: ${report.payload.step1.eventVenue || '-'} / ${getConfirmText(report)}`
  };

  Object.entries(map).forEach(([key, value]) => {
    const node = fragment.querySelector(`[data-field="${key}"]`);
    if (node) node.textContent = String(value);
  });

  const openButton = fragment.querySelector('[data-action="open"]');
  openButton.dataset.kind = 'report';
  openButton.dataset.id = report.id;
  openButton.dataset.source = sourceView;
  openButton.textContent = '編集';

  const previewButton = fragment.querySelector('[data-action="preview"]');
  previewButton.dataset.kind = 'preview';
  previewButton.dataset.id = report.id;
  previewButton.dataset.source = sourceView;
  previewButton.textContent = '内容確認';

  return fragment;
}

function createStaffNameCard(staffName, count, latestDate) {
  const fragment = elements.cardTemplate.content.cloneNode(true);
  const map = {
    primaryLabel: 'スタッフ名',
    primaryValue: staffName,
    secondaryValue: `件数: ${count}件 / 最新日付: ${latestDate || '-'}`
  };

  Object.entries(map).forEach(([key, value]) => {
    const node = fragment.querySelector(`[data-field="${key}"]`);
    if (node) node.textContent = String(value);
  });

  const openButton = fragment.querySelector('[data-action="open"]');
  openButton.dataset.kind = 'staff';
  openButton.dataset.staff = staffName;
  openButton.textContent = '日報を見る';

  const previewButton = fragment.querySelector('[data-action="preview"]');
  previewButton.style.display = 'none';

  return fragment;
}

function createAdminUncheckedCard(report) {
  const fragment = elements.cardTemplate.content.cloneNode(true);
  const map = {
    primaryLabel: 'スタッフ名',
    primaryValue: report.payload.step1.staffName || '-',
    secondaryValue: `稼働日: ${report.payload.step1.workDate || '-'}`
  };

  Object.entries(map).forEach(([key, value]) => {
    const node = fragment.querySelector(`[data-field="${key}"]`);
    if (node) node.textContent = String(value);
  });

  const openButton = fragment.querySelector('[data-action="open"]');
  openButton.dataset.kind = 'admin-unchecked';
  openButton.dataset.id = report.id;
  openButton.textContent = '日報を見る';

  const previewButton = fragment.querySelector('[data-action="preview"]');
  previewButton.style.display = 'none';

  return fragment;
}

function createLoadMoreButton(action, remaining) {
  const wrap = document.createElement('div');
  wrap.className = 'card-actions';
  const button = document.createElement('button');
  button.type = 'button';
  button.className = 'btn btn-ghost';
  button.dataset.action = action;
  button.textContent = `もっと見る（残り${remaining}件）`;
  wrap.appendChild(button);
  return wrap;
}

function renderStaffList() {
  const filtered = getStaffFilteredReports();
  elements.reportList.innerHTML = '';
  const groups = groupReportsByStaff(filtered);

  if (!state.selectedStaffName) {
    elements.staffListTitle.textContent = '日報一覧（スタッフ）';
    elements.staffBackButton.style.display = 'none';
    elements.listEmpty.style.display = groups.length === 0 ? 'block' : 'none';

    const list = document.createElement('div');
    list.className = 'card-grid';
    const visibleGroups = groups.slice(0, state.viewLimits.staffGroups);
    visibleGroups.forEach((group) => {
      const latestDate = group.items[0] ? group.items[0].payload.step1.workDate : '';
      list.appendChild(createStaffNameCard(group.staffName, group.items.length, latestDate));
    });
    elements.reportList.appendChild(list);
    if (groups.length > visibleGroups.length) {
      elements.reportList.appendChild(createLoadMoreButton('load-more-staff-groups', groups.length - visibleGroups.length));
    }
    return;
  }

  const targetGroup = groups.find((group) => group.staffName === state.selectedStaffName);
  const reports = targetGroup ? targetGroup.items : [];
  elements.staffListTitle.textContent = `日報一覧: ${state.selectedStaffName}`;
  elements.staffBackButton.style.display = 'inline-flex';
  elements.listEmpty.style.display = reports.length === 0 ? 'block' : 'none';

  const list = document.createElement('div');
  list.className = 'card-grid';
  const visibleReports = reports.slice(0, state.viewLimits.staffReports);
  visibleReports.forEach((report) => {
    list.appendChild(createSimpleReportCard(report, 'staff-list'));
  });
  elements.reportList.appendChild(list);
  if (reports.length > visibleReports.length) {
    elements.reportList.appendChild(createLoadMoreButton('load-more-staff-reports', reports.length - visibleReports.length));
  }
}

function getAdminFilteredReports() {
  return getSortedReports(state.reports);
}

function buildAdminSnippet(report) {
  const firstGood = (Array.isArray(report.payload.step4.cases) ? report.payload.step4.cases : []).find(hasFilledSuccessCase) || null;
  const firstBad = (Array.isArray(report.payload.step5.cases) ? report.payload.step5.cases : []).find(hasFilledImproveCase) || null;
  const good = firstGood ? `成約: ${firstGood.visitReason || firstGood.contractFactor || '-'}` : '';
  const bad = firstBad ? `改善: ${firstBad.improvePoint || firstBad.reason || '-'}` : '';
  const memo = report.payload.step6.impression ? `所感: ${report.payload.step6.impression}` : '';
  const summary = report.payload.step6.adminSummary ? `総括: ${report.payload.step6.adminSummary}` : '';
  return [good, bad, memo, summary].filter(Boolean).join(' / ') || '事例未入力';
}

function createAdminCard(report) {
  const fragment = elements.adminCardTemplate.content.cloneNode(true);
  const hasGood = (Array.isArray(report.payload.step4.cases) ? report.payload.step4.cases : []).some(hasFilledSuccessCase);
  const hasBad = (Array.isArray(report.payload.step5.cases) ? report.payload.step5.cases : []).some(hasFilledImproveCase);
  const map = {
    staffName: report.payload.step1.staffName || '-',
    workDate: report.payload.step1.workDate || '-',
    goodTag: hasGood ? '成約事例あり' : '成約事例なし',
    badTag: hasBad ? '改善事例あり' : '改善事例なし',
    confirmTag: report.confirmed ? `確認済み: ${report.confirmedBy || '不明'}` : '未確認',
    snippet: buildAdminSnippet(report)
  };

  Object.entries(map).forEach(([key, value]) => {
    const node = fragment.querySelector(`[data-field="${key}"]`);
    if (node) node.textContent = String(value);
  });

  const confirmButton = fragment.querySelector('[data-action="toggle-confirm"]');
  confirmButton.dataset.id = report.id;
  confirmButton.textContent = report.confirmed ? '未確認に戻す' : '確認済みにする';

  const editButton = fragment.querySelector('[data-action="edit"]');
  editButton.dataset.id = report.id;
  editButton.dataset.source = 'admin';

  return fragment;
}

function renderAdminLists() {
  const filtered = getAdminFilteredReports();
  const unchecked = filtered.filter((report) => !report.confirmed);
  const confirmed = filtered.filter((report) => report.confirmed);

  elements.adminUncheckedList.innerHTML = '';
  elements.adminUncheckedEmpty.style.display = unchecked.length === 0 ? 'block' : 'none';

  const visibleUnchecked = unchecked.slice(0, state.viewLimits.adminUnchecked);
  visibleUnchecked.forEach((report) => {
    elements.adminUncheckedList.appendChild(createAdminUncheckedCard(report));
  });
  if (unchecked.length > visibleUnchecked.length) {
    elements.adminUncheckedList.appendChild(createLoadMoreButton('load-more-admin-unchecked', unchecked.length - visibleUnchecked.length));
  }

  const confirmedGroups = groupReportsByStaff(confirmed);
  renderAdminConfirmedSection(confirmedGroups);
}

function renderAdminConfirmedSection(confirmedGroups) {
  elements.adminConfirmedStaffList.innerHTML = '';
  elements.adminConfirmedDateList.innerHTML = '';
  elements.adminConfirmedDetailContent.innerHTML = '';
  elements.adminConfirmedSummaryEditor.style.display = 'none';

  const hasConfirmed = confirmedGroups.length > 0;
  elements.adminConfirmedStaffEmpty.style.display = hasConfirmed ? 'none' : 'block';

  if (!state.adminConfirmedSelectedStaff) {
    elements.adminConfirmedDateWrap.style.display = 'none';
    elements.adminConfirmedDetailWrap.style.display = 'none';
    const visibleGroups = confirmedGroups.slice(0, state.viewLimits.adminConfirmedStaff);
    visibleGroups.forEach((group) => {
      const card = createStaffNameCard(
        group.staffName,
        group.items.length,
        group.items[0] ? group.items[0].payload.step1.workDate : ''
      );
      const openButton = card.querySelector('[data-action="open"]');
      openButton.dataset.kind = 'confirmed-staff';
      openButton.dataset.staff = group.staffName;
      openButton.textContent = '日付を見る';
      elements.adminConfirmedStaffList.appendChild(card);
    });
    if (confirmedGroups.length > visibleGroups.length) {
      elements.adminConfirmedStaffList.appendChild(createLoadMoreButton('load-more-admin-confirmed-staff', confirmedGroups.length - visibleGroups.length));
    }
    return;
  }

  const selectedGroup = confirmedGroups.find((group) => group.staffName === state.adminConfirmedSelectedStaff);
  if (!selectedGroup) {
    state.adminConfirmedSelectedStaff = '';
    state.adminConfirmedSelectedReportId = '';
    renderAdminConfirmedSection(confirmedGroups);
    return;
  }

  elements.adminConfirmedDateWrap.style.display = 'block';
  elements.adminConfirmedStaffTitle.textContent = `${selectedGroup.staffName} の日報`;
  const visibleDates = selectedGroup.items.slice(0, state.viewLimits.adminConfirmedDates);
  visibleDates.forEach((report) => {
    const card = createSimpleReportCard(report, 'admin');
    const openButton = card.querySelector('[data-action="open"]');
    openButton.dataset.kind = 'confirmed-date';
    openButton.dataset.id = report.id;
    openButton.textContent = '内容確認';
    const previewButton = card.querySelector('[data-action="preview"]');
    previewButton.style.display = 'none';
    elements.adminConfirmedDateList.appendChild(card);
  });
  if (selectedGroup.items.length > visibleDates.length) {
    elements.adminConfirmedDateList.appendChild(createLoadMoreButton('load-more-admin-confirmed-dates', selectedGroup.items.length - visibleDates.length));
  }

  elements.adminConfirmedDetailWrap.style.display = 'none';
  elements.adminConfirmedSummarySaveButton.dataset.id = '';
}

function onStaffListClick(event) {
  const actionButton = event.target.closest('[data-action]');
  if (actionButton && actionButton.dataset.action === 'load-more-staff-groups') {
    state.viewLimits.staffGroups += LIST_PAGE_SIZE;
    renderStaffList();
    return;
  }
  if (actionButton && actionButton.dataset.action === 'load-more-staff-reports') {
    state.viewLimits.staffReports += LIST_PAGE_SIZE;
    renderStaffList();
    return;
  }

  const previewTarget = event.target.closest('[data-action="preview"]');
  if (previewTarget && previewTarget.dataset.kind === 'preview') {
    openDetailView(previewTarget.dataset.id, previewTarget.dataset.source || 'staff-list');
    return;
  }

  const target = event.target.closest('[data-action="open"]');
  if (!target) return;
  if (target.dataset.kind === 'staff') {
    state.selectedStaffName = target.dataset.staff || '';
    state.viewLimits.staffReports = LIST_PAGE_SIZE;
    renderStaffList();
    return;
  }
  if (target.dataset.kind === 'report') {
    openEditView(target.dataset.id, target.dataset.source || 'staff-list');
  }
}

function onAdminListClick(event) {
  const actionButton = event.target.closest('[data-action]');
  if (actionButton && actionButton.dataset.action === 'load-more-admin-unchecked') {
    state.viewLimits.adminUnchecked += LIST_PAGE_SIZE;
    renderAdminLists();
    return;
  }

  const previewButton = event.target.closest('[data-action="preview"]');
  if (previewButton && previewButton.dataset.kind === 'preview') {
    openDetailView(previewButton.dataset.id, previewButton.dataset.source || 'admin');
    return;
  }

  const openButton = event.target.closest('[data-action="open"]');
  if (openButton && openButton.dataset.kind === 'report') {
    openEditView(openButton.dataset.id, openButton.dataset.source || 'admin');
    return;
  }
  if (openButton && openButton.dataset.kind === 'admin-unchecked') {
    openAdminReportView(openButton.dataset.id, 'admin');
    return;
  }

  const confirmButton = event.target.closest('[data-action="toggle-confirm"]');
  if (confirmButton) {
    updateConfirmedStatus(confirmButton.dataset.id);
    return;
  }

  const editButton = event.target.closest('[data-action="edit"]');
  if (editButton) {
    openEditView(editButton.dataset.id, editButton.dataset.source || 'admin');
  }
}

function onAdminConfirmedClick(event) {
  const actionButton = event.target.closest('[data-action]');
  if (actionButton && actionButton.dataset.action === 'load-more-admin-confirmed-staff') {
    state.viewLimits.adminConfirmedStaff += LIST_PAGE_SIZE;
    renderAdminLists();
    return;
  }
  if (actionButton && actionButton.dataset.action === 'load-more-admin-confirmed-dates') {
    state.viewLimits.adminConfirmedDates += LIST_PAGE_SIZE;
    renderAdminLists();
    return;
  }

  const openButton = event.target.closest('[data-action="open"]');
  if (!openButton) return;

  if (openButton.dataset.kind === 'confirmed-staff') {
    state.adminConfirmedSelectedStaff = openButton.dataset.staff || '';
    state.adminConfirmedSelectedReportId = '';
    state.viewLimits.adminConfirmedDates = LIST_PAGE_SIZE;
    renderAdminLists();
    return;
  }

  if (openButton.dataset.kind === 'confirmed-date') {
    const reportId = openButton.dataset.id || '';
    if (!reportId) return;
    openAdminReportView(reportId, 'admin-confirmed');
  }
}

function handleAdminConfirmedBack() {
  state.adminConfirmedSelectedStaff = '';
  state.adminConfirmedSelectedReportId = '';
  state.viewLimits.adminConfirmedDates = LIST_PAGE_SIZE;
  renderAdminLists();
}

function handleAdminConfirmedDetailEdit() {
  const reportId = elements.adminConfirmedDetailEditButton.dataset.id;
  if (!reportId) return;
  openEditView(reportId, 'admin');
}

function toggleAdminConfirmedSummaryEditor() {
  const reportId = elements.adminConfirmedSummarySaveButton.dataset.id;
  if (!reportId) return;
  const report = getReportById(reportId);
  if (!report) return;
  elements.adminConfirmedSummaryInput.value = report.payload.step6.adminSummary || '';
  const current = elements.adminConfirmedSummaryEditor.style.display;
  elements.adminConfirmedSummaryEditor.style.display = current === 'none' ? 'block' : 'none';
}

function handleAdminConfirmedSummarySave() {
  const reportId = elements.adminConfirmedSummarySaveButton.dataset.id;
  if (!reportId) return;
  const summary = elements.adminConfirmedSummaryInput.value.trim();
  const ok = updateAdminSummary(reportId, summary);
  if (!ok) return;
  elements.adminConfirmedSummaryEditor.style.display = 'none';
  renderAdminLists();
}

function updateConfirmedStatus(reportId) {
  const index = state.reports.findIndex((item) => item.id === reportId);
  if (index < 0) {
    showToast('対象の日報が見つかりませんでした');
    return;
  }
  const previous = deepCopy(state.reports[index]);
  state.reports[index].confirmed = !state.reports[index].confirmed;
  if (state.reports[index].confirmed) {
    state.reports[index].confirmedBy = state.adminUser ? state.adminUser.name : '管理者';
    state.reports[index].confirmedAt = new Date().toISOString();
  } else {
    state.reports[index].confirmedBy = '';
    state.reports[index].confirmedAt = '';
  }
  state.reports[index].updatedAt = new Date().toISOString();
  if (!saveReports()) {
    state.reports[index] = previous;
    return;
  }
  syncUpsert(state.reports[index]);
  showToast(state.reports[index].confirmed ? '確認済みにしました' : '未確認に戻しました');
  renderAdminView();
}

function updateAdminSummary(reportId, summaryText) {
  const index = state.reports.findIndex((item) => item.id === reportId);
  if (index < 0) {
    showToast('対象の日報が見つかりませんでした');
    return false;
  }

  const previous = deepCopy(state.reports[index]);
  state.reports[index].payload.step6.adminSummary = summaryText;
  state.reports[index].updatedAt = new Date().toISOString();

  if (!saveReports()) {
    state.reports[index] = previous;
    return false;
  }

  syncUpsert(state.reports[index]);
  showToast('管理者コメントを保存しました');
  return true;
}

function renderFormView() {
  elements.formTitle.textContent = state.mode === 'create' ? '日報作成' : '日報編集';
  elements.stepText.textContent = `STEP ${state.currentStep}/7`;
  elements.stepProgress.setAttribute('aria-valuemax', '7');
  elements.stepProgress.setAttribute('aria-valuenow', String(state.currentStep));
  elements.stepProgressBar.style.width = `${(state.currentStep / 7) * 100}%`;
  elements.prevStepButton.disabled = state.currentStep === 1;
  elements.nextStepButton.textContent = state.currentStep < 7 ? '次へ' : state.mode === 'create' ? '保存する' : '更新する';

  elements.stepContainer.innerHTML = renderStepHtml(state.currentStep);
  bindStepInputs(state.currentStep);

  if (state.mode === 'edit') {
    renderDeleteButton();
  }
}

function renderDeleteButton() {
  const old = elements.stepContainer.querySelector('.inline-actions');
  if (old) old.remove();

  const wrap = document.createElement('div');
  wrap.className = 'inline-actions';
  const button = document.createElement('button');
  button.type = 'button';
  button.className = 'btn btn-danger';
  button.textContent = 'この日報を削除';
  button.addEventListener('click', onDeleteCurrent);
  wrap.appendChild(button);
  elements.stepContainer.appendChild(wrap);
}

function goPrevStep() {
  if (state.currentStep <= 1) return;
  state.currentStep -= 1;
  state.errors = {};
  renderFormView();
}

function goNextStepOrSubmit() {
  if (state.photoUploading) {
    showToast('写真をアップロード中です。完了後に進んでください');
    return;
  }

  const stepErrors = validateStep(state.currentStep, state.form);
  if (Object.keys(stepErrors).length > 0) {
    state.errors = stepErrors;
    renderFormView();
    return;
  }

  if (state.currentStep < 7) {
    state.currentStep += 1;
    state.errors = {};
    renderFormView();
    return;
  }

  const firstInvalidStep = findFirstInvalidStep();
  if (firstInvalidStep > 0) {
    state.currentStep = firstInvalidStep;
    state.errors = validateStep(firstInvalidStep, state.form);
    renderFormView();
    return;
  }

  if (state.mode === 'create') {
    void createReport();
  } else {
    void updateReport();
  }
}

function findFirstInvalidStep() {
  for (let i = 1; i <= 7; i += 1) {
    const errors = validateStep(i, state.form);
    if (Object.keys(errors).length > 0) return i;
  }
  return 0;
}

async function createReport() {
  const prepared = await preparePhotoForPersist(state.form.step1);
  if (!prepared.ok) {
    showToast(prepared.message);
    return;
  }
  state.form.step1.photoDataUrl = prepared.photoDataUrl;
  state.form.step1.photoUrl = prepared.photoUrl;
  state.form.step1.photoFileId = prepared.photoFileId;
  state.form.step1.photos = prepared.photos;

  const now = new Date().toISOString();
  const report = {
    id: makeId(),
    createdAt: now,
    updatedAt: now,
    folder: '未分類',
    confirmed: false,
    confirmedBy: '',
    confirmedAt: '',
    payload: deepCopy(state.form)
  };
  state.reports.push(report);
  if (!saveReports()) {
    state.reports.pop();
    return;
  }
  syncUpsert(report);
  showToast('日報を保存しました');
  openEditView(report.id, state.returnView);
}

async function updateReport() {
  const index = state.reports.findIndex((item) => item.id === state.editingId);
  if (index < 0) {
    showToast('更新対象が見つかりませんでした');
    return;
  }

  const prepared = await preparePhotoForPersist(state.form.step1);
  if (!prepared.ok) {
    showToast(prepared.message);
    return;
  }
  state.form.step1.photoDataUrl = prepared.photoDataUrl;
  state.form.step1.photoUrl = prepared.photoUrl;
  state.form.step1.photoFileId = prepared.photoFileId;
  state.form.step1.photos = prepared.photos;

  const previous = deepCopy(state.reports[index]);
  state.reports[index] = {
    ...state.reports[index],
    updatedAt: new Date().toISOString(),
    payload: deepCopy(state.form)
  };
  if (!saveReports()) {
    state.reports[index] = previous;
    return;
  }
  syncUpsert(state.reports[index]);
  showToast('日報を更新しました');

  if (state.returnView === 'admin') {
    openAdminView();
  } else {
    openStaffListView();
  }
}

function onDeleteCurrent() {
  const ok = window.confirm('この日報を削除します。よろしいですか？');
  if (!ok) return;

  const deletedId = state.editingId;
  const previousReports = deepCopy(state.reports);
  state.reports = state.reports.filter((item) => item.id !== state.editingId);
  if (!saveReports()) {
    state.reports = previousReports;
    return;
  }
  syncDelete(deletedId);
  showToast('日報を削除しました');

  if (state.returnView === 'admin') {
    openAdminView();
  } else {
    openStaffListView();
  }
}

function showToast(message) {
  if (state.toastTimer) clearTimeout(state.toastTimer);
  elements.toast.textContent = message;
  elements.toast.classList.add('is-visible');
  state.toastTimer = setTimeout(() => {
    elements.toast.classList.remove('is-visible');
  }, 1800);
}

function toInt(value) {
  const n = Number(value);
  if (!Number.isFinite(n) || n < 0) return 0;
  return Math.floor(n);
}

function escapeHtml(value) {
  return String(value)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function errorText(field) {
  return state.errors[field] ? `<p class="error">${escapeHtml(state.errors[field])}</p>` : '';
}

function numberInput(label, path, value) {
  return `
    <div class="field-group">
      <label class="field-label" for="${path}">${label}</label>
      <input id="${path}" type="number" min="0" step="1" data-path="${path}" value="${escapeHtml(value)}" />
      ${errorText(path)}
    </div>
  `;
}

function textInput(label, path, value, required = false, type = 'text', hint = '') {
  return `
    <div class="field-group">
      <label class="field-label" for="${path}">${label}${required ? '<span class="required">必須</span>' : ''}</label>
      <input id="${path}" type="${type}" data-path="${path}" value="${escapeHtml(value || '')}" />
      ${hint ? `<p class="hint">${escapeHtml(hint)}</p>` : ''}
      ${errorText(path)}
    </div>
  `;
}

function textareaInput(label, path, value, required = false) {
  return `
    <div class="field-group">
      <label class="field-label" for="${path}">${label}${required ? '<span class="required">必須</span>' : ''}</label>
      <textarea id="${path}" data-path="${path}">${escapeHtml(value || '')}</textarea>
      ${errorText(path)}
    </div>
  `;
}

function selectInput(label, path, value, list, required = false) {
  const opts = list
    .map((option) => {
      const selected = option === value ? 'selected' : '';
      const text = option || '選択してください';
      return `<option value="${escapeHtml(option)}" ${selected}>${escapeHtml(text)}</option>`;
    })
    .join('');

  return `
    <div class="field-group">
      <label class="field-label" for="${path}">${label}${required ? '<span class="required">必須</span>' : ''}</label>
      <select id="${path}" data-path="${path}">${opts}</select>
      ${errorText(path)}
    </div>
  `;
}

function staffNameHybridInput(value, list, required = false) {
  const listId = 'step1-staffName-list';
  const selectedInList = list.includes(value) ? value : '';
  const selectOptions = ['']
    .concat(list)
    .map((option) => {
      const selected = option === selectedInList ? 'selected' : '';
      const text = option || 'プルダウンから選択';
      return `<option value="${escapeHtml(option)}" ${selected}>${escapeHtml(text)}</option>`;
    })
    .join('');
  const opts = list
    .map((option) => `<option value="${escapeHtml(option)}"></option>`)
    .join('');

  return `
    <div class="field-group">
      <label class="field-label" for="step1.staffName">${'スタッフ名'}${required ? '<span class="required">必須</span>' : ''}</label>
      <select id="step1.staffNameSelect" data-path="step1.staffNameSelect">${selectOptions}</select>
      <input id="step1.staffName" type="text" data-path="step1.staffName" list="${listId}" value="${escapeHtml(value || '')}" />
      <datalist id="${listId}">${opts}</datalist>
      <p class="hint">初回は下の入力欄に入力、2回目以降は上のプルダウンでも選べます</p>
      ${errorText('step1.staffName')}
    </div>
  `;
}

function getStaffNameOptions() {
  const set = new Set(STAFF_NAME_OPTIONS.filter(Boolean));
  state.reports.forEach((report) => {
    const name = normalizeStaffName(report && report.payload && report.payload.step1 ? report.payload.step1.staffName : '');
    if (name) set.add(name);
  });
  const current = normalizeStaffName(state.form && state.form.step1 ? state.form.step1.staffName : '');
  if (current) set.add(current);
  return Array.from(set);
}

function ltvBreakdownGroup(title, basePath, values) {
  return `
    <details class="collapse">
      <summary>${title} 内訳</summary>
      <div class="grid-2">
        ${numberInput(`${title} 新規`, `${basePath}.new`, values.new)}
        ${numberInput(`${title} ドコモ光から切替`, `${basePath}.fromDocomo`, values.fromDocomo)}
        ${numberInput(`${title} ソフトバンク光から切替`, `${basePath}.fromSoftbank`, values.fromSoftbank)}
        ${numberInput(`${title} その他から切替`, `${basePath}.fromOther`, values.fromOther)}
      </div>
    </details>
  `;
}

function isAdminEditor() {
  return state.mode === 'edit' && state.returnView === 'admin' && Boolean(state.adminUser);
}

function renderStepHtml(step) {
  const f = state.form;

  if (step === 1) {
    const photos = getStep1Photos(f.step1);
    const photoName = photos.length > 0 ? `選択済み: ${photos.length}枚` : '写真は未選択です';
    const isStoreSv = f.step1.workPlaceType === '店頭SV';
    const photoListHtml = photos
      .map((photo, index) => {
        const num = index + 1;
        const status = photo.url ? '（Drive保存済み）' : '（端末保持）';
        const name = photo.name || `写真${num}`;
        return `<p class="hint">写真${num}: ${escapeHtml(name)} ${status}</p>`;
      })
      .join('');
    return `
      <h3>STEP1: 基本情報</h3>
      <p class="hint">最初は軽い入力から始めます。</p>
      ${textInput('稼働日', 'step1.workDate', f.step1.workDate, true, 'date')}
      ${staffNameHybridInput(f.step1.staffName, getStaffNameOptions(), true)}
      ${selectInput('担当業務', 'step1.jobRole', f.step1.jobRole, options.jobRoles)}
      ${selectInput('区分（店頭SV / イベント）', 'step1.workPlaceType', f.step1.workPlaceType, options.workPlaceTypes, true)}
      ${textInput('店舗名', 'step1.storeName', f.step1.storeName, true)}
      ${isStoreSv ? '' : textInput('イベント会場', 'step1.eventVenue', f.step1.eventVenue, true)}
      ${isStoreSv ? '' : textInput('イベント全体目標（店舗とイベント）', 'step1.eventOverallTarget', f.step1.eventOverallTarget)}
      ${isStoreSv ? '' : textInput('イベント会場目標', 'step1.eventVenueTarget', f.step1.eventVenueTarget)}
      <div class="field-group">
        <label class="field-label" for="step1.photo">会場写真（任意・最大${PHOTO_MAX_COUNT}枚）</label>
        <input id="step1.photo" type="file" accept="image/*" multiple />
        ${state.photoUploading ? '<p class="hint">写真をGoogle Driveへアップロード中です...</p>' : ''}
        <p class="hint" id="photo-meta">${escapeHtml(photoName)}</p>
        ${photoListHtml || '<p class="hint">写真なし</p>'}
      </div>
    `;
  }

  if (step === 2) {
    return `
      <h3>STEP2: アプローチ状況</h3>
      <p class="hint">ここは数値入力だけです。</p>
      <div class="grid-2">
        ${numberInput('来店数', 'step2.visitors', f.step2.visitors)}
        ${numberInput('キャッチ数（反応数）', 'step2.catchCount', f.step2.catchCount)}
        ${numberInput('着座数', 'step2.seated', f.step2.seated)}
        ${numberInput('見込み', 'step2.prospects', f.step2.prospects)}
      </div>
      <h4>着座内訳</h4>
      <div class="grid-2">
        ${numberInput('au/UQ既存', 'step2.seatedBreakdown.auUqExisting', f.step2.seatedBreakdown.auUqExisting)}
        ${numberInput('SB／ワイモバイル', 'step2.seatedBreakdown.sbYmobile', f.step2.seatedBreakdown.sbYmobile)}
        ${numberInput('docomo／ahamo', 'step2.seatedBreakdown.docomoAhamo', f.step2.seatedBreakdown.docomoAhamo)}
        ${numberInput('楽天', 'step2.seatedBreakdown.rakuten', f.step2.seatedBreakdown.rakuten)}
        ${numberInput('その他', 'step2.seatedBreakdown.other', f.step2.seatedBreakdown.other)}
      </div>
    `;
  }

  if (step === 3) {
    return `
      <h3>STEP3: 獲得実績</h3>
      <p class="hint">カテゴリごとに開いて入力します。</p>

      <details class="collapse">
        <summary>新規獲得</summary>
        <div class="grid-2">
          ${numberInput('au MNP SIM単', 'step3.newAcquisitions.auMnpSim', f.step3.newAcquisitions.auMnpSim)}
          ${numberInput('au MNP HS', 'step3.newAcquisitions.auMnpHs', f.step3.newAcquisitions.auMnpHs)}
          ${numberInput('au純新規 SIM単', 'step3.newAcquisitions.auNewSim', f.step3.newAcquisitions.auNewSim)}
          ${numberInput('au純新規 HS', 'step3.newAcquisitions.auNewHs', f.step3.newAcquisitions.auNewHs)}
          ${numberInput('UQ MNP SIM単', 'step3.newAcquisitions.uqMnpSim', f.step3.newAcquisitions.uqMnpSim)}
        ${numberInput('UQ MNP HS', 'step3.newAcquisitions.uqMnpHs', f.step3.newAcquisitions.uqMnpHs)}
        ${numberInput('UQ純新規 SIM単', 'step3.newAcquisitions.uqNewSim', f.step3.newAcquisitions.uqNewSim)}
        ${numberInput('UQ純新規 HS', 'step3.newAcquisitions.uqNewHs', f.step3.newAcquisitions.uqNewHs)}
        ${numberInput('セルアップ', 'step3.newAcquisitions.cellUp', f.step3.newAcquisitions.cellUp)}
      </div>
      </details>

      <details class="collapse">
        <summary>LTV</summary>
        <div class="grid-2">
          ${numberInput('auでんき', 'step3.ltv.auDenki', f.step3.ltv.auDenki)}
          ${numberInput('ゴールドカード', 'step3.ltv.goldCard', f.step3.ltv.goldCard)}
          ${numberInput('シルバーカード', 'step3.ltv.silverCard', f.step3.ltv.silverCard)}
          ${numberInput('ランクアップ', 'step3.ltv.rankUp', f.step3.ltv.rankUp)}
          ${numberInput('じぶん銀行', 'step3.ltv.jibunBank', f.step3.ltv.jibunBank)}
          ${numberInput('ノートン', 'step3.ltv.norton', f.step3.ltv.norton)}
        </div>
      </details>

      <details class="collapse">
        <summary>光回線内訳</summary>
        ${ltvBreakdownGroup('auひかり', 'step3.ltv.auHikariBreakdown', f.step3.ltv.auHikariBreakdown)}
        ${ltvBreakdownGroup('BLひかり', 'step3.ltv.blHikariBreakdown', f.step3.ltv.blHikariBreakdown)}
        ${ltvBreakdownGroup('コミュファ光', 'step3.ltv.commufaHikariBreakdown', f.step3.ltv.commufaHikariBreakdown)}
      </details>
    `;
  }

  if (step === 4) {
    const cases = Array.isArray(f.step4.cases) && f.step4.cases.length > 0 ? f.step4.cases : [createEmptySuccessCase()];
    const caseHtml = cases
      .map((item, index) => `
        <div class="panel">
          <h4>成約事例 ${index + 1}</h4>
          ${selectInput('来店理由', `step4.cases.${index}.visitReason`, item.visitReason, options.successVisitReasons)}
          ${selectInput('客層', `step4.cases.${index}.customerType`, item.customerType, options.successCustomerTypes)}
          ${selectInput('決め手トーク（タグ）', `step4.cases.${index}.talkTag`, item.talkTag, options.successTalkTags)}
          ${textareaInput('決め手トーク（具体）', `step4.cases.${index}.talkDetail`, item.talkDetail)}
          ${textareaInput('成約要因', `step4.cases.${index}.contractFactor`, item.contractFactor)}
          ${textareaInput('その他', `step4.cases.${index}.other`, item.other)}
          <div class="card-actions">
            <button class="btn btn-ghost" type="button" data-action="remove-step4-case" data-index="${index}" ${cases.length === 1 ? 'disabled' : ''}>この事例を削除</button>
          </div>
        </div>
      `)
      .join('');
    return `
      <h3>STEP4: 成約事例</h3>
      <p class="hint">必要なら複数登録できます。</p>
      ${caseHtml}
      <div class="card-actions">
        <button class="btn btn-outline" type="button" data-action="add-step4-case">＋追加</button>
      </div>
    `;
  }

  if (step === 5) {
    const cases = Array.isArray(f.step5.cases) && f.step5.cases.length > 0 ? f.step5.cases : [createEmptyImproveCase()];
    const caseHtml = cases
      .map((item, index) => `
        <div class="panel">
          <h4>改善事例 ${index + 1}</h4>
          ${selectInput('改善ポイントは？', `step5.cases.${index}.improvePoint`, item.improvePoint, options.improvePoints)}
          ${textareaInput('理由（具体）', `step5.cases.${index}.reason`, item.reason)}
          ${textareaInput('その他', `step5.cases.${index}.other`, item.other)}
          <div class="card-actions">
            <button class="btn btn-ghost" type="button" data-action="remove-step5-case" data-index="${index}" ${cases.length === 1 ? 'disabled' : ''}>この改善事例を削除</button>
          </div>
        </div>
      `)
      .join('');
    return `
      <h3>STEP5: 改善事例</h3>
      ${caseHtml}
      <div class="card-actions">
        <button class="btn btn-outline" type="button" data-action="add-step5-case">＋追加</button>
      </div>
    `;
  }

  if (step === 6) {
    return `
      <h3>STEP6: イベント会場の評価</h3>
      <p class="hint">自由記述で入力してください（任意）。</p>
      ${textareaInput('イベント会場の評価', 'step5_5.venueEvaluation', f.step5_5.venueEvaluation)}
      ${textareaInput('その他', 'step5_5.other', f.step5_5.other)}
    `;
  }

  return `
    <h3>STEP7: 振り返り</h3>
    <p class="hint">最後は軽くまとめて完了します。</p>
    ${textareaInput('所感（短文）', 'step6.impression', f.step6.impression, true)}
    ${textareaInput('その他備考', 'step6.notes', f.step6.notes)}
    ${isAdminEditor() ? textareaInput('管理者専用 総括コメント', 'step6.adminSummary', f.step6.adminSummary) : ''}
  `;
}

function bindStepInputs(step) {
  const controls = elements.stepContainer.querySelectorAll('[data-path]');
  controls.forEach((control) => {
    control.addEventListener('input', onFieldInput);
    control.addEventListener('change', onFieldInput);
    if (control.type === 'number') {
      control.addEventListener('focus', onNumberFocus);
      control.addEventListener('blur', onNumberBlur);
    }
  });

  const actionButtons = elements.stepContainer.querySelectorAll('[data-action]');
  actionButtons.forEach((button) => {
    button.addEventListener('click', onStepActionClick);
  });

  if (step === 1) {
    const photoInput = document.getElementById('step1.photo');
    if (photoInput) photoInput.addEventListener('change', onPhotoChange);
  }
}

function onNumberFocus(event) {
  const input = event.target;
  if (input.value === '0') {
    input.value = '';
  }
}

function onNumberBlur(event) {
  const input = event.target;
  const path = input.dataset.path;
  if (!path) return;
  if (input.value.trim() !== '') return;
  input.value = '0';
  setByPath(state.form, path, 0);
}

function onStepActionClick(event) {
  const action = event.currentTarget.dataset.action;
  if (!action) return;

  if (action === 'add-step4-case') {
    state.form.step4.cases.push(createEmptySuccessCase());
    renderFormView();
    return;
  }

  if (action === 'remove-step4-case') {
    const index = Number(event.currentTarget.dataset.index || -1);
    if (!Number.isInteger(index) || index < 0) return;
    if (state.form.step4.cases.length <= 1) return;
    state.form.step4.cases.splice(index, 1);
    renderFormView();
    return;
  }

  if (action === 'add-step5-case') {
    state.form.step5.cases.push(createEmptyImproveCase());
    renderFormView();
    return;
  }

  if (action === 'remove-step5-case') {
    const index = Number(event.currentTarget.dataset.index || -1);
    if (!Number.isInteger(index) || index < 0) return;
    if (state.form.step5.cases.length <= 1) return;
    state.form.step5.cases.splice(index, 1);
    renderFormView();
  }
}

function onFieldInput(event) {
  const path = event.target.dataset.path;
  if (!path) return;

  const rawValue = event.target.value;
  if (path === 'step1.staffNameSelect') {
    const selected = normalizeStaffName(rawValue);
    state.form.step1.staffName = selected;
    const textInput = document.getElementById('step1.staffName');
    if (textInput) textInput.value = selected;
    if (state.errors['step1.staffName']) {
      delete state.errors['step1.staffName'];
    }
    return;
  }

  let value = event.target.type === 'number' ? toInt(rawValue) : rawValue;
  if (path === 'step1.staffName') {
    value = normalizeStaffName(value);
    event.target.value = value;
    const selectInput = document.getElementById('step1.staffNameSelect');
    if (selectInput) {
      const has = Array.from(selectInput.options).some((opt) => opt.value === value);
      selectInput.value = has ? value : '';
    }
  }
  setByPath(state.form, path, value);

  if (path === 'step1.workPlaceType') {
    if (value === '店頭SV') {
      state.form.step1.eventVenue = '';
      state.form.step1.eventOverallTarget = '';
      state.form.step1.eventVenueTarget = '';
      delete state.errors['step1.eventVenue'];
    }
    renderFormView();
    return;
  }

  if (state.errors[path]) {
    delete state.errors[path];
    renderFormView();
  }
}

function setByPath(obj, path, value) {
  const keys = path.split('.');
  const last = keys.pop();
  let cursor = obj;
  keys.forEach((key) => {
    cursor = cursor[key];
  });
  cursor[last] = value;
}

async function onPhotoChange(event) {
  const files = event.target.files ? Array.from(event.target.files) : [];
  if (files.length === 0) {
    state.form.step1.photoMeta = null;
    state.form.step1.photoDataUrl = '';
    state.form.step1.photoUrl = '';
    state.form.step1.photoFileId = '';
    state.form.step1.photos = [];
    state.photoPreview = null;
    state.photoUploading = false;
    renderFormView();
    return;
  }

  let targetFiles = files;
  if (targetFiles.length > PHOTO_MAX_COUNT) {
    targetFiles = targetFiles.slice(0, PHOTO_MAX_COUNT);
    showToast(`写真は最大${PHOTO_MAX_COUNT}枚までです。先頭${PHOTO_MAX_COUNT}枚を処理します`);
  }

  state.form.step1.photoMeta = null;
  state.form.step1.photoDataUrl = '';
  state.form.step1.photoUrl = '';
  state.form.step1.photoFileId = '';
  state.form.step1.photos = [];
  state.photoUploading = true;
  renderFormView();

  try {
    const uploadedPhotos = [];
    let hadUploadWarning = false;

    for (const file of targetFiles) {
      const compressed = await compressImageFile(file);
      try {
        const uploaded = await uploadPhotoToDrive(compressed, file);
        uploadedPhotos.push({
          name: file.name,
          type: file.type || 'image/jpeg',
          size: file.size || 0,
          url: uploaded.photoUrl,
          dataUrl: '',
          fileId: uploaded.photoFileId
        });
        if (uploaded.warning) hadUploadWarning = true;
      } catch {
        uploadedPhotos.push({
          name: file.name,
          type: file.type || 'image/jpeg',
          size: file.size || 0,
          url: '',
          dataUrl: compressed,
          fileId: ''
        });
      }
    }

    state.form.step1.photos = uploadedPhotos;
    syncLegacyPhotoFieldsFromPhotos(state.form.step1);
    const first = uploadedPhotos[0] || null;
    state.photoPreview = first ? (first.url || first.dataUrl || null) : null;

    if (hadUploadWarning) {
      showToast('一部の写真で共有設定に制限があります');
    }
  } catch (error) {
    state.form.step1.photos = [];
    syncLegacyPhotoFieldsFromPhotos(state.form.step1);
    state.photoPreview = null;
    const message = error && error.message ? `写真処理失敗: ${error.message}` : '写真の読み込みに失敗しました';
    showToast(message);
  } finally {
    state.photoUploading = false;
    renderFormView();
  }
}

function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ''));
    reader.onerror = () => reject(new Error('file read failed'));
    reader.readAsDataURL(file);
  });
}

function loadImage(dataUrl) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error('image decode failed'));
    img.src = dataUrl;
  });
}

async function compressImageFile(file) {
  const sourceDataUrl = await readFileAsDataUrl(file);
  const image = await loadImage(sourceDataUrl);

  const maxSide = Math.max(image.width, image.height);
  const scale = maxSide > PHOTO_MAX_EDGE_PX ? PHOTO_MAX_EDGE_PX / maxSide : 1;
  const width = Math.max(1, Math.round(image.width * scale));
  const height = Math.max(1, Math.round(image.height * scale));

  const canvas = document.createElement('canvas');
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext('2d');
  if (!ctx) throw new Error('canvas context failed');
  ctx.drawImage(image, 0, 0, width, height);

  return canvas.toDataURL('image/jpeg', PHOTO_JPEG_QUALITY);
}

async function uploadPhotoToDrive(photoDataUrl, file) {
  const endpoint = state.syncConfig.endpoint.trim();
  if (!endpoint) {
    throw new Error('sync endpoint not set');
  }

  const response = await fetch(endpoint, {
    method: 'POST',
    body: JSON.stringify({
      action: 'uploadPhoto',
      token: state.syncConfig.token || '',
      dataUrl: photoDataUrl,
      fileName: file && file.name ? file.name : `photo-${Date.now()}.jpg`,
      mimeType: file && file.type ? file.type : 'image/jpeg'
    })
  });

  if (!response.ok) {
    throw new Error(`upload failed: ${response.status}`);
  }
  const result = await response.json().catch(() => ({}));
  if (!result || result.ok !== true || !result.photoUrl) {
    throw new Error(result.error || 'upload rejected');
  }
  return {
    photoUrl: String(result.photoUrl),
    photoFileId: String(result.photoFileId || ''),
    warning: String(result.warning || '')
  };
}

async function preparePhotoForPersist(step1) {
  const rawPhotos = getStep1Photos(step1).slice(0, PHOTO_MAX_COUNT);
  const preparedPhotos = [];

  for (const photo of rawPhotos) {
    const record = {
      name: photo.name || '会場写真',
      type: photo.type || 'image/jpeg',
      size: Number(photo.size || 0),
      url: photo.url || '',
      dataUrl: photo.dataUrl || '',
      fileId: photo.fileId || ''
    };

    if (record.dataUrl && record.dataUrl.length > PHOTO_MAX_DATAURL_CHARS) {
      try {
        const image = await loadImage(record.dataUrl);
        const canvas = document.createElement('canvas');
        const maxSide = Math.max(image.width, image.height);
        const targetEdge = Math.min(960, maxSide);
        const scale = targetEdge / maxSide;
        canvas.width = Math.max(1, Math.round(image.width * scale));
        canvas.height = Math.max(1, Math.round(image.height * scale));
        const ctx = canvas.getContext('2d');
        if (!ctx) throw new Error('canvas context failed');
        ctx.drawImage(image, 0, 0, canvas.width, canvas.height);
        record.dataUrl = canvas.toDataURL('image/jpeg', 0.62);
      } catch {
        return {
          ok: false,
          photos: [],
          photoDataUrl: '',
          photoUrl: '',
          photoFileId: '',
          message: '写真の処理に失敗しました。再添付してください'
        };
      }
    }

    if (record.dataUrl && record.dataUrl.length > PHOTO_MAX_DATAURL_CHARS) {
      return {
        ok: false,
        photos: [],
        photoDataUrl: '',
        photoUrl: '',
        photoFileId: '',
        message: '写真サイズが大きすぎます。小さい写真を添付してください'
      };
    }

    preparedPhotos.push(record);
  }

  const first = preparedPhotos[0] || null;
  return {
    ok: true,
    photos: preparedPhotos,
    photoDataUrl: first ? first.dataUrl || '' : '',
    photoUrl: first ? first.url || '' : '',
    photoFileId: first ? first.fileId || '' : ''
  };
}

function validateStep(step, form) {
  const errors = {};

  if (step === 1) {
    if (!form.step1.workDate) errors['step1.workDate'] = '稼働日を入力してください';
    if (!form.step1.staffName.trim()) errors['step1.staffName'] = 'スタッフ名を選択してください';
    if (!form.step1.workPlaceType) errors['step1.workPlaceType'] = '区分を選択してください';
    if (!form.step1.storeName.trim()) errors['step1.storeName'] = '店舗名を入力してください';
    if (form.step1.workPlaceType !== '店頭SV' && !form.step1.eventVenue.trim()) {
      errors['step1.eventVenue'] = 'イベント会場を入力してください';
    }
  }

  if (step === 2) {
    const targets = [
      'step2.visitors',
      'step2.catchCount',
      'step2.seated',
      'step2.prospects',
      'step2.seatedBreakdown.auUqExisting',
      'step2.seatedBreakdown.sbYmobile',
      'step2.seatedBreakdown.docomoAhamo',
      'step2.seatedBreakdown.rakuten',
      'step2.seatedBreakdown.other'
    ];
    validateNumberPaths(targets, form, errors);
  }

  if (step === 3) {
    const targets = [
      'step3.newAcquisitions.auMnpSim',
      'step3.newAcquisitions.auMnpHs',
      'step3.newAcquisitions.auNewSim',
      'step3.newAcquisitions.auNewHs',
      'step3.newAcquisitions.uqMnpSim',
      'step3.newAcquisitions.uqMnpHs',
      'step3.newAcquisitions.uqNewSim',
      'step3.newAcquisitions.uqNewHs',
      'step3.newAcquisitions.cellUp',
      'step3.ltv.auDenki',
      'step3.ltv.goldCard',
      'step3.ltv.silverCard',
      'step3.ltv.rankUp',
      'step3.ltv.jibunBank',
      'step3.ltv.norton',
      'step3.ltv.auHikariBreakdown.new',
      'step3.ltv.auHikariBreakdown.fromDocomo',
      'step3.ltv.auHikariBreakdown.fromSoftbank',
      'step3.ltv.auHikariBreakdown.fromOther',
      'step3.ltv.blHikariBreakdown.new',
      'step3.ltv.blHikariBreakdown.fromDocomo',
      'step3.ltv.blHikariBreakdown.fromSoftbank',
      'step3.ltv.blHikariBreakdown.fromOther',
      'step3.ltv.commufaHikariBreakdown.new',
      'step3.ltv.commufaHikariBreakdown.fromDocomo',
      'step3.ltv.commufaHikariBreakdown.fromSoftbank',
      'step3.ltv.commufaHikariBreakdown.fromOther'
    ];
    validateNumberPaths(targets, form, errors);
  }

  if (step === 4) {
    // STEP4は任意入力
  }

  if (step === 5) {
    // STEP5は任意入力
  }

  if (step === 7) {
    if (!form.step6.impression.trim()) errors['step6.impression'] = '所感を入力してください';
  }

  return errors;
}

function validateNumberPaths(paths, form, errors) {
  paths.forEach((path) => {
    const value = getByPath(form, path);
    if (!Number.isInteger(value) || value < 0) {
      errors[path] = '0以上の整数を入力してください';
    }
  });
}

function getByPath(obj, path) {
  return path.split('.').reduce((acc, key) => acc[key], obj);
}
