const STORAGE_KEY = 'daily-report-app-v1';
const ADMIN_SESSION_KEY = 'daily-report-admin-session-v1';
const SYNC_CONFIG_KEY = 'daily-report-sync-config-v1';
const SYNC_POLL_INTERVAL_MS = 30000;
const PHOTO_MAX_EDGE_PX = 1280;
const PHOTO_JPEG_QUALITY = 0.72;
const PHOTO_MAX_DATAURL_CHARS = 500000;

// 全スタッフ端末で共通利用する既定の連携先。
// ここを設定しておくと、管理者以外でも自動同期されます。
const DEFAULT_SYNC_CONFIG = {
  endpoint: 'https://script.google.com/macros/s/AKfycbyQt0jv8wwTX2gJB3DTtPeAC-hLfksq7-TVjthaR2be1BGdJ5A-rAHy_5_y-59W0Dbw/exec',
  token: 'daily-report-token'
};

const ADMIN_USERS = [
  { id: 'admin01', name: '管理者A', password: 'admin1234' },
  { id: 'sv01', name: 'SV管理者', password: 'sv1234' }
];

const state = {
  reports: [],
  mode: 'staff-list',
  editingId: null,
  returnView: 'staff-list',
  detailReturnView: 'staff-list',
  detailReportId: '',
  adminFocusReportId: '',
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
  filters: {
    staffKeyword: '',
    staffType: 'all',
    adminKeyword: '',
    adminType: 'all',
    adminFolder: 'all'
  }
};

const views = {
  staffList: document.getElementById('view-list'),
  admin: document.getElementById('view-admin'),
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
  staffListTitle: document.getElementById('staff-list-title'),
  staffBackButton: document.getElementById('staff-back-button'),
  staffSearchKeyword: document.getElementById('staff-search-keyword'),
  staffSearchType: document.getElementById('staff-search-type'),
  staffClearButton: document.getElementById('staff-clear-button'),
  adminSearchKeyword: document.getElementById('admin-search-keyword'),
  adminSearchType: document.getElementById('admin-search-type'),
  adminFolderFilter: document.getElementById('admin-folder-filter'),
  adminClearButton: document.getElementById('admin-clear-button'),
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
  adminReportFolderInput: document.getElementById('admin-report-folder-input'),
  adminReportFolderSaveButton: document.getElementById('admin-report-folder-save-button'),
  adminReportPreviewButton: document.getElementById('admin-report-preview-button'),
  adminReportEditButton: document.getElementById('admin-report-edit-button'),
  syncEndpointInput: document.getElementById('sync-endpoint-input'),
  syncTokenInput: document.getElementById('sync-token-input'),
  syncSaveButton: document.getElementById('sync-save-button'),
  syncAllButton: document.getElementById('sync-all-button'),
  syncPullButton: document.getElementById('sync-pull-button'),
  syncStatusText: document.getElementById('sync-status-text')
};

const options = {
  customerSegments: ['', 'ファミリー', '単身', 'シニア', '学生', '法人', 'その他'],
  visitPurposes: ['', '料金見直し', 'MNP検討', '新規契約', '機種変更', '故障相談', 'その他'],
  failureCauses: ['', '在庫不足', 'ヒアリング不足', '競合優位', '訴求不足', '価格不一致', '時間不足', 'その他']
};

init();

function init() {
  state.reports = loadReports();
  state.adminUser = loadAdminSession();
  state.syncConfig = loadSyncConfig();
  bindEvents();
  void initialSyncFromSheet().finally(() => {
    renderStaffList();
    renderAdminView();
    openStaffListView();
    startSyncPolling();
  });
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
  elements.staffBackButton.addEventListener('click', backToStaffGroupList);
  elements.detailBackButton.addEventListener('click', backFromDetailView);

  elements.staffSearchKeyword.addEventListener('input', (event) => {
    state.filters.staffKeyword = event.target.value.trim();
    renderStaffList();
  });

  elements.staffSearchType.addEventListener('change', (event) => {
    state.filters.staffType = event.target.value;
    renderStaffList();
  });

  elements.staffClearButton.addEventListener('click', () => {
    state.filters.staffKeyword = '';
    state.filters.staffType = 'all';
    elements.staffSearchKeyword.value = '';
    elements.staffSearchType.value = 'all';
    renderStaffList();
  });

  elements.adminSearchKeyword.addEventListener('input', (event) => {
    state.filters.adminKeyword = event.target.value.trim();
    renderAdminLists();
  });

  elements.adminSearchType.addEventListener('change', (event) => {
    state.filters.adminType = event.target.value;
    renderAdminLists();
  });

  elements.adminFolderFilter.addEventListener('change', (event) => {
    state.filters.adminFolder = event.target.value;
    renderAdminLists();
  });

  elements.adminClearButton.addEventListener('click', () => {
    state.filters.adminKeyword = '';
    state.filters.adminType = 'all';
    state.filters.adminFolder = 'all';
    elements.adminSearchKeyword.value = '';
    elements.adminSearchType.value = 'all';
    elements.adminFolderFilter.value = 'all';
    renderAdminLists();
  });

  elements.adminLoginButton.addEventListener('click', handleAdminLogin);
  elements.adminLogoutButton.addEventListener('click', handleAdminLogout);
  elements.adminReportBackButton.addEventListener('click', openAdminView);
  elements.adminReportConfirmButton.addEventListener('click', handleAdminReportConfirm);
  elements.adminReportFolderSaveButton.addEventListener('click', handleAdminReportSaveFolder);
  elements.adminReportPreviewButton.addEventListener('click', handleAdminReportPreview);
  elements.adminReportEditButton.addEventListener('click', handleAdminReportEdit);
  elements.adminConfirmedBackButton.addEventListener('click', handleAdminConfirmedBack);
  elements.adminConfirmedDetailEditButton.addEventListener('click', handleAdminConfirmedDetailEdit);
  elements.syncSaveButton.addEventListener('click', handleSaveSyncConfig);
  elements.syncAllButton.addEventListener('click', handleSyncAllReports);
  elements.syncPullButton.addEventListener('click', handleSyncPullReports);
}

function createEmptyForm() {
  return {
    step1: {
      workDate: '',
      staffName: '',
      eventCompany: '',
      storeName: '',
      eventVenue: '',
      photoMeta: null,
      photoDataUrl: '',
      photoUrl: '',
      photoFileId: ''
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
        uqNewHs: 0
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
        silverCard: 0
      }
    },
    step4: {
      title: '',
      customerSegment: '',
      visitPurpose: '',
      keyTalk: '',
      reason: '',
      other: ''
    },
    step5: {
      title: '',
      cause: '',
      nextAction: '',
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
  return {
    id: report.id || makeId(),
    createdAt: report.createdAt || new Date().toISOString(),
    updatedAt: report.updatedAt || new Date().toISOString(),
    folder: typeof report.folder === 'string' && report.folder.trim() ? report.folder.trim() : '未分類',
    confirmed: Boolean(report.confirmed),
    payload
  };
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
    if (showToastOnSuccess) showToast('シートから取得しました');
  } catch {
    if (showToastOnSuccess) showToast('シート取得に失敗しました');
  }
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
}

function openStaffListView() {
  state.mode = 'staff-list';
  state.editingId = null;
  state.currentStep = 1;
  state.selectedStaffName = '';
  state.errors = {};
  state.photoPreview = null;
  setHeaderActiveRole('staff');
  switchView('staffList');
  renderStaffList();
}

function backToStaffGroupList() {
  state.selectedStaffName = '';
  renderStaffList();
}

function openAdminView() {
  state.mode = 'admin';
  state.editingId = null;
  state.currentStep = 1;
  state.adminConfirmedSelectedStaff = '';
  state.adminConfirmedSelectedReportId = '';
  state.errors = {};
  state.photoPreview = null;
  setHeaderActiveRole('admin');
  switchView('admin');
  renderAdminView();
}

function renderAdminView() {
  const loggedIn = Boolean(state.adminUser);

  elements.adminAuthPanel.style.display = loggedIn ? 'none' : 'block';
  elements.adminContent.style.display = loggedIn ? 'block' : 'none';

  if (!loggedIn) return;

  elements.adminUserLabel.textContent = `${state.adminUser.name}（${state.adminUser.id}）でログイン中`;
  renderSyncConfig();
  renderAdminFolderFilter();
  renderAdminLists();
}

function handleSaveSyncConfig() {
  state.syncConfig.endpoint = elements.syncEndpointInput.value.trim();
  state.syncConfig.token = elements.syncTokenInput.value.trim();
  saveSyncConfig();
  renderSyncConfig();
  startSyncPolling();
  showToast('連携設定を保存しました');
}

async function handleSyncAllReports() {
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
  showToast('管理者ログアウトしました');
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

function buildDetailHtml(report) {
  const f = report.payload;
  const photoSrc = f.step1.photoUrl || f.step1.photoDataUrl || '';
  const hasPhoto = Boolean(photoSrc);
  const hasDriveUrl = Boolean(f.step1.photoUrl);
  return `
    <h3>基本情報</h3>
    <p>日付: ${escapeHtml(f.step1.workDate || '-')}</p>
    <p>スタッフ: ${escapeHtml(f.step1.staffName || '-')}</p>
    <p>店舗名: ${escapeHtml(f.step1.storeName || '-')}</p>
    <p>イベント会場: ${escapeHtml(f.step1.eventVenue || '-')}</p>
    <p>会場写真: ${hasPhoto ? 'あり' : 'なし'}</p>
    ${hasDriveUrl ? `<p><a href="${escapeHtml(f.step1.photoUrl)}" target="_blank" rel="noopener noreferrer">写真を開く</a></p>` : ''}
    ${
      hasPhoto
        ? `<div class="photo-preview"><img src="${photoSrc}" alt="会場写真" /></div>`
        : ''
    }

    <h3>アプローチ状況</h3>
    <p>来店: ${f.step2.visitors} / キャッチ: ${f.step2.catchCount} / 着座: ${f.step2.seated} / 見込み: ${f.step2.prospects}</p>

    <h3>成功事例</h3>
    <p>タイトル: ${escapeHtml(f.step4.title || '-')}</p>
    <p>客層: ${escapeHtml(f.step4.customerSegment || '-')}</p>
    <p>来店目的: ${escapeHtml(f.step4.visitPurpose || '-')}</p>
    <p>決め手トーク: ${escapeHtml(f.step4.keyTalk || '-')}</p>
    <p>効いた理由: ${escapeHtml(f.step4.reason || '-')}</p>

    <h3>失敗事例</h3>
    <p>タイトル: ${escapeHtml(f.step5.title || '-')}</p>
    <p>原因: ${escapeHtml(f.step5.cause || '-')}</p>
    <p>次回: ${escapeHtml(f.step5.nextAction || '-')}</p>

    <h3>振り返り</h3>
    <p>所感: ${escapeHtml(f.step6.impression || '-')}</p>
    <p>備考: ${escapeHtml(f.step6.notes || '-')}</p>
    <p>管理者総括: ${escapeHtml(f.step6.adminSummary || '-')}</p>
  `;
}

function openAdminReportView(reportId) {
  const report = getReportById(reportId);
  if (!report) {
    showToast('対象の日報が見つかりませんでした');
    return;
  }
  state.adminFocusReportId = reportId;
  switchView('adminReport');
  elements.adminReportStaff.textContent = `スタッフ: ${report.payload.step1.staffName || '-'}`;
  elements.adminReportDate.textContent = `稼働日: ${report.payload.step1.workDate || '-'} / 会場: ${report.payload.step1.eventVenue || '-'}`;
  elements.adminReportFolderInput.value = report.folder || '未分類';
  elements.adminReportConfirmButton.disabled = report.confirmed;
  elements.adminReportConfirmButton.textContent = report.confirmed ? '確認済み（完了）' : '確認済みにする';
}

function handleAdminReportConfirm() {
  if (!state.adminFocusReportId) return;
  const index = state.reports.findIndex((item) => item.id === state.adminFocusReportId);
  if (index < 0) return;
  if (!state.reports[index].confirmed) {
    const previous = deepCopy(state.reports[index]);
    state.reports[index].confirmed = true;
    state.reports[index].updatedAt = new Date().toISOString();
    if (!saveReports()) {
      state.reports[index] = previous;
      return;
    }
    syncUpsert(state.reports[index]);
    showToast('確認済みにしました');
  }
  openAdminReportView(state.adminFocusReportId);
}

function handleAdminReportSaveFolder() {
  if (!state.adminFocusReportId) return;
  const folder = elements.adminReportFolderInput.value.trim() || '未分類';
  updateFolder(state.adminFocusReportId, folder);
  openAdminReportView(state.adminFocusReportId);
}

function handleAdminReportPreview() {
  if (!state.adminFocusReportId) return;
  openDetailView(state.adminFocusReportId, 'admin-report');
}

function handleAdminReportEdit() {
  if (!state.adminFocusReportId) return;
  openEditView(state.adminFocusReportId, 'admin');
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
    const name = report.payload.step1.staffName || '未設定スタッフ';
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

function getGoodText(report) {
  return [
    report.payload.step4.title,
    report.payload.step4.keyTalk,
    report.payload.step4.reason,
    report.payload.step4.other
  ]
    .join(' ')
    .toLowerCase();
}

function getBadText(report) {
  return [
    report.payload.step5.title,
    report.payload.step5.cause,
    report.payload.step5.nextAction,
    report.payload.step5.other
  ]
    .join(' ')
    .toLowerCase();
}

function matchesTypeAndKeyword(report, type, keyword) {
  const lowerKeyword = keyword.toLowerCase();
  const baseText = [
    report.payload.step1.staffName,
    report.payload.step1.workDate,
    report.payload.step6.impression,
    report.payload.step6.notes,
    report.payload.step6.adminSummary,
    report.folder || ''
  ]
    .join(' ')
    .toLowerCase();

  const goodText = getGoodText(report);
  const badText = getBadText(report);

  if (type === 'good') {
    return lowerKeyword ? goodText.includes(lowerKeyword) : goodText.trim().length > 0;
  }

  if (type === 'bad') {
    return lowerKeyword ? badText.includes(lowerKeyword) : badText.trim().length > 0;
  }

  if (!lowerKeyword) return true;
  return [baseText, goodText, badText].some((text) => text.includes(lowerKeyword));
}

function getStaffFilteredReports() {
  const keyword = state.filters.staffKeyword;
  const type = state.filters.staffType;
  return getSortedReports(state.reports).filter((report) => matchesTypeAndKeyword(report, type, keyword));
}

function getConfirmText(report) {
  return report.confirmed ? '確認済み' : '未確認';
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
    groups.forEach((group) => {
      const latestDate = group.items[0] ? group.items[0].payload.step1.workDate : '';
      list.appendChild(createStaffNameCard(group.staffName, group.items.length, latestDate));
    });
    elements.reportList.appendChild(list);
    return;
  }

  const targetGroup = groups.find((group) => group.staffName === state.selectedStaffName);
  const reports = targetGroup ? targetGroup.items : [];
  elements.staffListTitle.textContent = `日報一覧: ${state.selectedStaffName}`;
  elements.staffBackButton.style.display = 'inline-flex';
  elements.listEmpty.style.display = reports.length === 0 ? 'block' : 'none';

  const list = document.createElement('div');
  list.className = 'card-grid';
  reports.forEach((report) => {
    list.appendChild(createSimpleReportCard(report, 'staff-list'));
  });
  elements.reportList.appendChild(list);
}

function collectFolders() {
  const folders = new Set(['未分類']);
  state.reports.forEach((report) => folders.add(report.folder || '未分類'));
  return [...folders];
}

function renderAdminFolderFilter() {
  const folders = collectFolders();
  const selected = state.filters.adminFolder;
  elements.adminFolderFilter.innerHTML = '<option value="all">すべて</option>';

  folders.forEach((folder) => {
    const opt = document.createElement('option');
    opt.value = folder;
    opt.textContent = folder;
    if (folder === selected) opt.selected = true;
    elements.adminFolderFilter.appendChild(opt);
  });

  if (selected !== 'all' && !folders.includes(selected)) {
    state.filters.adminFolder = 'all';
    elements.adminFolderFilter.value = 'all';
  }
}

function getAdminFilteredReports() {
  const sorted = getSortedReports(state.reports);
  return sorted.filter((report) => {
    if (!matchesTypeAndKeyword(report, state.filters.adminType, state.filters.adminKeyword)) return false;
    if (state.filters.adminFolder === 'all') return true;
    return (report.folder || '未分類') === state.filters.adminFolder;
  });
}

function buildAdminSnippet(report) {
  const good = report.payload.step4.title ? `成功: ${report.payload.step4.title}` : '';
  const bad = report.payload.step5.title ? `失敗: ${report.payload.step5.title}` : '';
  const memo = report.payload.step6.impression ? `所感: ${report.payload.step6.impression}` : '';
  const summary = report.payload.step6.adminSummary ? `総括: ${report.payload.step6.adminSummary}` : '';
  return [good, bad, memo, summary].filter(Boolean).join(' / ') || '事例未入力';
}

function createAdminCard(report) {
  const fragment = elements.adminCardTemplate.content.cloneNode(true);
  const map = {
    staffName: report.payload.step1.staffName || '-',
    workDate: report.payload.step1.workDate || '-',
    folder: `フォルダ: ${report.folder || '未分類'}`,
    goodTag: report.payload.step4.title ? 'いい事例あり' : 'いい事例なし',
    badTag: report.payload.step5.title ? '悪い事例あり' : '悪い事例なし',
    confirmTag: report.confirmed ? '確認済み' : '未確認',
    snippet: buildAdminSnippet(report)
  };

  Object.entries(map).forEach(([key, value]) => {
    const node = fragment.querySelector(`[data-field="${key}"]`);
    if (node) node.textContent = String(value);
  });

  const folderInput = fragment.querySelector('[data-field="folderInput"]');
  folderInput.value = report.folder || '未分類';
  folderInput.dataset.id = report.id;

  const saveFolderButton = fragment.querySelector('[data-action="save-folder"]');
  saveFolderButton.dataset.id = report.id;

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

  unchecked.forEach((report) => {
    elements.adminUncheckedList.appendChild(createAdminUncheckedCard(report));
  });

  const confirmedGroups = groupReportsByStaff(confirmed);
  renderAdminConfirmedSection(confirmedGroups);
}

function renderAdminConfirmedSection(confirmedGroups) {
  elements.adminConfirmedStaffList.innerHTML = '';
  elements.adminConfirmedDateList.innerHTML = '';
  elements.adminConfirmedDetailContent.innerHTML = '';

  const hasConfirmed = confirmedGroups.length > 0;
  elements.adminConfirmedStaffEmpty.style.display = hasConfirmed ? 'none' : 'block';

  if (!state.adminConfirmedSelectedStaff) {
    elements.adminConfirmedDateWrap.style.display = 'none';
    elements.adminConfirmedDetailWrap.style.display = 'none';
    confirmedGroups.forEach((group) => {
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
  selectedGroup.items.forEach((report) => {
    const card = createSimpleReportCard(report, 'admin');
    const openButton = card.querySelector('[data-action="open"]');
    openButton.dataset.kind = 'confirmed-date';
    openButton.dataset.id = report.id;
    openButton.textContent = '内容確認';
    const previewButton = card.querySelector('[data-action="preview"]');
    previewButton.style.display = 'none';
    elements.adminConfirmedDateList.appendChild(card);
  });

  if (!state.adminConfirmedSelectedReportId) {
    elements.adminConfirmedDetailWrap.style.display = 'none';
    return;
  }

  const selectedReport = selectedGroup.items.find((report) => report.id === state.adminConfirmedSelectedReportId);
  if (!selectedReport) {
    state.adminConfirmedSelectedReportId = '';
    elements.adminConfirmedDetailWrap.style.display = 'none';
    return;
  }

  elements.adminConfirmedDetailWrap.style.display = 'block';
  elements.adminConfirmedDateTitle.textContent = `内容確認: ${selectedReport.payload.step1.workDate || '-'}`;
  elements.adminConfirmedDetailContent.innerHTML = buildDetailHtml(selectedReport);
  elements.adminConfirmedDetailEditButton.dataset.id = selectedReport.id;
}

function onStaffListClick(event) {
  const previewTarget = event.target.closest('[data-action="preview"]');
  if (previewTarget && previewTarget.dataset.kind === 'preview') {
    openDetailView(previewTarget.dataset.id, previewTarget.dataset.source || 'staff-list');
    return;
  }

  const target = event.target.closest('[data-action="open"]');
  if (!target) return;
  if (target.dataset.kind === 'staff') {
    state.selectedStaffName = target.dataset.staff || '';
    renderStaffList();
    return;
  }
  if (target.dataset.kind === 'report') {
    openEditView(target.dataset.id, target.dataset.source || 'staff-list');
  }
}

function onAdminListClick(event) {
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
    openAdminReportView(openButton.dataset.id);
    return;
  }

  const saveButton = event.target.closest('[data-action="save-folder"]');
  if (saveButton) {
    const id = saveButton.dataset.id;
    const card = saveButton.closest('.admin-card');
    const input = card.querySelector('[data-field="folderInput"]');
    const folder = input.value.trim() || '未分類';
    updateFolder(id, folder);
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
  const openButton = event.target.closest('[data-action="open"]');
  if (!openButton) return;

  if (openButton.dataset.kind === 'confirmed-staff') {
    state.adminConfirmedSelectedStaff = openButton.dataset.staff || '';
    state.adminConfirmedSelectedReportId = '';
    renderAdminLists();
    return;
  }

  if (openButton.dataset.kind === 'confirmed-date') {
    state.adminConfirmedSelectedReportId = openButton.dataset.id || '';
    renderAdminLists();
  }
}

function handleAdminConfirmedBack() {
  state.adminConfirmedSelectedStaff = '';
  state.adminConfirmedSelectedReportId = '';
  renderAdminLists();
}

function handleAdminConfirmedDetailEdit() {
  const reportId = elements.adminConfirmedDetailEditButton.dataset.id;
  if (!reportId) return;
  openEditView(reportId, 'admin');
}

function updateFolder(reportId, folderName) {
  const index = state.reports.findIndex((item) => item.id === reportId);
  if (index < 0) {
    showToast('フォルダ更新対象が見つかりませんでした');
    return;
  }
  const previous = deepCopy(state.reports[index]);
  state.reports[index].folder = folderName;
  state.reports[index].updatedAt = new Date().toISOString();
  if (!saveReports()) {
    state.reports[index] = previous;
    return;
  }
  syncUpsert(state.reports[index]);
  showToast('フォルダを更新しました');
  renderAdminView();
}

function updateConfirmedStatus(reportId) {
  const index = state.reports.findIndex((item) => item.id === reportId);
  if (index < 0) {
    showToast('対象の日報が見つかりませんでした');
    return;
  }
  const previous = deepCopy(state.reports[index]);
  state.reports[index].confirmed = !state.reports[index].confirmed;
  state.reports[index].updatedAt = new Date().toISOString();
  if (!saveReports()) {
    state.reports[index] = previous;
    return;
  }
  syncUpsert(state.reports[index]);
  showToast(state.reports[index].confirmed ? '確認済みにしました' : '未確認に戻しました');
  renderAdminView();
}

function renderFormView() {
  elements.formTitle.textContent = state.mode === 'create' ? '日報作成' : '日報編集';
  elements.stepText.textContent = `STEP ${state.currentStep}/6`;
  elements.stepProgress.setAttribute('aria-valuenow', String(state.currentStep));
  elements.stepProgressBar.style.width = `${(state.currentStep / 6) * 100}%`;
  elements.prevStepButton.disabled = state.currentStep === 1;
  elements.nextStepButton.textContent = state.currentStep < 6 ? '次へ' : state.mode === 'create' ? '保存する' : '更新する';

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

  if (state.currentStep < 6) {
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
  for (let i = 1; i <= 6; i += 1) {
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

  const now = new Date().toISOString();
  const report = {
    id: makeId(),
    createdAt: now,
    updatedAt: now,
    folder: '未分類',
    confirmed: false,
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
    const photoName = f.step1.photoMeta
      ? `選択済み: ${f.step1.photoMeta.name}${f.step1.photoUrl ? '（Drive保存済み）' : ''}`
      : f.step1.photoUrl
        ? 'Drive保存済みの写真があります'
        : '写真は未選択です';
    const effectivePreview = state.photoPreview || f.step1.photoUrl || f.step1.photoDataUrl || '';
    return `
      <h3>STEP1: 基本情報</h3>
      <p class="hint">最初は軽い入力から始めます。</p>
      ${textInput('稼働日', 'step1.workDate', f.step1.workDate, true, 'date')}
      ${textInput('スタッフ名', 'step1.staffName', f.step1.staffName, true)}
      ${textInput('店舗名', 'step1.storeName', f.step1.storeName, true)}
      ${textInput('イベント会場', 'step1.eventVenue', f.step1.eventVenue, true)}
      <div class="field-group">
        <label class="field-label" for="step1.photo">会場写真（任意）</label>
        <input id="step1.photo" type="file" accept="image/*" />
        ${state.photoUploading ? '<p class="hint">写真をGoogle Driveへアップロード中です...</p>' : ''}
        <p class="hint" id="photo-meta">${escapeHtml(photoName)}</p>
        <div class="photo-preview" id="photo-preview-wrap" ${effectivePreview ? '' : 'style="display:none"'}>
          <img id="photo-preview" alt="会場写真プレビュー" src="${effectivePreview}" />
        </div>
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
        </div>
      </details>

      <details class="collapse">
        <summary>LTV</summary>
        <div class="grid-2">
          ${numberInput('auでんき', 'step3.ltv.auDenki', f.step3.ltv.auDenki)}
          ${numberInput('ゴールドカード', 'step3.ltv.goldCard', f.step3.ltv.goldCard)}
          ${numberInput('シルバーカード', 'step3.ltv.silverCard', f.step3.ltv.silverCard)}
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
    return `
      <h3>STEP4: 成功事例</h3>
      <p class="hint">このステップが知識資産の核です。</p>
      ${textInput('成功事例タイトル', 'step4.title', f.step4.title, true, 'text', '20文字以内で入力')}
      ${selectInput('客層', 'step4.customerSegment', f.step4.customerSegment, options.customerSegments, true)}
      ${selectInput('来店目的', 'step4.visitPurpose', f.step4.visitPurpose, options.visitPurposes, true)}
      ${textareaInput('決め手トーク', 'step4.keyTalk', f.step4.keyTalk, true)}
      ${textareaInput('効いた理由', 'step4.reason', f.step4.reason, true)}
      ${textareaInput('その他', 'step4.other', f.step4.other)}
    `;
  }

  if (step === 5) {
    return `
      <h3>STEP5: 失敗事例</h3>
      ${textInput('NG事例タイトル', 'step5.title', f.step5.title, true)}
      ${selectInput('何がダメだった？', 'step5.cause', f.step5.cause, options.failureCauses, true)}
      ${textareaInput('次回どうする？', 'step5.nextAction', f.step5.nextAction, true)}
      ${textareaInput('その他', 'step5.other', f.step5.other)}
    `;
  }

  return `
    <h3>STEP6: 振り返り</h3>
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
  });

  if (step === 1) {
    const photoInput = document.getElementById('step1.photo');
    if (photoInput) photoInput.addEventListener('change', onPhotoChange);
  }
}

function onFieldInput(event) {
  const path = event.target.dataset.path;
  if (!path) return;

  const rawValue = event.target.value;
  const value = event.target.type === 'number' ? toInt(rawValue) : rawValue;
  setByPath(state.form, path, value);

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
  const file = event.target.files && event.target.files[0];
  if (!file) {
    state.form.step1.photoMeta = null;
    state.form.step1.photoDataUrl = '';
    state.form.step1.photoUrl = '';
    state.form.step1.photoFileId = '';
    state.photoPreview = null;
    state.photoUploading = false;
    renderFormView();
    return;
  }

  state.form.step1.photoMeta = {
    name: file.name,
    size: file.size,
    type: file.type
  };
  state.form.step1.photoUrl = '';
  state.form.step1.photoFileId = '';
  state.photoUploading = true;
  renderFormView();
  try {
    const compressed = await compressImageFile(file);
    state.photoPreview = compressed;
    state.form.step1.photoDataUrl = compressed;
    const uploaded = await uploadPhotoToDrive(compressed, file);
    state.form.step1.photoUrl = uploaded.photoUrl;
    state.form.step1.photoFileId = uploaded.photoFileId;
    state.form.step1.photoDataUrl = '';
    state.photoPreview = uploaded.photoUrl;
    const previewWrap = document.getElementById('photo-preview-wrap');
    const preview = document.getElementById('photo-preview');
    const photoMeta = document.getElementById('photo-meta');
    if (previewWrap) previewWrap.style.display = 'grid';
    if (preview) preview.src = uploaded.photoUrl;
    if (photoMeta) photoMeta.textContent = `選択済み: ${file.name}（Drive保存済み）`;
    if (uploaded.warning) {
      showToast('Drive保存は成功しましたが、共有設定に制限があります');
    }
  } catch (error) {
    if (!state.form.step1.photoDataUrl) {
      state.form.step1.photoMeta = null;
      state.photoPreview = null;
    }
    const photoMeta = document.getElementById('photo-meta');
    if (photoMeta) photoMeta.textContent = `選択済み: ${file.name}（Drive保存失敗）`;
    const message = error && error.message ? `Drive保存失敗: ${error.message}` : 'Drive保存に失敗しました。圧縮画像で保存します';
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
  const currentDataUrl = step1.photoDataUrl || '';
  const currentPhotoUrl = step1.photoUrl || '';
  const currentPhotoFileId = step1.photoFileId || '';

  if (currentPhotoUrl) {
    return {
      ok: true,
      photoDataUrl: '',
      photoUrl: currentPhotoUrl,
      photoFileId: currentPhotoFileId
    };
  }

  if (!currentDataUrl) {
    return {
      ok: true,
      photoDataUrl: '',
      photoUrl: '',
      photoFileId: ''
    };
  }

  let finalDataUrl = currentDataUrl;
  if (finalDataUrl.length > PHOTO_MAX_DATAURL_CHARS) {
    try {
      const image = await loadImage(finalDataUrl);
      const canvas = document.createElement('canvas');
      const maxSide = Math.max(image.width, image.height);
      const targetEdge = Math.min(960, maxSide);
      const scale = targetEdge / maxSide;
      canvas.width = Math.max(1, Math.round(image.width * scale));
      canvas.height = Math.max(1, Math.round(image.height * scale));
      const ctx = canvas.getContext('2d');
      if (!ctx) throw new Error('canvas context failed');
      ctx.drawImage(image, 0, 0, canvas.width, canvas.height);
      finalDataUrl = canvas.toDataURL('image/jpeg', 0.62);
    } catch {
      return { ok: false, photoDataUrl: '', photoUrl: '', photoFileId: '', message: '写真の処理に失敗しました。再添付してください' };
    }
  }

  if (finalDataUrl.length > PHOTO_MAX_DATAURL_CHARS) {
    return { ok: false, photoDataUrl: '', photoUrl: '', photoFileId: '', message: '写真サイズが大きすぎます。小さい写真を添付してください' };
  }

  return {
    ok: true,
    photoDataUrl: finalDataUrl,
    photoUrl: '',
    photoFileId: ''
  };
}

function validateStep(step, form) {
  const errors = {};

  if (step === 1) {
    if (!form.step1.workDate) errors['step1.workDate'] = '稼働日を入力してください';
    if (!form.step1.staffName.trim()) errors['step1.staffName'] = 'スタッフ名を入力してください';
    if (!form.step1.storeName.trim()) errors['step1.storeName'] = '店舗名を入力してください';
    if (!form.step1.eventVenue.trim()) errors['step1.eventVenue'] = 'イベント会場を入力してください';
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
      'step3.ltv.auDenki',
      'step3.ltv.goldCard',
      'step3.ltv.silverCard',
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
    if (!form.step4.title.trim()) errors['step4.title'] = '成功事例タイトルを入力してください';
    if (form.step4.title.trim().length > 20) errors['step4.title'] = 'タイトルは20文字以内で入力してください';
    if (!form.step4.customerSegment) errors['step4.customerSegment'] = '客層を選択してください';
    if (!form.step4.visitPurpose) errors['step4.visitPurpose'] = '来店目的を選択してください';
    if (!form.step4.keyTalk.trim()) errors['step4.keyTalk'] = '決め手トークを入力してください';
    if (!form.step4.reason.trim()) errors['step4.reason'] = '効いた理由を入力してください';
  }

  if (step === 5) {
    if (!form.step5.title.trim()) errors['step5.title'] = 'NG事例タイトルを入力してください';
    if (!form.step5.cause) errors['step5.cause'] = '原因を選択してください';
    if (!form.step5.nextAction.trim()) errors['step5.nextAction'] = '次回の対応を入力してください';
  }

  if (step === 6) {
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
