// ==================== UI 要素の取得 ====================

const storageHint = document.getElementById('storageHint');
const syncSchedulesButton = document.getElementById('syncSchedules');
const runMonthlyCalendarButton = document.getElementById('runMonthlyCalendar');
const runListCalendarButton = document.getElementById('runListCalendar');
const runVacationStatusMatrixButton = document.getElementById('runVacationStatusMatrix');
const runAggregateBpHoursButton = document.getElementById('runAggregateBpHours');
const runGenerateSubmitAllMembersButton = document.getElementById('runGenerateSubmitAllMembers');
const executionLog = document.getElementById('executionLog');
const calendarTargetMonthInput = document.getElementById('calendarTargetMonth');
const vacationStartDateInput = document.getElementById('vacationStartDate');
const vacationEndDateInput = document.getElementById('vacationEndDate');
const submitWriterSelect = document.getElementById('submitWriter');
const SUBMIT_WRITER_STORAGE_KEY = 'yojitsu.submitWriter';
const openCalendarGenerationButton = document.getElementById('openCalendarGeneration');
const backToRunFunctionMenuButton = document.getElementById('backToRunFunctionMenu');
const runFunctionMenuView = document.getElementById('runFunctionMenuView');
const calendarGenerationView = document.getElementById('calendarGenerationView');

const devSharePointPathInput = document.getElementById('devSharePointPath');
const devDownloadDirInput = document.getElementById('devDownloadDir');
const devDownloadSharePointFileButton = document.getElementById('devDownloadSharePointFile');
const devOpenSharePointExcelButton = document.getElementById('devOpenSharePointExcel');
const devRefreshMemberSettingsButton = document.getElementById('devRefreshMemberSettings');
const devRefreshProjectSettingsButton = document.getElementById('devRefreshProjectSettings');
const devExecutionLog = document.getElementById('devExecutionLog');

const settingToolRootDirInput = document.getElementById('settingToolRootDir');
const settingTemplateBaseDirInput = document.getElementById('settingTemplateBaseDir');
const settingOutputBaseDirInput = document.getElementById('settingOutputBaseDir');
const settingScheduleBaseDirInput = document.getElementById('settingScheduleBaseDir');
const settingRemoteWorkDirInput = document.getElementById('settingRemoteWorkDir');
const settingSharePointSiteUrlInput = document.getElementById('settingSharePointSiteUrl');
const pickSettingToolRootButton = document.getElementById('pickSettingToolRoot');
const pickSettingTemplateBaseButton = document.getElementById('pickSettingTemplateBase');
const pickSettingOutputBaseButton = document.getElementById('pickSettingOutputBase');
const pickSettingScheduleBaseButton = document.getElementById('pickSettingScheduleBase');
const saveBasicSettingsButton = document.getElementById('saveBasicSettings');
const basicSettingsMessage = document.getElementById('basicSettingsMessage');
const initSharePointSessionBasicButton = document.getElementById('initSharePointSessionBasic');
const deleteSharePointSessionBasicButton = document.getElementById('deleteSharePointSessionBasic');
const sharepointLoginStatus = document.getElementById('sharepointLoginStatus');

const menuToggle = document.getElementById('menuToggle');
const sidebarMenu = document.querySelector('.sidebar-menu');
const menuItems = document.querySelectorAll('.menu-item');
const toggleMemberGridEditButton = document.getElementById('toggleMemberGridEdit');
const saveMemberGridEditsButton = document.getElementById('saveMemberGridEdits');
const cancelMemberGridEditsButton = document.getElementById('cancelMemberGridEdits');
const holidayNameInput = document.getElementById('holidayName');
const holidayCategorySelect = document.getElementById('holidayCategory');
const holidayStartDateInput = document.getElementById('holidayStartDate');
const holidayEndDateInput = document.getElementById('holidayEndDate');
const holidayNotesInput = document.getElementById('holidayNotes');
const holidaySaveButton = document.getElementById('holidaySaveButton');
const holidayClearButton = document.getElementById('holidayClearButton');
const holidayReloadButton = document.getElementById('holidayReloadButton');
const holidayTableBody = document.getElementById('holidayTableBody');
const holidaySettingsMessage = document.getElementById('holidaySettingsMessage');

// ==================== 状態管理 ====================

let isBrowseDialogOpen = false;
let isTaskRunning = false;
let isDevActionRunning = false;
let memberSettingsData = [];
let memberSortState = { key: 'member_no', direction: 'asc' };
let isMemberGridEditMode = false;
let memberGridDraftMap = {};
let refreshMemberSettingsTable = null;
let refreshProjectSettingsTable = null;
let refreshHolidaySettingsTable = null;

// Python のログ（stderr）を、実行中の画面のログ欄へ流す
let taskLogSink = null;

// ==================== メニュー・タブ制御 ====================

function switchTab(tabId) {
  document.querySelectorAll('.tab-content').forEach(el => {
    el.classList.remove('active');
  });

  menuItems.forEach(item => {
    item.classList.remove('active');
  });

  const newTab = document.getElementById(tabId);
  if (newTab) {
    newTab.classList.add('active');
  }

  const activeMenuItem = document.querySelector(`[data-tab="${tabId}"]`);
  if (activeMenuItem) {
    activeMenuItem.classList.add('active');
  }

  if (tabId === 'tab-run') {
    switchRunView('menu');
  } else {
    updatePageHeader(tabId);
  }
}

function switchRunView(view) {
  const pageTitle = document.getElementById('pageTitle');
  const pageDescription = document.getElementById('pageDescription');

  if (view === 'calendar') {
    runFunctionMenuView.style.display = 'none';
    calendarGenerationView.style.display = 'block';
    pageTitle.textContent = 'カレンダー生成';
    pageDescription.textContent = '月間カレンダー作成と一覧作成を実行します。';
  } else {
    runFunctionMenuView.style.display = 'block';
    calendarGenerationView.style.display = 'none';
    pageTitle.textContent = '実行メニュー';
    pageDescription.textContent = '起動したい機能を選択してください。';
  }
}

function updatePageHeader(tabId) {
  const pageTitle = document.getElementById('pageTitle');
  const pageDescription = document.getElementById('pageDescription');

  const headers = {
    'tab-run': {
      title: '実行設定',
      description: '個別予定の取得と、各帳票の作成を行います。'
    },
    'tab-basic-settings': {
      title: '基本設定',
      description: 'ツール全体の基本ディレクトリを設定します。'
    },
    'tab-member-settings': {
      title: 'メンバ設定',
      description: 'プロジェクトメンバを管理します。'
    },
    'tab-holiday-settings': {
      title: '休日設定',
      description: '休日マスタ（共通カレンダ情報）を管理します。'
    },
    'tab-project-settings': {
      title: 'PJ設定',
      description: 'プロジェクトを管理します。'
    },
    'tab-dev': {
      title: '開発中',
      description: 'SharePoint ファイル操作の検証機能です。'
    }
  };

  const header = headers[tabId] || headers['tab-run'];
  pageTitle.textContent = header.title;
  pageDescription.textContent = header.description;
}

function toggleMenu() {
  sidebarMenu.classList.toggle('collapsed');
}

async function renderStorageHint() {
  if (!storageHint) return;
  try {
    const info = await window.electronAPI.getAppInfo();
    if (!info.dbExists) {
      storageHint.textContent = `⚠️ DB が見つかりません: ${info.dbFile}`;
      return;
    }
    const label = info.storageBackend === 'dummy' ? 'ダミー（ローカルフォルダ）' : 'SharePoint';
    storageHint.textContent = `取得元: ${label}。リモートの作業実績格納先から、個別予定ファイルを個別予定Dirへ取得します。`;
  } catch (_error) {
    // 表示用なので失敗しても無視
  }
}

function setRunButtonsDisabled(disabled) {
  if (syncSchedulesButton) syncSchedulesButton.disabled = disabled;
  runMonthlyCalendarButton.disabled = disabled;
  runListCalendarButton.disabled = disabled;
  if (runVacationStatusMatrixButton) runVacationStatusMatrixButton.disabled = disabled;
  if (runAggregateBpHoursButton) runAggregateBpHoursButton.disabled = disabled;
  if (runGenerateSubmitAllMembersButton) runGenerateSubmitAllMembersButton.disabled = disabled;
}

function setBrowseButtonsDisabled(disabled) {
  if (pickSettingToolRootButton) pickSettingToolRootButton.disabled = disabled;
  if (pickSettingTemplateBaseButton) pickSettingTemplateBaseButton.disabled = disabled;
  if (pickSettingOutputBaseButton) pickSettingOutputBaseButton.disabled = disabled;
  if (pickSettingScheduleBaseButton) pickSettingScheduleBaseButton.disabled = disabled;
}

function setDevButtonsDisabled(disabled) {
  if (devDownloadSharePointFileButton) devDownloadSharePointFileButton.disabled = disabled;
  if (devOpenSharePointExcelButton) devOpenSharePointExcelButton.disabled = disabled;
  if (devRefreshMemberSettingsButton) devRefreshMemberSettingsButton.disabled = disabled;
  if (devRefreshProjectSettingsButton) devRefreshProjectSettingsButton.disabled = disabled;
}

function setSharePointSessionActionButtonsDisabled(disabled) {
  if (initSharePointSessionBasicButton) {
    initSharePointSessionBasicButton.disabled = disabled;
  }
  if (deleteSharePointSessionBasicButton) {
    deleteSharePointSessionBasicButton.disabled = disabled;
  }
}

function setSharePointLoginStatus(message, color = 'var(--accent)') {
  if (!sharepointLoginStatus) return;
  sharepointLoginStatus.textContent = message;
  sharepointLoginStatus.style.color = color;
}

async function runWithBrowseLock(action) {
  if (isBrowseDialogOpen) {
    return '';
  }

  isBrowseDialogOpen = true;
  setBrowseButtonsDisabled(true);

  try {
    return await action();
  } finally {
    isBrowseDialogOpen = false;
    setBrowseButtonsDisabled(false);
  }
}

function appendExecutionLog(message) {
  const timestamp = new Date().toLocaleTimeString();
  executionLog.textContent += `[${timestamp}] ${message}\n`;
  executionLog.scrollTop = executionLog.scrollHeight;
}

function appendDevLog(message) {
  if (!devExecutionLog) return;
  const timestamp = new Date().toLocaleTimeString();
  devExecutionLog.textContent += `[${timestamp}] ${message}\n`;
  devExecutionLog.scrollTop = devExecutionLog.scrollHeight;
}

function normalizeWindowsPath(value) {
  return String(value || '').replace(/\\\\+/g, '\\').trim();
}

function isServerRelativePath(value) {
  const text = String(value || '').trim();
  return text.startsWith('/sites/');
}

function bindPathNormalization(inputElement) {
  if (!inputElement) return;
  inputElement.addEventListener('blur', () => {
    inputElement.value = normalizeWindowsPath(inputElement.value);
  });
}

function buildCalendarExecutionOptions() {
  const monthValue = (calendarTargetMonthInput?.value || '').trim();
  const yearMonth = monthValue ? monthValue.replace('-', '') : '';
  const vacationStartDate = (vacationStartDateInput?.value || '').trim();
  const vacationEndDate = (vacationEndDateInput?.value || '').trim();
  return {
    targetYearMonth: yearMonth,
    vacationStartDate,
    vacationEndDate,
  };
}

function validateVacationDateRange(executionOptions) {
  const startDate = String(executionOptions?.vacationStartDate || '').trim();
  const endDate = String(executionOptions?.vacationEndDate || '').trim();
  if (!startDate || !endDate) {
    return '休暇ステータス一覧作成には開始日と終了日の両方を指定してください';
  }
  if (startDate > endDate) {
    return '休暇ステータス一覧の開始日は終了日以前を指定してください';
  }
  return '';
}

const TASK_LABELS = {
  'monthly-calendar': '月間カレンダー作成',
  'list-calendar': '一覧作成',
  'leave-matrix': '休暇ステータス一覧作成',
  'hours-summary': '実績時間集計',
  'submit-files': '全メンバー実績表作成',
  sync: '個別予定の取得',
};

function streamTaskLog(chunk) {
  if (!taskLogSink) return;
  String(chunk).split(/\r?\n/).filter((line) => line.trim()).forEach((line) => taskLogSink(line));
}

function reportResult(label, result, log) {
  const outputs = result.outputs || [];
  if (outputs.length > 10) {
    log(`📄 ${outputs.length}件（${outputs[0]} ほか）`);
  } else {
    outputs.forEach((output) => log(`📄 ${output}`));
  }
  (result.warnings || []).forEach((w) => log(`⚠️ ${w.member ? `${w.member}: ` : ''}${w.message}`));
  if (result.success) {
    log(`✅ ${label} 完了${result.warnings?.length ? `（警告 ${result.warnings.length}件）` : ''}`);
  } else {
    log(`❌ ${label} 失敗: ${result.error || '不明なエラー'}`);
  }
}

async function runTask(task) {
  if (isTaskRunning) return;
  const label = TASK_LABELS[task];
  const options = buildCalendarExecutionOptions();
  const payload = { task };

  if (task === 'leave-matrix') {
    const message = validateVacationDateRange(options);
    if (message) {
      appendExecutionLog(`❌ 検証エラー: ${message}`);
      return;
    }
    payload.startDate = options.vacationStartDate;
    payload.endDate = options.vacationEndDate;
  } else if (task !== 'sync') {
    if (!options.targetYearMonth) {
      appendExecutionLog('❌ 検証エラー: 対象年月を指定してください');
      return;
    }
    payload.yearMonth = options.targetYearMonth;
    if (task === 'submit-files' && submitWriterSelect) {
      payload.writer = submitWriterSelect.value;
    }
  }

  isTaskRunning = true;
  setRunButtonsDisabled(true);
  taskLogSink = appendExecutionLog;
  appendExecutionLog(`--- ${label} を開始 ---`);
  if (payload.writer) {
    appendExecutionLog(`書き込み方式: ${submitWriterSelect.selectedOptions[0]?.textContent || payload.writer}`);
  }
  try {
    const result = task === 'sync'
      ? await window.electronAPI.syncSchedules()
      : await window.electronAPI.runTask(payload);
    reportResult(label, result, appendExecutionLog);
  } catch (error) {
    appendExecutionLog(`❌ エラー: ${error.message}`);
  } finally {
    taskLogSink = null;
    isTaskRunning = false;
    setRunButtonsDisabled(false);
  }
}

function init() {
  menuItems.forEach((item) => {
    item.addEventListener('click', (e) => {
      const tabId = e.currentTarget.getAttribute('data-tab');
      switchTab(tabId);

      if (window.innerWidth <= 720) {
        sidebarMenu.classList.remove('mobile');
      }
    });
  });

  if (menuToggle) {
    menuToggle.addEventListener('click', toggleMenu);
  }

  switchRunView('menu');
  renderStorageHint();
  window.electronAPI.onTaskLog(streamTaskLog);

  if (calendarTargetMonthInput) {
    const now = new Date();
    const y = now.getFullYear();
    const m = String(now.getMonth() + 1).padStart(2, '0');
    calendarTargetMonthInput.value = `${y}-${m}`;
  }

  if (syncSchedulesButton) {
    syncSchedulesButton.addEventListener('click', () => runTask('sync'));
  }

  // 書き込み方式は端末ごとに覚えておく（Excel の有無は端末で決まるため）
  if (submitWriterSelect) {
    try {
      const saved = localStorage.getItem(SUBMIT_WRITER_STORAGE_KEY);
      if (saved && [...submitWriterSelect.options].some((o) => o.value === saved)) {
        submitWriterSelect.value = saved;
      }
    } catch (_error) {
      // 保存領域が使えなくても既定値で動く
    }
    submitWriterSelect.addEventListener('change', () => {
      try {
        localStorage.setItem(SUBMIT_WRITER_STORAGE_KEY, submitWriterSelect.value);
      } catch (_error) {
        // 同上
      }
    });
  }

  if (openCalendarGenerationButton) {
    openCalendarGenerationButton.addEventListener('click', () => {
      switchRunView('calendar');
    });
  }

  if (backToRunFunctionMenuButton) {
    backToRunFunctionMenuButton.addEventListener('click', () => {
      switchRunView('menu');
    });
  }

  if (runMonthlyCalendarButton) {
    runMonthlyCalendarButton.addEventListener('click', async () => {
      await runTask('monthly-calendar');
    });
  }

  if (runListCalendarButton) {
    runListCalendarButton.addEventListener('click', async () => {
      await runTask('list-calendar');
    });
  }

  if (runVacationStatusMatrixButton) {
    runVacationStatusMatrixButton.addEventListener('click', async () => {
      await runTask('leave-matrix');
    });
  }

  if (runAggregateBpHoursButton) {
    runAggregateBpHoursButton.addEventListener('click', async () => {
      await runTask('hours-summary');
    });
  }

  if (runGenerateSubmitAllMembersButton) {
    runGenerateSubmitAllMembersButton.addEventListener('click', async () => {
      await runTask('submit-files');
    });
  }

  initializeBasicSettingsTab();
  initializeMemberSettingsTab();
  initializeHolidaySettingsTab();
  initializeProjectSettingsTab();
  initializeDevTab();

  bindPathNormalization(settingToolRootDirInput);
  bindPathNormalization(settingTemplateBaseDirInput);
  bindPathNormalization(settingOutputBaseDirInput);
  bindPathNormalization(settingScheduleBaseDirInput);
  bindPathNormalization(devDownloadDirInput);
}

async function initializeBasicSettingsTab() {
  await loadBasicSettings();
  await refreshSharePointSessionStatus();

  pickSettingToolRootButton.addEventListener('click', async () => {
    const selected = await runWithBrowseLock(() => window.electronAPI.pickDirectory());
    if (selected) settingToolRootDirInput.value = normalizeWindowsPath(selected);
  });

  pickSettingTemplateBaseButton.addEventListener('click', async () => {
    const selected = await runWithBrowseLock(() => window.electronAPI.pickDirectory());
    if (selected) settingTemplateBaseDirInput.value = normalizeWindowsPath(selected);
  });

  pickSettingOutputBaseButton.addEventListener('click', async () => {
    const selected = await runWithBrowseLock(() => window.electronAPI.pickDirectory());
    if (selected) settingOutputBaseDirInput.value = normalizeWindowsPath(selected);
  });

  pickSettingScheduleBaseButton.addEventListener('click', async () => {
    const selected = await runWithBrowseLock(() => window.electronAPI.pickDirectory());
    if (selected) settingScheduleBaseDirInput.value = normalizeWindowsPath(selected);
  });

  saveBasicSettingsButton.addEventListener('click', async () => {
    await saveBasicSettings();
  });

  if (initSharePointSessionBasicButton) {
    initSharePointSessionBasicButton.addEventListener('click', async () => {
      await runSharePointSessionInitFromBasicSettings();
    });
  }

  if (deleteSharePointSessionBasicButton) {
    deleteSharePointSessionBasicButton.addEventListener('click', async () => {
      await deleteSharePointSessionFromBasicSettings();
    });
  }
}

async function refreshSharePointSessionStatus() {
  try {
    const result = await window.electronAPI.getSharePointSessionStatus();
    if (result && result.exists) {
      setSharePointLoginStatus('セッション保存済み', 'green');
    } else {
      setSharePointLoginStatus('セッションファイル無し（初期化してください）', '#b35a00');
    }
  } catch (_error) {
    setSharePointLoginStatus('セッションファイル無し（初期化してください）', '#b35a00');
  }
}

async function runSharePointSessionInitFromBasicSettings() {
  if (isDevActionRunning) return;
  isDevActionRunning = true;
  setDevButtonsDisabled(true);
  setSharePointSessionActionButtonsDisabled(true);
  setSharePointLoginStatus('ログイン画面でサインインしてください…', 'var(--accent)');

  try {
    const result = await window.electronAPI.loginSharePoint();
    if (result.success) {
      setSharePointLoginStatus('ログイン成功（セッション保存済み）', 'green');
    } else {
      setSharePointLoginStatus(`ログイン失敗: ${result.error || '不明なエラー'}`, 'red');
    }
  } catch (error) {
    setSharePointLoginStatus(`ログイン失敗: ${error.message}`, 'red');
  } finally {
    isDevActionRunning = false;
    setDevButtonsDisabled(false);
    setSharePointSessionActionButtonsDisabled(false);
  }
}

async function deleteSharePointSessionFromBasicSettings() {
  if (isDevActionRunning) return;
  isDevActionRunning = true;
  setDevButtonsDisabled(true);
  setSharePointSessionActionButtonsDisabled(true);

  try {
    const result = await window.electronAPI.deleteSharePointSession();
    if (!result || !result.success) {
      throw new Error(result?.error || 'セッションの削除に失敗しました');
    }
    setSharePointLoginStatus('セッションファイル無し（初期化してください）', '#b35a00');
  } catch (error) {
    setSharePointLoginStatus(`削除エラー: ${error.message}`, 'red');
  } finally {
    isDevActionRunning = false;
    setDevButtonsDisabled(false);
    setSharePointSessionActionButtonsDisabled(false);
  }
}

async function getConfigValue(toolName, configKey) {
  const result = await queryDatabase('get-config', { tool_name: toolName, config_key: configKey });
  if (!result || !result.success) {
    throw new Error(result?.error || `設定 ${toolName}.${configKey} の取得に失敗しました`);
  }
  return String(result.data?.value || '').trim();
}

async function setConfigValue(toolName, configKey, value) {
  const result = await queryDatabase('set-config', {
    tool_name: toolName,
    config_key: configKey,
    config_value: value,
    data_type: 'string',
  });
  if (!result || !result.success) {
    throw new Error(result?.error || `設定 ${toolName}.${configKey} の保存に失敗しました`);
  }
}

async function loadBasicSettings() {
  try {
    const toolSettings = await fetchToolSettings();
    settingToolRootDirInput.value = normalizeWindowsPath(toolSettings.tool_root_dir || '');
    settingTemplateBaseDirInput.value = normalizeWindowsPath(toolSettings.template_base_dir || '');
    settingOutputBaseDirInput.value = normalizeWindowsPath(toolSettings.output_base_dir || '');
    settingScheduleBaseDirInput.value = normalizeWindowsPath(toolSettings.schedule_base_dir || '');
    if (settingRemoteWorkDirInput) {
      settingRemoteWorkDirInput.value = await getConfigValue('sharepoint', 'remote_work_dir');
    }
    if (settingSharePointSiteUrlInput) {
      settingSharePointSiteUrlInput.value = await getConfigValue('sharepoint', 'site_url');
    }
  } catch (error) {
    basicSettingsMessage.textContent = `基本設定の読み込みに失敗しました: ${error.message}`;
    basicSettingsMessage.style.color = 'red';
  }
}

async function saveBasicSettings() {
  const paths = {
    tool_root_dir: normalizeWindowsPath(settingToolRootDirInput.value),
    template_base_dir: normalizeWindowsPath(settingTemplateBaseDirInput.value),
    output_base_dir: normalizeWindowsPath(settingOutputBaseDirInput.value),
    schedule_base_dir: normalizeWindowsPath(settingScheduleBaseDirInput.value),
  };
  const remoteWorkDir = (settingRemoteWorkDirInput?.value || '').trim();
  const siteUrl = (settingSharePointSiteUrlInput?.value || '').trim();

  const showError = (message) => {
    basicSettingsMessage.textContent = message;
    basicSettingsMessage.style.color = 'red';
  };
  if (Object.values(paths).some((value) => !value)) {
    showError('4つの基本設定パスをすべて指定してください。');
    return;
  }
  if (remoteWorkDir && !isServerRelativePath(remoteWorkDir)) {
    showError('作業実績格納先は /sites/... で始まる server-relative path で指定してください（共有用の短縮URLは使えません）。');
    return;
  }
  if (siteUrl && !/^https:\/\//i.test(siteUrl)) {
    showError('SharePoint サイトURLは https:// で始まるURLを指定してください。');
    return;
  }

  try {
    const result = await queryDatabase('update-tool-settings', paths);
    if (!result || !result.success) {
      throw new Error(result?.error || '設定保存に失敗しました');
    }
    await setConfigValue('sharepoint', 'remote_work_dir', remoteWorkDir);
    await setConfigValue('sharepoint', 'site_url', siteUrl);
    await loadBasicSettings();
    basicSettingsMessage.textContent = '✅ 設定を保存しました。';
    basicSettingsMessage.style.color = 'green';
  } catch (error) {
    showError(`設定保存に失敗しました: ${error.message}`);
  }
}

async function initializeMemberSettingsTab() {
  const memberTableBody = document.getElementById('memberTableBody');
  const memberSettingsMessage = document.getElementById('memberSettingsMessage');
  const sortHeaders = document.querySelectorAll('#memberTable thead th[data-member-sort]');

  if (!memberTableBody || !memberSettingsMessage) {
    return;
  }

  const normalizeForSort = (value) => String(value || '').toLocaleLowerCase();
  const editableFields = ['member_no', 'full_name', 'display_name', 'group_name', 'organization', 'abbreviation', 'file_storage_location', 'is_proprietary'];

  const normalizeMemberDraft = (member) => ({
    member_no: String(member.member_no || '').trim(),
    full_name: String(member.full_name || '').trim(),
    display_name: String(member.display_name || '').trim(),
    group_name: String(member.group_name || '').trim(),
    organization: String(member.organization || '').trim(),
    abbreviation: String(member.abbreviation || '').trim(),
    file_storage_location: String(member.file_storage_location || 'internal').trim(),
    is_proprietary: member.is_proprietary ? 1 : 0,
  });

  const resetGridDraftFromData = () => {
    memberGridDraftMap = {};
    memberSettingsData.forEach((member) => {
      memberGridDraftMap[member.member_id] = normalizeMemberDraft(member);
    });
  };

  const hasGridChanges = () => {
    return memberSettingsData.some((member) => {
      const draft = memberGridDraftMap[member.member_id] || normalizeMemberDraft(member);
      return editableFields.some((field) => String(draft[field] || '') !== String(member[field] || ''));
    });
  };

  const setGridEditControls = () => {
    if (!toggleMemberGridEditButton || !saveMemberGridEditsButton || !cancelMemberGridEditsButton) {
      return;
    }

    toggleMemberGridEditButton.textContent = isMemberGridEditMode ? '通常モードに戻す' : 'グリッド編集モード';
    saveMemberGridEditsButton.style.display = isMemberGridEditMode ? 'inline-block' : 'none';
    cancelMemberGridEditsButton.style.display = isMemberGridEditMode ? 'inline-block' : 'none';
    saveMemberGridEditsButton.disabled = !isMemberGridEditMode || !hasGridChanges();
    cancelMemberGridEditsButton.disabled = !isMemberGridEditMode;
  };

  const applyMemberSort = (items) => {
    const sorted = [...items].sort((a, b) => {
      const key = memberSortState.key;
      let comp = 0;

      if (key === 'is_proprietary') {
        const av = Number(a.is_proprietary ? 1 : 0);
        const bv = Number(b.is_proprietary ? 1 : 0);
        comp = av - bv;
      } else {
        comp = normalizeForSort(a[key]).localeCompare(normalizeForSort(b[key]), 'ja');
      }

      if (comp === 0) {
        comp = normalizeForSort(a.member_id).localeCompare(normalizeForSort(b.member_id), 'ja');
      }

      return memberSortState.direction === 'asc' ? comp : -comp;
    });

    return sorted;
  };

  const updateSortHeaderLabels = () => {
    sortHeaders.forEach((header) => {
      const key = header.getAttribute('data-member-sort');
      const label = header.getAttribute('data-member-sort-label') || '';
      if (key === memberSortState.key) {
        header.textContent = `${label} ${memberSortState.direction === 'asc' ? '↑' : '↓'}`;
      } else {
        header.textContent = label;
      }
    });
  };

  const escapeHtml = (value) => String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');

  const renderMemberRows = () => {
    const sortedMembers = applyMemberSort(memberSettingsData);
    memberTableBody.innerHTML = '';

    sortedMembers.forEach((member) => {
      const row = document.createElement('tr');
      const staffType = member.is_proprietary ? 'プロパー' : 'BP';

      if (isMemberGridEditMode) {
        const draft = memberGridDraftMap[member.member_id] || normalizeMemberDraft(member);
        const isDirty = (field) => String(draft[field] || '') !== String(member[field] || '');

        row.innerHTML = `
          <td><input class="member-edit-input ${isDirty('member_no') ? 'dirty' : ''}" data-member-grid-field="member_no" data-member-id="${escapeHtml(member.member_id)}" type="text" value="${escapeHtml(draft.member_no)}"></td>
          <td><input class="member-edit-input ${isDirty('full_name') ? 'dirty' : ''}" data-member-grid-field="full_name" data-member-id="${escapeHtml(member.member_id)}" type="text" value="${escapeHtml(draft.full_name)}"></td>
          <td><input class="member-edit-input ${isDirty('display_name') ? 'dirty' : ''}" data-member-grid-field="display_name" data-member-id="${escapeHtml(member.member_id)}" type="text" value="${escapeHtml(draft.display_name)}"></td>
          <td><input class="member-edit-input ${isDirty('group_name') ? 'dirty' : ''}" data-member-grid-field="group_name" data-member-id="${escapeHtml(member.member_id)}" type="text" value="${escapeHtml(draft.group_name)}"></td>
          <td><select class="member-edit-input ${isDirty('is_proprietary') ? 'dirty' : ''}" data-member-grid-field="is_proprietary" data-member-id="${escapeHtml(member.member_id)}">
            <option value="0" ${draft.is_proprietary === 0 ? 'selected' : ''}>BP</option>
            <option value="1" ${draft.is_proprietary === 1 ? 'selected' : ''}>プロパー</option>
          </select></td>
          <td><select class="member-edit-input ${isDirty('file_storage_location') ? 'dirty' : ''}" data-member-grid-field="file_storage_location" data-member-id="${escapeHtml(member.member_id)}">
            <option value="internal" ${draft.file_storage_location === 'internal' ? 'selected' : ''}>内部</option>
            <option value="external" ${draft.file_storage_location === 'external' ? 'selected' : ''}>外部</option>
          </select></td>
          <td><input class="member-edit-input ${isDirty('organization') ? 'dirty' : ''}" data-member-grid-field="organization" data-member-id="${escapeHtml(member.member_id)}" type="text" value="${escapeHtml(draft.organization)}"></td>
          <td><input class="member-edit-input ${isDirty('abbreviation') ? 'dirty' : ''}" data-member-grid-field="abbreviation" data-member-id="${escapeHtml(member.member_id)}" type="text" value="${escapeHtml(draft.abbreviation)}"></td>
          <td>-</td>
        `;
        memberTableBody.appendChild(row);
        return;
      }

      const storageLabel = member.file_storage_location === 'external' ? '外部' : '内部';
      row.innerHTML = `
        <td>${escapeHtml(member.member_no || '')}</td>
        <td>${escapeHtml(member.full_name)}</td>
        <td>${escapeHtml(member.display_name)}</td>
        <td>${escapeHtml(member.group_name)}</td>
        <td>${staffType}</td>
        <td>${storageLabel}</td>
        <td>${escapeHtml(member.organization)}</td>
        <td>${escapeHtml(member.abbreviation)}</td>
        <td>-</td>
      `;

      memberTableBody.appendChild(row);
    });

    setGridEditControls();
  };

  const loadMembers = async () => {
    try {
      memberSettingsData = await getMembers();
      resetGridDraftFromData();
      updateSortHeaderLabels();
      renderMemberRows();
      memberSettingsMessage.textContent = '';
    } catch (error) {
      memberSettingsMessage.textContent = `エラー: ${error.message}`;
      memberSettingsMessage.style.color = 'red';
      memberTableBody.innerHTML = '';
    }
  };

  refreshMemberSettingsTable = loadMembers;

  sortHeaders.forEach((header) => {
    header.addEventListener('click', () => {
      const nextKey = header.getAttribute('data-member-sort');
      if (!nextKey) return;

      if (memberSortState.key === nextKey) {
        memberSortState.direction = memberSortState.direction === 'asc' ? 'desc' : 'asc';
      } else {
        memberSortState.key = nextKey;
        memberSortState.direction = 'asc';
      }

      updateSortHeaderLabels();
      renderMemberRows();
    });
  });

  if (toggleMemberGridEditButton && saveMemberGridEditsButton && cancelMemberGridEditsButton) {
    toggleMemberGridEditButton.addEventListener('click', () => {
      isMemberGridEditMode = !isMemberGridEditMode;
      if (isMemberGridEditMode) {
        resetGridDraftFromData();
        memberSettingsMessage.textContent = 'グリッド編集モード: 複数行を編集して一括保存できます。';
        memberSettingsMessage.style.color = 'var(--accent)';
      } else {
        memberSettingsMessage.textContent = '';
      }
      renderMemberRows();
    });

    cancelMemberGridEditsButton.addEventListener('click', () => {
      resetGridDraftFromData();
      memberSettingsMessage.textContent = 'グリッド編集の変更を破棄しました。';
      memberSettingsMessage.style.color = 'var(--accent)';
      renderMemberRows();
    });

    saveMemberGridEditsButton.addEventListener('click', async () => {
      const changedMembers = memberSettingsData
        .map((member) => {
          const draft = memberGridDraftMap[member.member_id] || normalizeMemberDraft(member);
          const changed = editableFields.some((field) => String(draft[field] || '') !== String(member[field] || ''));
          return changed ? { member, draft } : null;
        })
        .filter(Boolean);

      if (changedMembers.length === 0) {
        memberSettingsMessage.textContent = '変更はありません。';
        memberSettingsMessage.style.color = 'var(--accent)';
        return;
      }

      const invalid = changedMembers.find(({ draft }) => !String(draft.full_name || '').trim());
      if (invalid) {
        memberSettingsMessage.textContent = 'フルネームは必須です。';
        memberSettingsMessage.style.color = 'red';
        return;
      }

      saveMemberGridEditsButton.disabled = true;
      cancelMemberGridEditsButton.disabled = true;

      try {
        for (const item of changedMembers) {
          const { member, draft } = item;
          const result = await queryDatabase('update-member', {
            member_id: member.member_id,
            member_no: draft.member_no,
            full_name: draft.full_name,
            display_name: draft.display_name,
            group_name: draft.group_name,
            organization: draft.organization,
            abbreviation: draft.abbreviation,
            file_storage_location: draft.file_storage_location,
            is_proprietary: draft.is_proprietary,
          });

          if (!result || !result.success) {
            throw new Error(result?.error || `メンバ更新に失敗しました: ${member.member_id}`);
          }
        }

        memberSettingsData = memberSettingsData.map((member) => {
          const draft = memberGridDraftMap[member.member_id] || normalizeMemberDraft(member);
          return {
            ...member,
            member_no: draft.member_no,
            full_name: draft.full_name,
            display_name: draft.display_name,
            group_name: draft.group_name,
            organization: draft.organization,
            abbreviation: draft.abbreviation,
            file_storage_location: draft.file_storage_location,
            is_proprietary: draft.is_proprietary,
          };
        });

        resetGridDraftFromData();
        memberSettingsMessage.textContent = `✅ 一括保存完了（${changedMembers.length}件）`;
        memberSettingsMessage.style.color = 'green';
        renderMemberRows();
      } catch (error) {
        memberSettingsMessage.textContent = `一括保存エラー: ${error.message}`;
        memberSettingsMessage.style.color = 'red';
      } finally {
        saveMemberGridEditsButton.disabled = false;
        cancelMemberGridEditsButton.disabled = false;
      }
    });
  }

  memberTableBody.addEventListener('input', (event) => {
    if (!isMemberGridEditMode) {
      return;
    }

    const input = event.target.closest('input[data-member-grid-field][data-member-id]');
    if (!input) {
      return;
    }

    const memberId = input.getAttribute('data-member-id');
    const field = input.getAttribute('data-member-grid-field');
    if (!memberId || !field) {
      return;
    }

    if (!memberGridDraftMap[memberId]) {
      const target = memberSettingsData.find((member) => member.member_id === memberId);
      if (!target) {
        return;
      }
      memberGridDraftMap[memberId] = normalizeMemberDraft(target);
    }

    memberGridDraftMap[memberId][field] = String(input.value || '').trim();

    const base = memberSettingsData.find((member) => member.member_id === memberId);
    if (base) {
      const dirty = String(memberGridDraftMap[memberId][field] || '') !== String(base[field] || '');
      input.classList.toggle('dirty', dirty);
    }

    setGridEditControls();
  });

  memberTableBody.addEventListener('change', (event) => {
    if (!isMemberGridEditMode) {
      return;
    }

    const select = event.target.closest('select[data-member-grid-field][data-member-id]');
    if (!select) {
      return;
    }

    const memberId = select.getAttribute('data-member-id');
    const field = select.getAttribute('data-member-grid-field');
    if (!memberId || !field) {
      return;
    }

    if (!memberGridDraftMap[memberId]) {
      const target = memberSettingsData.find((member) => member.member_id === memberId);
      if (!target) {
        return;
      }
      memberGridDraftMap[memberId] = normalizeMemberDraft(target);
    }

    memberGridDraftMap[memberId][field] = String(select.value || '').trim();

    const base = memberSettingsData.find((member) => member.member_id === memberId);
    if (base) {
      const dirty = String(memberGridDraftMap[memberId][field] || '') !== String(base[field] || '');
      select.classList.toggle('dirty', dirty);
    }

    setGridEditControls();
  });

  await loadMembers();
}

async function initializeProjectSettingsTab() {
  const projectTableBody = document.getElementById('projectTableBody');
  const projectSettingsMessage = document.getElementById('projectSettingsMessage');

  const loadProjects = async () => {
    try {
      const projects = await getProjects();
      projectTableBody.innerHTML = '';

      projects.forEach(project => {
        const startDate = project.start_date || '-';
        const endDate = project.end_date || '-';
        const period = `${startDate} ～ ${endDate}`;
        const status = project.is_active ? '有効' : '無効';

        const row = document.createElement('tr');
        row.innerHTML = `
          <td>${project.project_name}</td>
          <td>${project.admin_name || 'N/A'}</td>
          <td>${period}</td>
          <td>${status}</td>
          <td>
            <button class="secondary" type="button" onclick="editProject('${project.project_id}')">編集</button>
          </td>
        `;
        projectTableBody.appendChild(row);
      });

      if (projectSettingsMessage) {
        projectSettingsMessage.textContent = '';
      }
    } catch (error) {
      if (projectSettingsMessage) {
        projectSettingsMessage.textContent = `エラー: ${error.message}`;
      }
      projectTableBody.innerHTML = '';
    }
  };

  refreshProjectSettingsTable = loadProjects;

  await loadProjects();

  const backButton = document.getElementById('backToProjectList');
  if (backButton) {
    backButton.addEventListener('click', () => {
      switchProjectView('list');
    });
  }
}

async function initializeHolidaySettingsTab() {
  if (!holidayTableBody || !holidaySaveButton || !holidaySettingsMessage) {
    return;
  }

  let editingHolidayId = null;

  const escapeHtml = (value) => String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');

  const clearForm = () => {
    editingHolidayId = null;
    holidayNameInput.value = '';
    holidayCategorySelect.value = '祝日';
    holidayStartDateInput.value = '';
    holidayEndDateInput.value = '';
    holidayNotesInput.value = '';
    holidaySaveButton.textContent = '追加';
  };

  const fillForm = (holiday) => {
    editingHolidayId = holiday.holiday_id;
    holidayNameInput.value = holiday.name || '';
    holidayCategorySelect.value = holiday.category || '祝日';
    holidayStartDateInput.value = holiday.start_date || '';
    holidayEndDateInput.value = holiday.end_date || '';
    holidayNotesInput.value = holiday.notes || '';
    holidaySaveButton.textContent = '更新';
  };

  const renderRows = (holidays) => {
    holidayTableBody.innerHTML = '';
    holidays.forEach((holiday) => {
      const row = document.createElement('tr');
      row.innerHTML = `
        <td>${escapeHtml(holiday.holiday_id)}</td>
        <td>${escapeHtml(holiday.name)}</td>
        <td>${escapeHtml(holiday.category)}</td>
        <td>${escapeHtml(holiday.start_date)}</td>
        <td>${escapeHtml(holiday.end_date || '')}</td>
        <td>${escapeHtml(holiday.notes || '')}</td>
        <td>
          <button class="secondary holiday-edit" type="button" data-holiday-id="${escapeHtml(holiday.holiday_id)}">編集</button>
          <button class="secondary holiday-delete" type="button" data-holiday-id="${escapeHtml(holiday.holiday_id)}">削除</button>
        </td>
      `;
      holidayTableBody.appendChild(row);
    });
  };

  const loadHolidays = async () => {
    const holidays = await getHolidays();
    renderRows(holidays);
    if (!editingHolidayId) {
      holidaySettingsMessage.textContent = '';
    }
  };

  refreshHolidaySettingsTable = loadHolidays;

  holidaySaveButton.addEventListener('click', async () => {
    const name = String(holidayNameInput.value || '').trim();
    const category = String(holidayCategorySelect.value || '').trim();
    const startDate = String(holidayStartDateInput.value || '').trim();
    const endDate = String(holidayEndDateInput.value || '').trim();
    const notes = String(holidayNotesInput.value || '').trim();

    if (!name || !category || !startDate) {
      holidaySettingsMessage.textContent = '名前・区分・From は必須です。';
      holidaySettingsMessage.style.color = 'red';
      return;
    }

    try {
      let result;
      if (editingHolidayId) {
        result = await queryDatabase('update-holiday', {
          holiday_id: editingHolidayId,
          name,
          category,
          start_date: startDate,
          end_date: endDate,
          notes,
        });
      } else {
        result = await queryDatabase('add-holiday', {
          name,
          category,
          start_date: startDate,
          end_date: endDate,
          notes,
        });
      }

      if (!result || !result.success) {
        throw new Error(result?.error || '休日保存に失敗しました');
      }

      await loadHolidays();
      const actionLabel = editingHolidayId ? '更新' : '追加';
      holidaySettingsMessage.textContent = `✅ 休日を${actionLabel}しました。`;
      holidaySettingsMessage.style.color = 'green';
      clearForm();
    } catch (error) {
      holidaySettingsMessage.textContent = `休日保存エラー: ${error.message}`;
      holidaySettingsMessage.style.color = 'red';
    }
  });

  holidayClearButton.addEventListener('click', () => {
    clearForm();
    holidaySettingsMessage.textContent = '入力をクリアしました。';
    holidaySettingsMessage.style.color = 'var(--accent)';
  });

  holidayReloadButton.addEventListener('click', async () => {
    try {
      await loadHolidays();
      holidaySettingsMessage.textContent = '✅ 休日一覧を再取得しました。';
      holidaySettingsMessage.style.color = 'green';
    } catch (error) {
      holidaySettingsMessage.textContent = `再取得エラー: ${error.message}`;
      holidaySettingsMessage.style.color = 'red';
    }
  });

  holidayTableBody.addEventListener('click', async (event) => {
    const editButton = event.target.closest('button.holiday-edit[data-holiday-id]');
    const deleteButton = event.target.closest('button.holiday-delete[data-holiday-id]');

    if (editButton) {
      const holidayId = Number.parseInt(editButton.getAttribute('data-holiday-id') || '', 10);
      const holidays = await getHolidays();
      const target = holidays.find((h) => Number(h.holiday_id) === holidayId);
      if (!target) {
        holidaySettingsMessage.textContent = '対象の休日が見つかりません。';
        holidaySettingsMessage.style.color = 'red';
        return;
      }
      fillForm(target);
      holidaySettingsMessage.textContent = `holiday_id=${holidayId} を編集中です。`;
      holidaySettingsMessage.style.color = 'var(--accent)';
      return;
    }

    if (deleteButton) {
      const holidayId = Number.parseInt(deleteButton.getAttribute('data-holiday-id') || '', 10);
      if (!window.confirm(`holiday_id=${holidayId} を削除します。よろしいですか？`)) {
        return;
      }
      try {
        const result = await queryDatabase('delete-holiday', { holiday_id: holidayId });
        if (!result || !result.success) {
          throw new Error(result?.error || '休日削除に失敗しました');
        }
        await loadHolidays();
        if (editingHolidayId === holidayId) {
          clearForm();
        }
        holidaySettingsMessage.textContent = '✅ 休日を削除しました。';
        holidaySettingsMessage.style.color = 'green';
      } catch (error) {
        holidaySettingsMessage.textContent = `休日削除エラー: ${error.message}`;
        holidaySettingsMessage.style.color = 'red';
      }
    }
  });

  clearForm();
  await loadHolidays();
}

async function fetchToolSettings() {
  const result = await queryDatabase('get-tool-settings', {});
  if (!result || !result.success) {
    throw new Error(result?.error || '基本設定の取得に失敗しました');
  }
  return result.data || {};
}

async function runDevAction(label, action) {
  if (isDevActionRunning) return;
  isDevActionRunning = true;
  setDevButtonsDisabled(true);
  taskLogSink = appendDevLog;
  appendDevLog(`--- ${label} を開始 ---`);
  try {
    await action();
  } catch (error) {
    appendDevLog(`❌ エラー: ${error.message}`);
  } finally {
    taskLogSink = null;
    isDevActionRunning = false;
    setDevButtonsDisabled(false);
  }
}

async function initializeDevTab() {
  if (!devDownloadSharePointFileButton || !devOpenSharePointExcelButton) {
    return;
  }

  try {
    const settings = await fetchToolSettings();
    const toolRootDir = (settings.tool_root_dir || '').trim();
    if (!devDownloadDirInput.value && toolRootDir) {
      devDownloadDirInput.value = normalizeWindowsPath(`${toolRootDir.replace(/[\\/]+$/, '')}\\download`);
    }
  } catch (error) {
    appendDevLog(`⚠️ 基本設定の取得に失敗: ${error.message}`);
  }

  devDownloadSharePointFileButton.addEventListener('click', () => {
    const remotePath = (devSharePointPathInput?.value || '').trim();
    const downloadDir = normalizeWindowsPath(devDownloadDirInput?.value || '');
    if (!remotePath || !downloadDir) {
      appendDevLog('❌ SharePoint パスとダウンロード先を入力してください。');
      return;
    }
    return runDevAction('ファイルダウンロード', async () => {
      const result = await window.electronAPI.downloadRemoteFile({ remotePath, downloadDir });
      reportResult('ファイルダウンロード', result, appendDevLog);
    });
  });

  devOpenSharePointExcelButton.addEventListener('click', () => {
    const remotePath = (devSharePointPathInput?.value || '').trim();
    if (!remotePath) {
      appendDevLog('❌ SharePoint パスを入力してください。');
      return;
    }
    return runDevAction('ローカルExcel起動', async () => {
      const result = await window.electronAPI.openRemoteInExcel({ remotePath });
      appendDevLog(result.success ? `✅ Excel で開きました: ${result.fileUrl}` : `❌ Excel 起動失敗: ${result.error}`);
    });
  });

  if (devRefreshMemberSettingsButton) {
    devRefreshMemberSettingsButton.addEventListener('click', () => runDevAction('メンバ設定の再取得', async () => {
      await (typeof refreshMemberSettingsTable === 'function' ? refreshMemberSettingsTable() : initializeMemberSettingsTab());
      appendDevLog('✅ メンバ設定を再取得しました');
    }));
  }

  if (devRefreshProjectSettingsButton) {
    devRefreshProjectSettingsButton.addEventListener('click', () => runDevAction('PJ設定の再取得', async () => {
      await (typeof refreshProjectSettingsTable === 'function' ? refreshProjectSettingsTable() : initializeProjectSettingsTab());
      appendDevLog('✅ PJ設定を再取得しました');
    }));
  }
}

function switchProjectView(view) {
  const listView = document.getElementById('projectListView');
  const detailView = document.getElementById('projectDetailView');

  if (view === 'list') {
    listView.style.display = 'block';
    detailView.style.display = 'none';
  } else if (view === 'detail') {
    listView.style.display = 'none';
    detailView.style.display = 'block';
  }
}

async function editProject(projectId) {
  try {
    const projects = await getProjects();
    const project = projects.find(p => p.project_id === projectId);

    if (!project) {
      throw new Error('プロジェクトが見つかりません');
    }

    document.getElementById('projectDetailTitle').textContent = `${project.project_name} - メンバ管理`;
    document.getElementById('detailProjectName').textContent = project.project_name;
    document.getElementById('detailAdminName').textContent = project.admin_name || 'N/A';

    const startDate = project.start_date || '-';
    const endDate = project.end_date || '-';
    document.getElementById('detailPeriod').textContent = `${startDate} ～ ${endDate}`;
    document.getElementById('detailStatus').textContent = project.is_active ? '有効' : '無効';

    const members = await getProjectMembers(projectId);
    const projectMemberTableBody = document.getElementById('projectMemberTableBody');
    projectMemberTableBody.innerHTML = '';

    members.forEach(member => {
      const memberStartDate = member.start_date || '-';
      const memberEndDate = member.end_date || '現在';
      const period = `${memberStartDate} ～ ${memberEndDate}`;
      const staffType = member.is_proprietary ? 'プロパー' : 'BP';

      const row = document.createElement('tr');
      row.innerHTML = `
        <td>${member.full_name || ''}</td>
        <td>${member.display_name || ''}</td>
        <td>${member.group_name || ''}</td>
        <td>${staffType}</td>
        <td>${period}</td>
        <td>${member.include_in_monthly ? '○' : ''}</td>
        <td>
          <button class="secondary" type="button" onclick="editProjectMember('${projectId}', '${member.member_id}')">編集</button>
        </td>
      `;
      projectMemberTableBody.appendChild(row);
    });

    switchProjectView('detail');
  } catch (error) {
    document.getElementById('projectSettingsMessage').textContent = `エラー: ${error.message}`;
  }
}

function editProjectMember(projectId, memberId) {
  alert(`メンバ編集: プロジェクト=${projectId}, メンバ=${memberId}`);
}

async function queryDatabase(action, params = {}) {
  try {
    if (!window.electronAPI || typeof window.electronAPI.queryDb !== 'function') {
      throw new Error('electronAPI が初期化されていません（preload 読み込み失敗の可能性）');
    }

    return await window.electronAPI.queryDb({ action, params });
  } catch (error) {
    return {
      success: false,
      error: String(error),
    };
  }
}

async function getProjects() {
  const result = await queryDatabase('list-projects', { active_only: 'true' });
  if (result.success) {
    return result.data.projects;
  }
  throw new Error(result.error || '不明なエラー');
}

async function getMembers() {
  const result = await queryDatabase('list-members', {});
  if (result.success) {
    return result.data.members;
  }
  throw new Error(result.error || '不明なエラー');
}

async function getProjectMembers(projectId, asOfDate = null) {
  const params = { project_id: projectId };
  if (asOfDate) {
    params.as_of_date = asOfDate;
  }

  const result = await queryDatabase('get-project-members', params);
  if (result.success) {
    return result.data.members;
  }
  throw new Error(result.error || '不明なエラー');
}

async function getHolidays() {
  const result = await queryDatabase('list-holidays', {});
  if (result.success) {
    return result.data.holidays || [];
  }
  throw new Error(result.error || '不明なエラー');
}

init();
