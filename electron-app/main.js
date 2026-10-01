const fs = require('fs');
const path = require('path');
const { app, BrowserWindow, dialog, ipcMain, shell } = require('electron');
const { dbFile, sharePointStateFile, windowStateFile } = require('./lib/paths');
const { runYojitsu } = require('./lib/python');
const sharePointLogin = require('./lib/sharepoint-login');

const DEFAULT_BOUNDS = { width: 980, height: 720 };
const MIN_SIZE = { width: 860, height: 620 };
const TASKS = new Set(['monthly-calendar', 'list-calendar', 'leave-matrix', 'hours-summary', 'submit-files']);

// ---------- ウィンドウ ----------

function loadWindowBounds() {
  try {
    const saved = JSON.parse(fs.readFileSync(windowStateFile(app), 'utf8'));
    if (['x', 'y', 'width', 'height'].every((key) => Number.isFinite(saved[key]))) return saved;
  } catch (_error) {
    // 初回起動など
  }
  return { ...DEFAULT_BOUNDS };
}

function createWindow() {
  const win = new BrowserWindow({
    ...loadWindowBounds(),
    minWidth: MIN_SIZE.width,
    minHeight: MIN_SIZE.height,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });
  win.loadFile(path.join(__dirname, 'src', 'index.html'));
  win.on('close', () => {
    try {
      fs.writeFileSync(windowStateFile(app), JSON.stringify(win.getBounds()), 'utf8');
    } catch (_error) {
      // 位置の保存に失敗しても終了は妨げない
    }
  });
}

app.whenReady().then(() => {
  createWindow();
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

// ---------- IPC ----------

const logTo = (event) => (chunk) => {
  try {
    event.sender.send('task:log', String(chunk));
  } catch (_error) {
    // 画面が閉じられた後のログは捨てる
  }
};

async function getConfig(toolName, configKey) {
  const result = await runYojitsu(['db', '--action', 'get-config', '--tool-name', toolName, '--config-key', configKey]);
  return result.success ? String(result.data?.value || '').trim() : '';
}

ipcMain.handle('app:info', async () => ({
  dbFile: dbFile(),
  dbExists: fs.existsSync(dbFile()),
  storageBackend: process.env.YOJITSU_STORAGE_BACKEND || (await getConfig('storage', 'backend')) || 'sharepoint',
}));

ipcMain.handle('dialog:pick-directory', async (event) => {
  const result = await dialog.showOpenDialog(BrowserWindow.fromWebContents(event.sender), {
    properties: ['openDirectory'],
  });
  return result.canceled ? '' : result.filePaths[0] || '';
});

ipcMain.handle('task:run', (event, { task, yearMonth, startDate, endDate } = {}) => {
  if (!TASKS.has(task)) return { success: false, error: `不明な処理です: ${task}` };
  const args = ['run', task];
  if (yearMonth) args.push('--year-month', String(yearMonth));
  if (startDate) args.push('--start-date', String(startDate));
  if (endDate) args.push('--end-date', String(endDate));
  return runYojitsu(args, { onLog: logTo(event) });
});

ipcMain.handle('storage:sync', (event) => runYojitsu(['sync'], { onLog: logTo(event) }));

ipcMain.handle('storage:download', (event, { remotePath, downloadDir } = {}) => {
  if (!remotePath || !downloadDir) return { success: false, error: 'リモートパスと保存先を指定してください' };
  return runYojitsu(['download', '--remote-path', String(remotePath), '--download-dir', String(downloadDir)],
    { onLog: logTo(event) });
});

ipcMain.handle('sharepoint:open-in-excel', async (_event, { remotePath } = {}) => {
  try {
    const raw = String(remotePath || '').trim();
    if (!raw) throw new Error('SharePoint のパスを指定してください');
    let fileUrl = raw;
    if (!/^https?:\/\//i.test(raw)) {
      const siteUrl = await getConfig('sharepoint', 'site_url');
      if (!siteUrl) throw new Error('SharePoint サイトURL（sharepoint.site_url）が未設定です');
      fileUrl = new URL(siteUrl).origin + (raw.startsWith('/') ? raw : `/${raw}`);
    }
    fileUrl = encodeURI(fileUrl);
    await shell.openExternal(`ms-excel:ofe|u|${fileUrl}`);
    return { success: true, fileUrl };
  } catch (error) {
    return { success: false, error: String(error.message || error) };
  }
});

ipcMain.handle('sharepoint:session-status', () => ({ exists: fs.existsSync(sharePointStateFile()) }));

ipcMain.handle('sharepoint:login', async (event) => {
  const startUrl = await getConfig('sharepoint', 'site_url');
  return sharePointLogin.login({
    startUrl,
    stateFile: sharePointStateFile(),
    parent: BrowserWindow.fromWebContents(event.sender),
  });
});

ipcMain.handle('sharepoint:session-delete', async () => {
  try {
    await sharePointLogin.clearSession(sharePointStateFile());
    return { success: true };
  } catch (error) {
    return { success: false, error: String(error.message || error) };
  }
});

ipcMain.handle('db:query', async (_event, { action, params } = {}) => {
  const args = ['db', '--action', String(action)];
  for (const [key, value] of Object.entries(params || {})) {
    if (value !== null && value !== undefined) args.push(`--${key.replace(/_/g, '-')}`, String(value));
  }
  return runYojitsu(args);
});
