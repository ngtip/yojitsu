// SharePoint の手動ログイン（MFA 含む）とセッション保存。
// パスワードの保存・自動入力はしない。ログイン完了後の認証 Cookie だけを
// Playwright の storage_state 形式で保存し、Python 側の SharePoint ストレージが使う。
const fs = require('fs');
const path = require('path');
const { BrowserWindow, session } = require('electron');

const PARTITION = 'persist:sharepoint';
const LOGIN_TIMEOUT_MS = 5 * 60 * 1000;
const DEFAULT_START_URL = 'https://www.office.com/launch/sharepoint';

function isMicrosoftDomain(domain) {
  const host = String(domain || '').replace(/^\./, '').toLowerCase();
  return host.endsWith('sharepoint.com') || host.endsWith('microsoftonline.com');
}

function toPlaywrightCookie(cookie) {
  const sameSite = { strict: 'Strict', lax: 'Lax', no_restriction: 'None' }[String(cookie.sameSite).toLowerCase()] || 'Lax';
  return {
    name: cookie.name,
    value: cookie.value,
    domain: cookie.domain,
    path: cookie.path || '/',
    expires: Number.isFinite(cookie.expirationDate) ? cookie.expirationDate : -1,
    httpOnly: Boolean(cookie.httpOnly),
    secure: Boolean(cookie.secure),
    sameSite,
  };
}

async function hasAuthCookie(ses) {
  const cookies = await ses.cookies.get({});
  return cookies.some((c) => isMicrosoftDomain(c.domain) && ['fedauth', 'rtfa'].includes(String(c.name).toLowerCase()));
}

async function saveStorageState(ses, stateFile) {
  const cookies = (await ses.cookies.get({})).filter((c) => isMicrosoftDomain(c.domain));
  fs.mkdirSync(path.dirname(stateFile), { recursive: true });
  fs.writeFileSync(stateFile, JSON.stringify({ cookies: cookies.map(toPlaywrightCookie), origins: [] }, null, 2), 'utf8');
}

/** ログイン用ウィンドウを開き、認証 Cookie を検出したら保存して閉じる */
function login({ startUrl, stateFile, parent }) {
  const url = /^https?:\/\//i.test(String(startUrl || '')) ? startUrl : DEFAULT_START_URL;
  const ses = session.fromPartition(PARTITION);
  const win = new BrowserWindow({
    width: 1000,
    height: 760,
    parent: parent || undefined,
    autoHideMenuBar: true,
    title: 'SharePoint ログイン',
    webPreferences: { partition: PARTITION, contextIsolation: true, nodeIntegration: false, sandbox: true },
  });

  return new Promise((resolve) => {
    let settled = false;
    const finish = (result) => {
      if (settled) return;
      settled = true;
      clearInterval(timer);
      clearTimeout(timeout);
      if (!win.isDestroyed()) win.close();
      resolve(result);
    };

    const check = async () => {
      try {
        if (await hasAuthCookie(ses)) {
          await saveStorageState(ses, stateFile);
          finish({ success: true, stateFile });
        }
      } catch (error) {
        finish({ success: false, error: String(error.message || error) });
      }
    };

    const timer = setInterval(check, 1200);
    const timeout = setTimeout(() => finish({ success: false, error: 'ログイン完了を検出できませんでした（時間切れ）' }),
      LOGIN_TIMEOUT_MS);
    win.webContents.on('did-navigate', check);
    win.on('closed', () => finish({ success: false, error: 'ログインウィンドウが閉じられました' }));
    win.loadURL(url).catch((error) => finish({ success: false, error: String(error.message || error) }));
  });
}

async function clearSession(stateFile) {
  if (fs.existsSync(stateFile)) fs.unlinkSync(stateFile);
  await session.fromPartition(PARTITION).clearStorageData();
}

module.exports = { login, clearSession };
