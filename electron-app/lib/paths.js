const path = require('path');

const APP_ROOT = path.resolve(__dirname, '..');

function envPath(name) {
  const value = String(process.env[name] || '').trim();
  return value ? path.resolve(value) : '';
}

module.exports = {
  APP_ROOT,
  PYTHON_DIR: path.join(APP_ROOT, 'python'),
  // 環境変数で差し替えられる（サンプル環境での起動用）
  dbFile: () => envPath('YOJITSU_DB') || path.join(APP_ROOT, 'assets', 'db', 'management.sqlite'),
  sharePointStateFile: () => envPath('PLAYWRIGHT_STATE_FILE') || path.join(APP_ROOT, 'assets', 'ms365_storage_state.json'),
  windowStateFile: (app) => path.join(app.getPath('userData'), 'window-state.json'),
};
