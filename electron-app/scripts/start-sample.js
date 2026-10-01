// 架空データのサンプル環境（sample-env/）を使って Electron を起動する
const path = require('path');
const fs = require('fs');
const { spawn } = require('child_process');

const appRoot = path.resolve(__dirname, '..');
const sampleRoot = path.join(appRoot, 'sample-env');
const db = path.join(sampleRoot, 'db', 'management.sqlite');

if (!fs.existsSync(db)) {
  console.error('サンプル環境がありません。先に npm run sample を実行してください。');
  process.exit(1);
}

const electron = require('electron');
const child = spawn(electron, [appRoot], {
  stdio: 'inherit',
  env: {
    ...process.env,
    YOJITSU_DB: db,
    YOJITSU_SITE_SETTINGS: path.join(sampleRoot, 'site-settings.json'),
    YOJITSU_SUBMIT_WRITER: process.env.YOJITSU_SUBMIT_WRITER || 'openpyxl',
  },
});
child.on('close', (code) => process.exit(code ?? 0));
