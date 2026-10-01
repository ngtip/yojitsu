// npm scripts から python/.venv の Python で yojitsu を実行する
const path = require('path');
const fs = require('fs');
const { spawnSync } = require('child_process');

const pythonDir = path.resolve(__dirname, '..', 'python');
const venvPython = process.platform === 'win32'
  ? path.join(pythonDir, '.venv', 'Scripts', 'python.exe')
  : path.join(pythonDir, '.venv', 'bin', 'python');
const python = fs.existsSync(venvPython) ? venvPython : 'python';

const result = spawnSync(python, ['-X', 'utf8', ...process.argv.slice(2)], { cwd: pythonDir, stdio: 'inherit' });
process.exit(result.status ?? 1);
