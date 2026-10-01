// Python（yojitsu パッケージ）の呼び出し。stdout は結果 JSON 1件、stderr はログ
const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');
const { PYTHON_DIR, dbFile } = require('./paths');

function pythonCandidates() {
  const candidates = [];
  if (process.env.YOJITSU_PYTHON) {
    candidates.push({ command: process.env.YOJITSU_PYTHON, prefix: [] });
  }
  const venvPython = process.platform === 'win32'
    ? path.join(PYTHON_DIR, '.venv', 'Scripts', 'python.exe')
    : path.join(PYTHON_DIR, '.venv', 'bin', 'python');
  if (fs.existsSync(venvPython)) {
    candidates.push({ command: venvPython, prefix: [] });
  }
  candidates.push({ command: 'python', prefix: [] }, { command: 'py', prefix: ['-3'] });
  return candidates;
}

function spawnOnce(command, args, onLog) {
  return new Promise((resolve) => {
    let stdout = '';
    let stderr = '';
    let child;
    try {
      child = spawn(command, args, {
        cwd: PYTHON_DIR,
        shell: false,
        windowsHide: true,
        env: { ...process.env, PYTHONUTF8: '1', PYTHONIOENCODING: 'utf-8' },
      });
    } catch (error) {
      resolve({ started: false, error: String(error) });
      return;
    }
    child.stdout.setEncoding('utf8');
    child.stderr.setEncoding('utf8');
    child.stdout.on('data', (chunk) => { stdout += chunk; });
    child.stderr.on('data', (chunk) => {
      stderr += chunk;
      if (onLog) onLog(chunk);
    });
    child.on('error', (error) => resolve({ started: false, error: String(error) }));
    child.on('close', (code) => resolve({ started: true, code: code ?? 1, stdout, stderr }));
  });
}

function parseResult(stdout) {
  const lines = String(stdout || '').trim().split(/\r?\n/).filter(Boolean);
  for (let i = lines.length - 1; i >= 0; i -= 1) {
    try {
      return JSON.parse(lines[i]);
    } catch (_error) {
      // JSON 以外の行は読み飛ばす
    }
  }
  return null;
}

/**
 * python -m yojitsu <args> --db-file <db> を実行する
 * @returns {{success: boolean, outputs: string[], warnings: object[], error?: string, data?: any, stderr: string}}
 */
async function runYojitsu(args, { onLog, withDb = true } = {}) {
  const fullArgs = ['-X', 'utf8', '-m', 'yojitsu', ...args];
  if (withDb) fullArgs.push('--db-file', dbFile());

  for (const { command, prefix } of pythonCandidates()) {
    const result = await spawnOnce(command, [...prefix, ...fullArgs], onLog);
    if (!result.started) continue;

    const parsed = parseResult(result.stdout);
    if (!parsed) {
      return {
        success: false,
        outputs: [],
        warnings: [],
        error: `Python の結果を読み取れませんでした（終了コード ${result.code}）`,
        stderr: result.stderr,
      };
    }
    return { outputs: [], warnings: [], ...parsed, stderr: result.stderr };
  }
  return {
    success: false,
    outputs: [],
    warnings: [],
    error: 'Python が見つかりません（python/.venv を作成するか、python / py を PATH に通してください）',
    stderr: '',
  };
}

module.exports = { runYojitsu };
