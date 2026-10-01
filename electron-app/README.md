# PJメンバ予実管理ツール

メンバごとの個別予定ファイル（Excel）を集め、月間カレンダー・一覧カレンダー・休暇ステータス一覧・
実績時間集計・作業実績表（提出用）を作る Electron + Python のツール。

## 構成

```
electron-app/
  main.js / preload.js      Electron（IPC は lib/ に分割）
  lib/python.js             Python 呼び出し（stdout=結果JSON, stderr=ログ）
  lib/sharepoint-login.js   SharePoint 手動ログインとセッション保存
  src/                      画面
  config/                   環境固有設定（*.example.json 以外は Git 管理外）
  python/yojitsu/           Python 本体（python -m yojitsu）
    db.py                     SQLite（PJ・メンバ・祝日・設定の正本）
    schedule.py               個別予定ファイルの読み込み（全機能で共通）
    holidays.py / members.py  祝日カレンダー / PJメンバ
    features/                 各帳票
    storage/                  リモートストレージ（sharepoint / dummy）
    sample.py                 架空データのテスト環境
  python/tests/             pytest
```

正本データはすべて SQLite（既定: `assets/db/management.sqlite`）。制御ブック（PJメンバ予実管理.xlsm）は使わない。

## 開発・テスト（実データなし）

```bash
cd electron-app
python -m venv python/.venv
python/.venv/Scripts/python -m pip install -r python/requirements-dev.txt
npm install
npm test               # pytest
npm run sample         # 架空データの sample-env/ を作る
npm run start:sample   # sample-env を使って起動
```

sample-env ではストレージが `dummy` になっており、`sample-env/remote/` を SharePoint に見立てて
「リモートから個別予定を取得」が動く。人名・社名はすべて架空。

## 実環境への移行

1. 依存を入れる: `python/.venv/Scripts/python -m pip install -r python/requirements.txt`
2. `assets/db/management.sqlite`、`assets/templates/`、`config/site-settings.json` を実環境の物で配置する
   （いずれも Git 管理外。`config/site-settings.example.json` が雛形）
3. 既存 DB に次の設定を入れる（画面の基本設定から保存できる）
   - `sharepoint.site_url` … SharePoint サイトURL（旧 .env の SHAREPOINT_SITE_URL）
   - `sharepoint.remote_work_dir` … 作業実績格納先。`/sites/...` の server-relative path（短縮URLは不可）
4. テンプレート名が既定と違う場合は `tool_config` に設定する
   - `monthly_calendar.template_file`（既定: 月間カレンダーテンプレ.xlsx）
   - `list_calendar.template_file`（既定: 一覧カレンダーテンプレ.xlsx）
5. 基本設定の「ログインセッション初期化」で手動ログイン（MFA 含む）してから使う

ストレージは `storage.backend`（`sharepoint` / `dummy`）で切り替える。未設定なら `sharepoint`。
環境変数 `YOJITSU_STORAGE_BACKEND` があればそちらが優先。

### 実環境でしか確認できていないこと

- SharePoint 取得: 旧実装と同じ REST API・Cookie を使うが、ブラウザ起動ではなく Playwright の
  APIRequestContext で呼ぶ形に変えた
- 作業実績表: 実環境は Excel COM で書く（テンプレートの画像を保持するため）。COM が使えない環境では
  openpyxl に切り替わる（`YOJITSU_SUBMIT_WRITER=com|openpyxl` で固定可）

## 環境変数

| 変数 | 用途 |
|---|---|
| `YOJITSU_DB` | DB ファイルの場所 |
| `YOJITSU_SITE_SETTINGS` | site-settings.json の場所 |
| `YOJITSU_STORAGE_BACKEND` / `YOJITSU_DUMMY_ROOT` | ストレージの切替 |
| `YOJITSU_SUBMIT_WRITER` | 作業実績表の書き込み方式（auto / com / openpyxl） |
| `YOJITSU_PYTHON` | 使う Python（未指定なら python/.venv → python → py） |
| `SHAREPOINT_SITE_URL` / `PLAYWRIGHT_STATE_FILE` | SharePoint サイトURL / セッションファイル |

## コマンド

```bash
python -m yojitsu run monthly-calendar --db-file DB --year-month 202610
python -m yojitsu run leave-matrix --db-file DB --start-date 2026-08-01 --end-date 2026-08-31
python -m yojitsu sync --db-file DB
python -m yojitsu db --db-file DB --action list-holidays
```
