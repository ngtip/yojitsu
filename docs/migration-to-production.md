# 本番環境への移行手順（AI 作業用）

この文書は、本番環境（会社PC）で作業する AI エージェント向けの手順書です。
旧構成（`python/scripts/*.py` を Electron から直接呼ぶ構成、制御ブック `PJメンバ予実管理.xlsm` を使う構成）から、
このリポジトリの新構成（`python -m yojitsu`）へ移行します。

## 前提と禁止事項

- 本番の DB・テンプレート・設定ファイル・SharePoint セッションは Git 管理外。**リポジトリに追加・コミットしないこと**
  （`.gitignore` 済み: `electron-app/assets/`、`electron-app/config/*.json`（example 以外）、`.env`）
- このリポジトリは公開リポジトリ。本番の人名・社名・テナント名・パスを含むファイルやコミットを作らないこと
- 各手順の前後で、ユーザーに確認を取ってから破壊的な操作（削除・上書き）をすること

## 1. 現状の確認とバックアップ

1. 旧構成のフォルダ（以下「旧フォルダ」。例: `...\electron-app`）の場所をユーザーに確認する
2. 次をまとめてバックアップする（コピーを別フォルダへ）
   - `assets/db/management.sqlite`（本番 DB）
   - `assets/templates/`（帳票テンプレート）
   - `config/site-settings.json`、`config/auth-settings.json`
   - `assets/ms365_storage_state.json`（あれば）
   - `.env`（あれば）

## 2. コードの配置

1. このリポジトリを clone する（旧フォルダとは別の場所でよい）
2. 旧フォルダから次を新しい `electron-app/` 配下の同じ位置へコピーする
   - `assets/db/management.sqlite` → `electron-app/assets/db/management.sqlite`
   - `assets/templates/*` → `electron-app/assets/templates/`
   - `config/site-settings.json` → `electron-app/config/site-settings.json`
   - `assets/ms365_storage_state.json` は**コピー不要**（ログインし直す）
3. 依存を入れる

   ```bash
   cd electron-app
   python -m venv python/.venv
   python/.venv/Scripts/python -m pip install -r python/requirements.txt
   npm install
   ```

   `requirements.txt` に `lxml`・`playwright`・`pywin32` が含まれる。
   Playwright はブラウザを起動しない方式（APIRequestContext）なので `playwright install` は不要。

## 3. DB の変更（自動）

変更は **`members.schedule_file_name` 列の追加と値の移行のみ**。ほかのテーブルの構造は変えない。

- 新しいコードが DB を最初に開いたときに自動で実行される（`yojitsu/db.py` の `_migrate`）
  - `ALTER TABLE members ADD COLUMN schedule_file_name TEXT`
  - 各メンバの `project_members.assignment_name`（`start_date` が最新のもの）を `schedule_file_name` へコピー
- 実行と確認:

  ```bash
  cd electron-app/python
  .venv/Scripts/python -m yojitsu db --db-file ../assets/db/management.sqlite --action list-members
  ```

  確認用 SQL（件数だけ見ればよい。人名を画面やログに残さない）:

  ```sql
  SELECT COUNT(*), SUM(schedule_file_name IS NOT NULL) FROM members;
  -- PJ参画履歴が無いメンバだけが NULL になる
  SELECT COUNT(*) FROM members m
  WHERE m.schedule_file_name IS NULL
    AND EXISTS (SELECT 1 FROM project_members pm WHERE pm.member_id = m.member_id);  -- 0 であること
  ```

- `project_members.assignment_name` は残るが、今後は使わない（ファイル名はメンバ側で管理）
- 個別予定のファイル名（拡張子なし）と `schedule_file_name` が一致していることを、`schedules` フォルダの一覧と突き合わせて確認する

## 4. 設定値の確認・修正（DB の tool_config / tool_settings）

| 設定 | 期待する値 | 備考 |
|---|---|---|
| `tool_settings.tool_root_dir` | 新しい `electron-app` の絶対パス | 旧フォルダのパスのままなら更新する |
| `tool_settings.template_base_dir` | `assets/templates` | tool_root_dir からの相対パス |
| `tool_settings.output_base_dir` / `schedule_base_dir` | 運用中の値 | 相対パスなら tool_root_dir 基準 |
| `tool_config` `sharepoint.site_url` | SharePoint サイトURL（`https://<tenant>.sharepoint.com/sites/<site>`） | 旧 `.env` の `SHAREPOINT_SITE_URL` に相当。`.env` は読まない |
| `tool_config` `sharepoint.remote_work_dir` | `/sites/...` で始まる server-relative path | 共有用の短縮URL（`/:f:/` を含むもの）は使えない |
| `tool_config` `storage.backend` | `sharepoint` | 未設定でも `sharepoint` 扱い |

- 画面の「基本設定」からも保存できる（パス4つ、SharePoint サイトURL、作業実績格納先）
- 使わなくなった設定（残っていても害はない）: `sharepoint.remote_work_dir_sub_url`、
  `monthly_calendar.default_template` / `list_calendar.default_template`、`*.output_subdir`、`system.default_worked_days`
- テンプレート名は `tool_config` の `monthly_calendar.template_file` / `list_calendar.template_file`
  （無ければ `月間カレンダーテンプレ.xlsx` / `一覧カレンダーテンプレ.xlsx`）

## 5. site-settings.json の確認・修正

雛形は `electron-app/config/site-settings.example.json`。

- `own_template_file`: 自社向け作業実績表のテンプレートファイル名
- `pj_template_file`: PJ向け作業実績表のテンプレートファイル名
- `groups`: 拠点。**配列の順番が一覧カレンダーの並び順**になる。並び順はユーザーに確認して合わせる
- `groups[].keywords`（新規・任意）: 一覧カレンダーの拠点別人数は「行先にこの文字を含む人」を数える。
  拠点名と行先の表記が違う拠点に設定する（例: 行先に拠点名ではなく最寄り駅名を書く運用なら、その駅名）。
  未指定なら `label` で数える。行先の実際の表記はユーザーに確認すること
- `date_highlight`: 日付列を黄色にする拠点ラベルと人数

## 6. 旧ファイルの片付け

- `config/auth-settings.json`: 旧構成で **SharePoint のパスワードを平文で保存していたファイル**。新構成では使わない。
  ユーザーに確認のうえ、バックアップも含めて削除する
- 旧フォルダの `python/scripts/`、`python/testscripts/`、制御ブック `PJメンバ予実管理.xlsm`: 新構成では使わない。
  動作確認が済むまでは旧フォルダごと残しておき、確認後に削除してよいかユーザーに聞く

## 7. 動作確認

起動:

```bash
cd electron-app
npm start
```

| # | 操作 | 確認すること |
|---|---|---|
| 1 | 基本設定 →「ログインセッション初期化」 | ログイン画面で手動ログイン（MFA 含む）できる。パスワードの自動入力はしない仕様 |
| 2 | 実行 →「リモートから個別予定を取得」 | 個別予定 Dir にファイルが揃う。外部格納（`file_storage_location = external`）のメンバは取得されない |
| 3 | 月間カレンダー・一覧カレンダー作成 | 旧版と同じ見た目か。PC持出は「有」のときだけ★・黄色になる（旧版は「無」でも付いていた） |
| 4 | 実績時間集計 | L列「参画期間」、M列「期間外の実績」が追加されている。営業日数は参画期間内で数える |
| 5 | 全メンバー実績表作成（Excel COM） | 出力に `XX月実績` シートが残らない。**別の Excel ブックを開いたまま実行し、それが閉じられないこと** |
| 6 | 全メンバー実績表作成（XML直接編集） | 出力を Excel で開いて「修復」ダイアログが出ないこと、画像が残っていること。COM の出力と内容が一致すること |
| 7 | 休暇ステータス一覧 | 期間 × メンバの記号が正しい |

問題があれば、画面の実行ログ（Python のログが流れる）と、出力ファイルの該当箇所をユーザーに報告する。
人名を含むログをリポジトリや外部に出さないこと。

## 8. 戻し方

新構成で問題が出た場合は、旧フォルダをそのまま使えば旧構成に戻る（旧フォルダには手を入れていない前提）。
DB は手順1のバックアップを戻す（`schedule_file_name` 列が増えた DB でも旧構成は動くが、戻すならバックアップを使う）。
