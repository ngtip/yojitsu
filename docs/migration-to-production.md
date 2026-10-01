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
   - `config/auth-settings.json`（あれば。手順6で削除する）
   - `assets/ms365_storage_state.json`（あれば）
   - `.env`（あれば）

## 2. コードの配置

1. このリポジトリを clone する（旧フォルダとは別の場所でよい）
2. 旧フォルダから次を新しい `electron-app/` 配下の同じ位置へコピーする
   - `assets/db/management.sqlite` → `electron-app/assets/db/management.sqlite`
   - `assets/templates/*` → `electron-app/assets/templates/`
   - `config/site-settings.json` は**旧構成には存在しない**。手順5で新しく作る
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

## 5. site-settings.json の作成

旧構成には `site-settings.json` が無い。旧構成では拠点の並び順・人数のしきい値・自社名・テンプレート名が
Python スクリプトに直接書かれていた。新構成ではこれらを `electron-app/config/site-settings.json`（Git 管理外）に置く。
旧フォルダのスクリプトから値を読み取り、雛形 `electron-app/config/site-settings.example.json` をコピーして作る。

| 項目 | 意味 | 旧構成で値が書かれている場所（目安） |
|---|---|---|
| `groups[].name` | DB の `members.group_name` と完全一致する拠点名 | `generate_list_calendar.py` / `generate_calendar.py` の拠点の並び順の定義 |
| `groups` の配列順 | 一覧・月間カレンダーの拠点の並び順 | 同上（並び順の定義） |
| `groups[].label` | 一覧カレンダー B 列に出す拠点の表示名（人数集計のキーも兼ねる） | `generate_list_calendar.py` の「拠点 → 表示名」の対応 |
| `groups[].threshold` | 拠点の座席数。その拠点に出社する人数がこれを**超えたら**（座席数+1人以上で）拠点の範囲を黄色にする。拠点ごとに値が違ってよい | `generate_list_calendar.py` の拠点別の人数しきい値。座席数をユーザーに確認する |
| `date_highlight.label` / `min_count` | この拠点の人数が `min_count` **以上**なら日付列（A列）を黄色にする | `generate_list_calendar.py` の日付強調の判定 |
| `own_company_shortname` | PJ向け作業実績表のファイル名の括弧内に入る自社略称 | `generate_submit_files_com.py` の PJ向けファイル名の組み立て |
| `own_company_fullname` | PJ向け作業実績表の H7 に入る自社名 | `generate_submit_files_com.py` の PJ向けの H7 書き込み |
| `own_template_file` / `pj_template_file` | 自社向け / PJ向けのテンプレートのファイル名 | `generate_submit_files_com.py` で開いているテンプレート（自社向けと PJ向けの2つ） |
| `groups[].keywords`（新規・任意） | 行先にこの文字を含む人をその拠点の人数として数える。未指定なら `label` | 旧構成には無い。行先の実際の表記（最寄り駅名など）をユーザーに確認する |

手順:

1. 旧スクリプトを読み、上の表の値を集める。拠点名・社名などの値はこの文書やリポジトリに書き戻さないこと
2. DB の拠点名の表記を確認し、`groups[].name` と一致させる（全角・半角、前後の空白に注意）

   ```sql
   SELECT DISTINCT group_name, length(group_name) FROM members;
   ```

3. 集めた値と、拠点の並び順・`keywords` をユーザーに見せて確認を取ってから保存する（文字コードは UTF-8）
4. 一覧カレンダーを作成し、実行ログに site-settings の警告が出ないこと、拠点ごとに並ぶこと、
   B 列の人数が旧ツールの出力と合うことを確認する

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

## 付録: 他社テナント（ゲスト参加）の SharePoint から取得できない場合の診断

ゲストとして参加している他社テナントのファイルが取得できない場合は、原因を切り分けるため診断コマンドを実行する。
出力（`assets/sharepoint-diagnose.md`）にはテナント名・サイト名・人名・Cookie の値を含めないので、
ユーザーはこれを開発環境に持ち帰って調査できる。

```bash
cd electron-app/python
# 1) 今のセッションのまま診断
.venv/Scripts/python -m yojitsu diagnose-sharepoint --db-file ../assets/db/management.sqlite \
  --target "https://<接続先テナント>.sharepoint.com/sites/<site>/Shared%20Documents/<file>.xlsx"
# 2) ブラウザ（Edge）で接続先にログインし直したセッションで診断（別ファイルに保存。通常のセッションは変えない）
.venv/Scripts/python -m yojitsu diagnose-sharepoint --db-file ../assets/db/management.sqlite --interactive-login \
  --target "https://<接続先テナント>.sharepoint.com/sites/<site>/Shared%20Documents/<file>.xlsx"
```

- `--target` は取得できないファイル（またはフォルダ）の https URL。URL はユーザーに確認する
- 2) はユーザーが画面でログイン操作（MFA 含む）をする。AI が認証情報を入力しないこと
- 「原因2（相手先のアクセス制限）」と出た場合は、回避を試みずユーザーに報告して終える
- 終わったら `assets/ms365_diagnose_state.json`（診断用のセッション）は削除してよい。`sharepoint-diagnose.md` をユーザーに渡す

## 付録: 一覧カレンダーの並びが拠点ごとにならない場合

プロパー → BP の順に全拠点が混ざって並ぶ場合は、`site-settings.json` の拠点設定が効いていない。

- 実行ログ（画面）に「site-settings.json が無いため既定値で作成しました」が出ている
  → `electron-app/config/site-settings.json` が無い。手順5に沿って作る（旧構成には存在しないのでコピーはできない）
- 「拠点「…」が site-settings.json の groups にありません」が出ている
  → `groups[].name` と DB の拠点名の表記が違う。次で DB 側の表記を確認し、`name` を完全に一致させる
  （全角・半角、前後の空白に注意）

  ```sql
  SELECT DISTINCT group_name, length(group_name) FROM members;
  ```
