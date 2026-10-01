"""コマンドライン入口（Electron から呼ぶ）

  python -m yojitsu run <task> --db-file DB [--project-id ID] --year-month YYYYMM
  python -m yojitsu run leave-matrix --db-file DB --start-date D --end-date D
  python -m yojitsu sync --db-file DB
  python -m yojitsu download --db-file DB --remote-path PATH --download-dir DIR
  python -m yojitsu db --db-file DB --action <action> [...]
  python -m yojitsu sample --dest DIR

stdout には結果 JSON を1つだけ出す。ログと進捗は stderr。
"""

import argparse
import json
import logging
import sys
from pathlib import Path
from typing import Any, Callable, Dict, Optional

from .context import RunContext
from .dates import parse_date_arg, parse_yyyymm
from .db import Database

logger = logging.getLogger('yojitsu')

TASKS = ('monthly-calendar', 'list-calendar', 'leave-matrix', 'hours-summary', 'submit-files')


def _run_task(args) -> Dict[str, Any]:
    from .features import hours_summary, leave_matrix, list_calendar, monthly_calendar, submit_files

    with Database(args.db_file) as db:
        ctx = RunContext(db, args.project_id)
        logger.info(f"PJ: {ctx.project_name} / 個別予定: {ctx.paths.schedules} / 出力: {ctx.paths.output}")
        if args.task == 'leave-matrix':
            if not (args.start_date and args.end_date):
                raise ValueError('休暇ステータス一覧には --start-date と --end-date が必要です')
            result = leave_matrix.run(ctx, parse_date_arg(args.start_date), parse_date_arg(args.end_date))
        else:
            if not args.year_month:
                raise ValueError('--year-month を指定してください')
            year, month = parse_yyyymm(args.year_month)
            if args.task == 'monthly-calendar':
                result = monthly_calendar.run(ctx, year, month)
            elif args.task == 'list-calendar':
                result = list_calendar.run(ctx, year, month)
            elif args.task == 'hours-summary':
                result = hours_summary.run(ctx, year, month)
            else:
                result = submit_files.run(ctx, year, month, args.writer)
        result['warnings'] = [w.to_dict() for w in ctx.warnings]
        return result


def _sync(args) -> Dict[str, Any]:
    from .settings import resolve_tool_paths
    from .storage import create_storage, remote_work_dir, sync_directory

    with Database(args.db_file) as db:
        local_dir = resolve_tool_paths(db).schedules
        remote_dir = remote_work_dir(db)
        exclude = db.get_external_member_aliases()
        storage = create_storage(db)
    try:
        logger.info(f"同期: {remote_dir} -> {local_dir}")
        saved = sync_directory(storage, remote_dir, local_dir, exclude)
    finally:
        storage.close()
    return {'outputs': [str(p) for p in saved], 'data': {'downloaded_count': len(saved)}}


def _download(args) -> Dict[str, Any]:
    from .storage import RemoteFile, create_storage
    from .storage.base import sanitize
    from .storage.sharepoint import to_server_relative

    with Database(args.db_file) as db:
        storage = create_storage(db)
    try:
        remote = RemoteFile(to_server_relative(args.remote_path))
        target_dir = Path(args.download_dir)
        target_dir.mkdir(parents=True, exist_ok=True)
        target = target_dir / sanitize(remote.name)
        storage.download(remote, target)
    finally:
        storage.close()
    return {'outputs': [str(target)]}


def _diagnose(args) -> Dict[str, Any]:
    from .settings import APP_ROOT
    from .storage import DEFAULT_STATE_FILE
    from .storage.diagnose import diagnose, interactive_login, to_markdown

    own_site_url = ''
    if args.db_file:
        with Database(args.db_file) as db:
            own_site_url = db.get_config('sharepoint', 'site_url') or ''
    state_file = Path(args.state_file) if args.state_file else DEFAULT_STATE_FILE
    if args.interactive_login:
        # 通常のセッションとは別のファイルに保存する（自社テナント用のセッションを壊さない）
        state_file = APP_ROOT / 'assets' / 'ms365_diagnose_state.json'
        interactive_login(args.target, state_file, channel=args.channel)
    report = diagnose(args.target, state_file, own_site_url)
    for finding in report['findings']:
        logger.info(f"見立て: {finding}")
    report_path = Path(args.report) if args.report else APP_ROOT / 'assets' / 'sharepoint-diagnose.md'
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(to_markdown(report), encoding='utf-8')
    return {'outputs': [str(report_path)], 'data': report}


def _sample(args) -> Dict[str, Any]:
    from .sample import create_sample_environment

    paths = create_sample_environment(Path(args.dest), force=args.force)
    return {'outputs': [str(p) for p in paths]}


# ---------- db（設定画面用） ----------

def _opt_bool(value: Optional[str], default: bool = True) -> bool:
    return default if value is None else str(value).lower() in ('1', 'true', 'yes')


DB_ACTIONS: Dict[str, Callable[[Database, argparse.Namespace], Any]] = {
    'list-projects': lambda db, a: {'projects': (p := db.get_projects(_opt_bool(a.active_only))), 'count': len(p)},
    'list-members': lambda db, a: {'members': (m := db.get_members()), 'count': len(m)},
    'update-member': lambda db, a: db.update_member(
        _required(a.member_id, 'member_id'),
        full_name=a.full_name, display_name=a.display_name, group_name=a.group_name,
        organization=a.organization, abbreviation=a.abbreviation, member_no=a.member_no,
        file_storage_location=a.file_storage_location,
        is_proprietary=int(a.is_proprietary) if a.is_proprietary is not None else None,
    ),
    'get-project-members': lambda db, a: {
        'project_id': a.project_id,
        'members': (m := db.get_project_members(_required(a.project_id, 'project_id'), a.as_of_date or _today())),
        'count': len(m),
    },
    'list-holidays': lambda db, a: {
        'project_id': a.project_id,
        'holidays': (h := db.get_holidays(a.project_id, a.start_date, a.end_date)),
        'count': len(h),
    },
    'add-holiday': lambda db, a: {'holiday_id': db.add_holiday(
        a.name, a.category, a.start_date, a.end_date, a.project_id, a.notes)},
    'update-holiday': lambda db, a: db.update_holiday(
        int(_required(a.holiday_id, 'holiday_id')), name=a.name, category=a.category,
        start_date=a.start_date, end_date=a.end_date, project_id=a.project_id, notes=a.notes),
    'delete-holiday': lambda db, a: db.delete_holiday(int(_required(a.holiday_id, 'holiday_id'))),
    'get-tool-settings': lambda db, a: db.get_tool_settings(),
    'update-tool-settings': lambda db, a: db.update_tool_settings(
        tool_root_dir=a.tool_root_dir, template_base_dir=a.template_base_dir,
        output_base_dir=a.output_base_dir, schedule_base_dir=a.schedule_base_dir),
    'get-config': lambda db, a: {
        'value': db.get_config(_required(a.tool_name, 'tool_name'), _required(a.config_key, 'config_key'),
                               a.project_id),
        'tool_name': a.tool_name, 'config_key': a.config_key, 'project_id': a.project_id,
    },
    'set-config': lambda db, a: db.set_config(
        _required(a.tool_name, 'tool_name'), _required(a.config_key, 'config_key'),
        a.config_value if a.config_value is not None else '', a.project_id, a.data_type or 'string'),
}


def _required(value, name: str):
    if value in (None, ''):
        raise ValueError(f"{name} は必須です")
    return value


def _today() -> str:
    from datetime import date
    return date.today().isoformat()


def _db(args) -> Dict[str, Any]:
    with Database(args.db_file) as db:
        data = DB_ACTIONS[args.action](db, args)
    return {'data': data}


# ---------- 引数 ----------

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog='yojitsu', description='PJメンバ予実管理ツール')
    sub = parser.add_subparsers(dest='command', required=True)

    run = sub.add_parser('run', help='帳票を作成する')
    run.add_argument('task', choices=TASKS)
    run.add_argument('--db-file', required=True)
    run.add_argument('--project-id')
    run.add_argument('--year-month')
    run.add_argument('--start-date')
    run.add_argument('--end-date')
    run.add_argument('--writer', choices=('auto', 'com', 'xml'), help='作業実績表の書き込み方式')
    run.set_defaults(handler=_run_task)

    sync = sub.add_parser('sync', help='リモートの個別予定を個別予定Dirへ取得する')
    sync.add_argument('--db-file', required=True)
    sync.set_defaults(handler=_sync)

    download = sub.add_parser('download', help='リモートのファイルを1件取得する')
    download.add_argument('--db-file', required=True)
    download.add_argument('--remote-path', required=True)
    download.add_argument('--download-dir', required=True)
    download.set_defaults(handler=_download)

    diag = sub.add_parser('diagnose-sharepoint', help='SharePoint 接続の診断（他社テナントへのゲスト参加など）')
    diag.add_argument('--target', required=True, help='取得できないファイル／フォルダの https URL')
    diag.add_argument('--db-file', help='自社テナントを見分けるため sharepoint.site_url を読む（任意）')
    diag.add_argument('--state-file', help='診断に使うセッション（既定: assets/ms365_storage_state.json）')
    diag.add_argument('--interactive-login', action='store_true',
                      help='ブラウザで接続先にログインし直したセッションで診断する（別ファイルに保存）')
    diag.add_argument('--channel', default='msedge', help='--interactive-login で使うブラウザ')
    diag.add_argument('--report', help='診断結果の Markdown の保存先（既定: assets/sharepoint-diagnose.md）')
    diag.set_defaults(handler=_diagnose)

    sample = sub.add_parser('sample', help='架空データでテスト環境を作る')
    sample.add_argument('--dest', required=True)
    sample.add_argument('--force', action='store_true', help='既存の dest を作り直す')
    sample.set_defaults(handler=_sample)

    db = sub.add_parser('db', help='設定画面用の DB 操作')
    db.add_argument('--db-file', required=True)
    db.add_argument('--action', required=True, choices=sorted(DB_ACTIONS))
    for name in ('project-id', 'tool-name', 'config-key', 'config-value', 'data-type',
                 'tool-root-dir', 'template-base-dir', 'output-base-dir', 'schedule-base-dir',
                 'as-of-date', 'start-date', 'end-date', 'active-only',
                 'member-id', 'full-name', 'display-name', 'group-name', 'organization', 'abbreviation',
                 'member-no', 'file-storage-location', 'is-proprietary',
                 'holiday-id', 'name', 'category', 'notes'):
        db.add_argument(f'--{name}')
    db.set_defaults(handler=_db)
    return parser


def main(argv=None) -> int:
    logging.basicConfig(level=logging.INFO, stream=sys.stderr,
                        format='%(asctime)s %(levelname)s %(message)s', datefmt='%H:%M:%S')
    # Electron 経由でも Windows コンソールでも文字化けしないよう UTF-8 に固定
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, 'reconfigure'):
            stream.reconfigure(encoding='utf-8')

    args = build_parser().parse_args(argv)
    try:
        result = {'success': True, 'outputs': [], 'warnings': []}
        result.update(args.handler(args) or {})
        code = 0
    except Exception as e:
        logger.exception(f"失敗: {e}")
        result = {'success': False, 'error': str(e), 'outputs': [], 'warnings': []}
        code = 1
    json.dump(result, sys.stdout, ensure_ascii=False, default=str)
    sys.stdout.write('\n')
    return code
