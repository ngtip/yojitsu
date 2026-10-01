"""リモートストレージ（個別予定ファイルの置き場所）

backend:
  sharepoint ... 実環境。Electron の手動ログインで保存したセッションで SharePoint REST API を使う
  dummy      ... ローカルフォルダを SharePoint に見立てる。リモートパス /sites/... をそのまま
                 dummy_root 配下の相対パスとして扱うので、実環境と同じ設定値で動かせる

切替: 環境変数 YOJITSU_STORAGE_BACKEND > DB tool_config(storage, backend) > sharepoint
"""

import os
from pathlib import Path
from typing import Optional

from ..db import Database
from ..settings import APP_ROOT
from .base import RemoteFile, RemoteStorage, sync_directory

DEFAULT_STATE_FILE = APP_ROOT / 'assets' / 'ms365_storage_state.json'


def storage_setting(db: Database, key: str, env: Optional[str] = None) -> Optional[str]:
    return (os.environ.get(env) if env else None) or db.get_config('storage', key)


def remote_work_dir(db: Database) -> str:
    value = db.get_config('sharepoint', 'remote_work_dir') or ''
    if not value.strip():
        raise ValueError('リモートの作業実績格納先（sharepoint.remote_work_dir）が未設定です')
    return value.strip()


def create_storage(db: Database) -> RemoteStorage:
    backend = (storage_setting(db, 'backend', 'YOJITSU_STORAGE_BACKEND') or 'sharepoint').lower()
    if backend == 'dummy':
        from .dummy import DummyStorage

        root = storage_setting(db, 'dummy_root', 'YOJITSU_DUMMY_ROOT')
        if not root:
            raise ValueError('dummy ストレージのルート（storage.dummy_root）が未設定です')
        root_path = Path(root)
        if not root_path.is_absolute():
            root_path = APP_ROOT / root_path
        return DummyStorage(root_path)
    if backend == 'sharepoint':
        from .sharepoint import SharePointStorage

        site_url = os.environ.get('SHAREPOINT_SITE_URL') or db.get_config('sharepoint', 'site_url')
        if not site_url:
            raise ValueError('SharePoint サイトURL（sharepoint.site_url または SHAREPOINT_SITE_URL）が未設定です')
        state_file = Path(os.environ.get('PLAYWRIGHT_STATE_FILE') or DEFAULT_STATE_FILE)
        return SharePointStorage(site_url, state_file)
    raise ValueError(f"不明なストレージ backend です: {backend}")


__all__ = ['RemoteFile', 'RemoteStorage', 'create_storage', 'remote_work_dir', 'sync_directory']
