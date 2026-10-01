"""SharePoint 実装

認証は Electron 側の手動ログイン（MFA 含む）で保存した Cookie（Playwright の storage_state 形式）を
そのまま使う。パスワードの保存や自動入力はしない。
API 呼び出しは Playwright の APIRequestContext で行うため、ブラウザ画面は起動しない。
"""

import logging
import re
from html import unescape
from pathlib import Path
from typing import List, Set
from urllib.parse import parse_qs, unquote, urljoin, urlparse, urlunparse

from .base import RemoteFile

logger = logging.getLogger(__name__)

SESSION_EXPIRED = 'SharePoint のセッションが切れています。基本設定から「ログインセッション初期化」を実行してください'


def to_server_relative(value: str) -> str:
    """server-relative path（/sites/...）か、id= / RootFolder= を含むフォルダURLを受け付ける"""
    raw = value.strip()
    if raw.startswith('/'):
        return raw.rstrip('/')
    parsed = urlparse(raw)
    for key in ('id', 'RootFolder'):
        for candidate in parse_qs(parsed.query).get(key, []):
            candidate = unquote(candidate).strip()
            if candidate.startswith('/'):
                return candidate.rstrip('/')
    path = unquote(parsed.path or '')
    if path.startswith('/sites/') and '/:' not in path:
        return path.rstrip('/')
    raise ValueError(
        f"リモートフォルダは server-relative path（/sites/...）で指定してください（共有用の短縮URLは不可）: {value}")


class SharePointStorage:
    def __init__(self, site_url: str, state_file: Path):
        if not state_file.exists():
            raise FileNotFoundError(f"ログインセッションがありません。基本設定から初期化してください: {state_file}")
        try:
            from playwright.sync_api import sync_playwright
        except ImportError as e:
            raise ImportError('playwright が未インストールです（pip install playwright）') from e

        self.site_url = site_url.rstrip('/')
        parsed = urlparse(self.site_url)
        self.origin = f"{parsed.scheme}://{parsed.netloc}"
        self._playwright = sync_playwright().start()
        self._request = self._playwright.request.new_context(storage_state=str(state_file))

    def close(self) -> None:
        self._request.dispose()
        self._playwright.stop()

    def _get_json(self, url: str) -> dict:
        response = self._request.get(url, headers={'accept': 'application/json;odata=nometadata'})
        if response.status in (401, 403):
            raise PermissionError(SESSION_EXPIRED)
        if response.status != 200:
            raise RuntimeError(f"SharePoint API エラー status={response.status} url={url}\n{response.text()[:300]}")
        return response.json()

    def _children(self, folder: str, kind: str) -> List[dict]:
        escaped = folder.replace("'", "''")
        # URL エンコードは Playwright に任せる（旧実装と同じ。実環境で動作確認済みの形）
        url = (f"{self.site_url}/_api/web/GetFolderByServerRelativePath(decodedurl='{escaped}')"
               f"/{kind}?$select=Name,ServerRelativeUrl")
        return self._get_json(url).get('value') or []

    def list_files(self, remote_dir: str, skip_top_level: Set[str] = frozenset()) -> List[RemoteFile]:
        root = to_server_relative(remote_dir)
        files: List[RemoteFile] = []
        stack = [root]
        while stack:
            folder = stack.pop()
            for item in self._children(folder, 'Files'):
                name = item.get('Name') or ''
                if name in skip_top_level or Path(name).stem in skip_top_level:
                    logger.info(f"除外: {name}")
                    continue
                files.append(RemoteFile(item['ServerRelativeUrl']))
            for item in self._children(folder, 'Folders'):
                name = item.get('Name') or ''
                if name.lower() == 'forms' or (folder == root and name in skip_top_level):
                    continue
                stack.append(item['ServerRelativeUrl'])
        return files

    def download(self, remote: RemoteFile, local_path: Path) -> None:
        url = urlunparse(urlparse(self.origin + remote.path)._replace(query='download=1'))
        response = self._follow_auto_post(self._request.get(url), url)
        content_type = (response.headers.get('content-type') or '').lower()
        if response.status in (401, 403) or 'text/html' in content_type:
            raise PermissionError(SESSION_EXPIRED)
        if response.status != 200:
            raise RuntimeError(f"ダウンロード失敗 status={response.status} path={remote.path}")
        local_path.write_bytes(response.body())

    def _follow_auto_post(self, response, url: str, max_hops: int = 3):
        """認証の中継ページ（自動 POST フォーム）を辿る"""
        for _ in range(max_hops):
            if 'text/html' not in (response.headers.get('content-type') or '').lower():
                return response
            body = response.text()
            if '<title>Working...</title>' not in body and '_forms/default.aspx' not in body:
                return response
            action = re.search(r"<form[^>]*action=[\"']([^\"']+)[\"']", body, re.IGNORECASE)
            fields = re.findall(
                r"<input[^>]*type=[\"']hidden[\"'][^>]*name=[\"']([^\"']+)[\"'][^>]*value=[\"']([^\"']*)[\"']",
                body, re.IGNORECASE)
            if not action or not fields:
                return response
            response = self._request.post(urljoin(url, unescape(action.group(1))),
                                          form={unescape(k): unescape(v) for k, v in fields})
        return response
