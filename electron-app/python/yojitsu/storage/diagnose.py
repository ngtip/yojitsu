"""SharePoint 接続の診断（ゲスト参加している他社テナントなど）

ダウンロードが失敗する原因を切り分けるための情報を集める。出力はそのまま持ち出せるよう、
テナント名・サイト名・人名・Cookie の値は含めない（ホストは役割、パスは種類に置き換える）。

想定する原因:
  1. 保存したセッションに接続先テナントの認証 Cookie が無い（ログイン判定が早すぎた等）… コードで直せる
  2. 接続先テナントのアクセス制限（条件付きアクセス、管理外デバイスのダウンロード禁止など）
     … 相手側の方針なので回避しない
  3. 共有リンク（短縮URL）経由でしか権限が無い
"""

import logging
import re
import time
from html import unescape
from pathlib import Path
from typing import Dict, List, Optional
from urllib.parse import urljoin, urlparse, urlunparse

logger = logging.getLogger(__name__)

LOGIN_HOSTS = ('login.microsoftonline.com', 'login.live.com', 'login.windows.net', 'login.microsoft.com')
SP_AUTH_COOKIES = ('fedauth', 'rtfa', 'spoidcrl')
LOGIN_AUTH_COOKIES = ('estsauth', 'estsauthpersistent', 'estsauthlight', 'buid', 'fpc')
# 理由が書かれていて値を出してよい応答ヘッダ
SAFE_HEADERS = ('x-forms_based_auth_required', 'x-msdavext_error', 'x-ms-forbidreason', 'sprequestguid',
                'x-ms-diagnostics')
MAX_HOPS = 12

PAGE_PATTERNS = [
    ('blocked_by_policy', ('ConditionalAccess', 'UnmanagedDevice', 'BlockDownload', 'limited access',
                           'download is blocked', 'ダウンロードがブロック', '制限付きアクセス', 'アクセスがブロック',
                           "can't access this", 'Your sign-in was successful but', 'デバイスが準拠')),
    ('access_denied', ('AccessDenied.aspx', 'Access denied', 'アクセスが拒否', 'アクセス権がありません',
                       'need permission', 'アクセス許可が必要')),
    ('auto_post', ('<title>Working...</title>', '_forms/default.aspx')),
    ('login', ('Sign in to your account', 'アカウントにサインイン', 'loginfmt', 'login.microsoftonline.com/')),
]


def host_role(host: str, target_host: str, own_host: str = '') -> str:
    host = (host or '').lower()
    if host == target_host:
        return '接続先テナント'
    if own_host and host == own_host:
        return '自社テナント'
    if host.endswith(LOGIN_HOSTS):
        return 'Microsoftログイン'
    if host.endswith('.sharepoint.com'):
        return '別のSharePointテナント'
    return 'その他'


def path_kind(path: str) -> str:
    lower = path.lower()
    for marker, kind in (('/_layouts/15/accessdenied.aspx', 'アクセス拒否ページ'),
                         ('/_layouts/15/authenticate.aspx', '認証ページ'),
                         ('/_layouts/15/guestaccess.aspx', 'ゲスト用リンク'),
                         ('/_layouts/15/download.aspx', 'ダウンロード'),
                         ('/_forms/default.aspx', '認証の中継ページ'),
                         ('/_api/', 'REST API'),
                         ('/oauth2/', 'OAuth 認可'),
                         ('/kmsi', 'サインイン状態の維持'),
                         ('/login', 'ログイン')):
        if marker in lower:
            return kind
    if re.match(r'^/:[a-z]:/', lower):
        return '共有リンク（短縮URL）'
    if re.match(r'^/(sites|teams|personal)/', lower):
        return 'サイト内のパス'
    return 'その他'


def classify_html(html: str) -> str:
    for kind, needles in PAGE_PATTERNS:
        if any(needle.lower() in html.lower() for needle in needles):
            return kind
    return 'sharepoint_page'


def summarize_cookies(state: dict, target_host: str, own_host: str = '') -> List[dict]:
    """ドメインの役割ごとに、認証に関わる Cookie の名前と期限切れの数だけをまとめる"""
    now = time.time()
    groups: Dict[str, dict] = {}
    for cookie in state.get('cookies', []):
        role = host_role(str(cookie.get('domain', '')).lstrip('.'), target_host, own_host)
        name = str(cookie.get('name', ''))
        group = groups.setdefault(role, {'role': role, 'auth_cookies': set(), 'other_count': 0, 'expired': 0})
        if name.lower() in SP_AUTH_COOKIES + LOGIN_AUTH_COOKIES:
            group['auth_cookies'].add(name)
        else:
            group['other_count'] += 1
        expires = cookie.get('expires', -1)
        if isinstance(expires, (int, float)) and 0 < expires < now:
            group['expired'] += 1
    return [{**g, 'auth_cookies': sorted(g['auth_cookies'])} for g in groups.values()]


def site_root(path: str) -> Optional[str]:
    match = re.match(r'^(/(?:sites|teams)/[^/]+)', path, re.IGNORECASE)
    return match.group(1) if match else None


def _hop(response, url: str, target_host: str, own_host: str) -> dict:
    parsed = urlparse(url)
    headers = {k.lower(): v for k, v in response.headers.items()}
    return {
        'host': host_role(parsed.hostname or '', target_host, own_host),
        'path': path_kind(parsed.path),
        'status': response.status,
        'content_type': (headers.get('content-type') or '').split(';')[0],
        'headers': {k: headers[k][:200] for k in SAFE_HEADERS if k in headers},
        'www_authenticate': 'www-authenticate' in headers,
    }


def probe(request, url: str, target_host: str, own_host: str = '') -> List[dict]:
    """リダイレクトを1段ずつ辿り、各段階を記録する（応答本文は保存しない）"""
    hops = []
    method, form = 'GET', None
    for _ in range(MAX_HOPS):
        if method == 'GET':
            response = request.get(url, max_redirects=0, fail_on_status_code=False)
        else:
            response = request.post(url, form=form, max_redirects=0, fail_on_status_code=False)
        hop = _hop(response, url, target_host, own_host)
        hops.append(hop)
        location = response.headers.get('location')
        if 300 <= response.status < 400 and location:
            url, method, form = urljoin(url, location), 'GET', None
            continue
        if hop['content_type'] == 'text/html':
            body = response.text()
            hop['page'] = classify_html(body)
            if hop['page'] == 'auto_post':
                action = re.search(r"<form[^>]*action=[\"']([^\"']+)[\"']", body, re.IGNORECASE)
                fields = re.findall(
                    r"<input[^>]*type=[\"']hidden[\"'][^>]*name=[\"']([^\"']+)[\"'][^>]*value=[\"']([^\"']*)[\"']",
                    body, re.IGNORECASE)
                if action and fields:
                    url, method = urljoin(url, unescape(action.group(1))), 'POST'
                    form = {unescape(k): unescape(v) for k, v in fields}
                    continue
        break
    return hops


def current_user(request, origin: str, site: str) -> dict:
    """接続先サイトでログイン済み利用者として認識されているか（ID とゲストか否かだけを取る）"""
    url = f"{origin}{site}/_api/web/currentuser?$select=Id,IsShareByEmailGuestUser"
    response = request.get(url, headers={'accept': 'application/json;odata=nometadata'},
                           max_redirects=0, fail_on_status_code=False)
    result = {'status': response.status}
    if response.status == 200:
        try:
            data = response.json()
            result['recognized'] = bool(data.get('Id'))
            result['share_by_email_guest'] = data.get('IsShareByEmailGuestUser')
        except ValueError:
            result['recognized'] = False
    return result


def verdict(cookies: List[dict], hops: List[dict], user: Optional[dict], target_path: str) -> List[str]:
    findings = []
    target = next((c for c in cookies if c['role'] == '接続先テナント'), None)
    target_auth = [n for n in (target or {}).get('auth_cookies', []) if n.lower() in ('fedauth', 'rtfa')]
    final = hops[-1] if hops else {}
    went_login = any(h['host'] == 'Microsoftログイン' for h in hops)
    pages = {h.get('page') for h in hops}

    if not target_auth:
        findings.append('原因1の可能性: 保存したセッションに接続先テナントの認証Cookie（FedAuth/rtFa）がありません。'
                        '接続先にログインし直したセッションで再診断してください（--interactive-login）。')
    if 'blocked_by_policy' in pages:
        findings.append('原因2の可能性: 接続先テナントのアクセス制限（条件付きアクセス／管理外デバイスの制限）と思われる'
                        'ページが返っています。相手先の方針なので回避はせず、相手先の管理者に相談してください。')
    if 'access_denied' in pages or final.get('status') == 403:
        findings.append('権限不足の可能性: アクセス拒否が返っています。ゲストとしての権限がこのファイル／フォルダに'
                        '付いているか、共有リンク経由でしか権限が無いかを確認してください。')
    if went_login and target_auth:
        findings.append('認証Cookieはあるのにログインへ戻されています。セッションの期限切れか、接続先テナントが'
                        '再認証（MFA等）を求めている可能性があります。')
    if path_kind(target_path) == '共有リンク（短縮URL）':
        findings.append('原因3の可能性: 対象が共有リンク（短縮URL）です。ゲストの受け入れが済んでいないと'
                        'ブラウザ以外からは取得できないことがあります。')
    if user and user.get('status') == 200 and user.get('recognized') and final.get('content_type') != 'text/html' \
            and final.get('status') == 200:
        findings.append('接続先でログイン済みと認識され、ファイルも取得できました。この状態のセッションなら取得できます'
                        '（原因1が濃厚。ログイン判定の修正で直せる見込み）。')
    if not findings:
        findings.append('明確な原因を特定できませんでした。経路（hops）を確認してください。')
    return findings


def interactive_login(target_url: str, state_out: Path, channel: str = 'msedge', timeout_s: int = 300) -> None:
    """画面付きのブラウザで接続先にログインしてもらい、接続先テナントの認証Cookieが付いたら保存する"""
    from playwright.sync_api import sync_playwright

    target_host = (urlparse(target_url).hostname or '').lower()
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(channel=channel, headless=False)
        try:
            context = browser.new_context()
            context.new_page().goto(target_url, wait_until='domcontentloaded')
            logger.info('ブラウザで接続先にログインしてください（MFA 含む）。完了を自動で検出します')
            deadline = time.time() + timeout_s
            while time.time() < deadline:
                if any(str(c.get('domain', '')).lstrip('.').lower() == target_host
                       and str(c.get('name', '')).lower() in ('fedauth', 'rtfa') for c in context.cookies()):
                    break
                time.sleep(1)
            else:
                raise TimeoutError('接続先テナントの認証Cookieを検出できませんでした（時間切れ）')
            state_out.parent.mkdir(parents=True, exist_ok=True)
            context.storage_state(path=str(state_out))
        finally:
            browser.close()


def diagnose(target: str, state_file: Path, own_site_url: str = '') -> dict:
    import json

    from playwright.sync_api import sync_playwright

    parsed = urlparse(target)
    if not parsed.scheme:
        raise ValueError('--target は https:// から始まるURLで指定してください（接続先テナントのホスト名が必要なため）')
    target_host = (parsed.hostname or '').lower()
    own_host = (urlparse(own_site_url).hostname or '').lower() if own_site_url else ''
    if not state_file.exists():
        raise FileNotFoundError(f"セッションファイルがありません: {state_file}")
    state = json.loads(state_file.read_text(encoding='utf-8'))
    cookies = summarize_cookies(state, target_host, own_host)

    # ファイルらしければダウンロード指定を付ける（フォルダや共有リンクはそのまま）
    url = target
    if re.search(r'\.[A-Za-z0-9]{2,5}$', parsed.path) and 'download=1' not in parsed.query:
        url = urlunparse(parsed._replace(query=(parsed.query + '&' if parsed.query else '') + 'download=1'))

    with sync_playwright() as playwright:
        request = playwright.request.new_context(storage_state=str(state_file))
        try:
            hops = probe(request, url, target_host, own_host)
            site = site_root(parsed.path)
            user = current_user(request, f"{parsed.scheme}://{parsed.netloc}", site) if site else None
        finally:
            request.dispose()

    return {
        'target_kind': path_kind(parsed.path),
        'cookies': cookies,
        'hops': hops,
        'current_user': user,
        'findings': verdict(cookies, hops, user, parsed.path),
    }


def to_markdown(report: dict) -> str:
    lines = ['# SharePoint 接続診断', '',
             '※ テナント名・サイト名・人名・Cookie の値は含めていません', '',
             f"- 対象の種類: {report['target_kind']}", '', '## 見立て', '']
    lines += [f"- {f}" for f in report['findings']]
    lines += ['', '## セッション内の Cookie（名前のみ）', '', '| ドメイン | 認証Cookie | その他の数 | 期限切れ |', '|---|---|---|---|']
    lines += [f"| {c['role']} | {', '.join(c['auth_cookies']) or '-'} | {c['other_count']} | {c['expired']} |"
              for c in report['cookies']]
    lines += ['', '## アクセスの経路', '', '| # | ホスト | パスの種類 | 状態 | 種類 | ページ | ヘッダ |', '|---|---|---|---|---|---|---|']
    for i, h in enumerate(report['hops'], start=1):
        headers = '; '.join(f"{k}={v}" for k, v in h['headers'].items())
        lines.append(f"| {i} | {h['host']} | {h['path']} | {h['status']} | {h['content_type']} | "
                     f"{h.get('page', '')} | {headers} |")
    user = report.get('current_user')
    lines += ['', '## 接続先サイトでの認識', '',
              f"- {user}" if user else '- サイトのパスが分からないため未確認']
    return '\n'.join(lines) + '\n'
