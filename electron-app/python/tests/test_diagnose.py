"""SharePoint 接続診断。偽の SharePoint（ローカル HTTP サーバ）を相手に経路の記録と見立てを確かめる"""

import json
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer

import pytest

from yojitsu.storage.diagnose import classify_html, diagnose, host_role, path_kind, summarize_cookies, to_markdown

pytest.importorskip('playwright')


class FakeSharePoint(BaseHTTPRequestHandler):
    """Cookie と URL で振る舞いを変える。FedAuth=ok ならファイルを返す"""

    def log_message(self, *_args):
        pass

    def _send(self, status, body=b'', content_type='text/html', headers=None):
        self.send_response(status)
        self.send_header('Content-Type', content_type)
        for key, value in (headers or {}).items():
            self.send_header(key, value)
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        authed = 'FedAuth=ok' in (self.headers.get('Cookie') or '')
        if self.path.startswith('/sites/guest/_api/web/currentuser'):
            if authed:
                return self._send(200, json.dumps({'Id': 7, 'IsShareByEmailGuestUser': False}).encode(),
                                  'application/json')
            return self._send(401, b'', headers={'WWW-Authenticate': 'Bearer', 'X-Forms_Based_Auth_Required': 'x'})
        if self.path.startswith('/sites/guest/blocked/'):
            return self._send(200, '<title>アクセスがブロックされました</title>UnmanagedDevice'.encode())
        if self.path.startswith('/sites/guest/denied/'):
            return self._send(302, headers={'Location': '/sites/guest/_layouts/15/AccessDenied.aspx'})
        if self.path.startswith('/sites/guest/_layouts/15/AccessDenied.aspx'):
            return self._send(403, b'<title>Access denied</title>', headers={'SPRequestGuid': 'guid-1'})
        if self.path.startswith('/sites/guest/'):
            if authed:
                return self._send(200, b'PK\x03\x04', 'application/octet-stream')
            return self._send(302, headers={'Location': '/_forms/default.aspx?ReturnUrl=x'})
        if self.path.startswith('/_forms/default.aspx'):
            return self._send(200, b'<html><body>Sign in to your account <input name="loginfmt"></body></html>')
        return self._send(404)


@pytest.fixture(scope='module')
def server():
    httpd = HTTPServer(('127.0.0.1', 0), FakeSharePoint)
    thread = threading.Thread(target=httpd.serve_forever, daemon=True)
    thread.start()
    yield f"http://127.0.0.1:{httpd.server_address[1]}"
    httpd.shutdown()


def _state(tmp_path, cookies):
    path = tmp_path / 'state.json'
    path.write_text(json.dumps({'cookies': [
        {'name': n, 'value': v, 'domain': d, 'path': '/', 'expires': -1, 'httpOnly': True, 'secure': False,
         'sameSite': 'Lax'} for n, v, d in cookies], 'origins': []}), encoding='utf-8')
    return path


def test_missing_target_cookies_is_cause_1(server, tmp_path):
    # 自社テナントの Cookie しか無いセッション（旧版のログイン判定で起こる状態）
    state = _state(tmp_path, [('FedAuth', 'secret', 'own.sharepoint.com'), ('ESTSAUTH', 'secret', 'login.microsoftonline.com')])
    report = diagnose(f"{server}/sites/guest/Shared%20Documents/a.xlsx", state, 'https://own.sharepoint.com/sites/x')
    assert report['findings'][0].startswith('原因1の可能性')
    assert [h['path'] for h in report['hops']] == ['サイト内のパス', '認証の中継ページ']
    assert report['hops'][-1]['page'] == 'login'
    assert report['current_user'] == {'status': 401}
    roles = {c['role']: c['auth_cookies'] for c in report['cookies']}
    assert roles == {'自社テナント': ['FedAuth'], 'Microsoftログイン': ['ESTSAUTH']}
    assert 'secret' not in to_markdown(report) and 'own.sharepoint.com' not in to_markdown(report)


def test_working_session_points_to_cause_1_fix(server, tmp_path):
    state = _state(tmp_path, [('FedAuth', 'ok', '127.0.0.1')])
    report = diagnose(f"{server}/sites/guest/Shared%20Documents/a.xlsx", state)
    assert report['hops'][-1]['status'] == 200 and report['hops'][-1]['content_type'] == 'application/octet-stream'
    assert report['current_user'] == {'status': 200, 'recognized': True, 'share_by_email_guest': False}
    assert any('ログイン判定の修正で直せる' in f for f in report['findings'])


def test_policy_block_and_access_denied(server, tmp_path):
    state = _state(tmp_path, [('FedAuth', 'ok', '127.0.0.1')])
    blocked = diagnose(f"{server}/sites/guest/blocked/a.xlsx", state)
    assert any(f.startswith('原因2の可能性') for f in blocked['findings'])
    denied = diagnose(f"{server}/sites/guest/denied/a.xlsx", state)
    assert denied['hops'][-1]['path'] == 'アクセス拒否ページ' and denied['hops'][-1]['headers'] == {'sprequestguid': 'guid-1'}
    assert any(f.startswith('権限不足の可能性') for f in denied['findings'])


def test_classifiers():
    assert host_role('x.sharepoint.com', 'y.sharepoint.com') == '別のSharePointテナント'
    assert path_kind('/:x:/s/site/AbCd') == '共有リンク（短縮URL）'
    assert classify_html('<title>Working...</title>') == 'auto_post'
    # 普通の SharePoint ページに「サインイン」の文字があってもログイン画面とはみなさない
    assert classify_html('<a>別のユーザーとしてサインイン</a>') == 'sharepoint_page'
    expired = summarize_cookies({'cookies': [{'name': 'rtFa', 'domain': '.t.sharepoint.com', 'expires': 1}]},
                                't.sharepoint.com')
    assert expired == [{'role': '接続先テナント', 'auth_cookies': ['rtFa'], 'other_count': 0, 'expired': 1}]
