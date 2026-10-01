import json
import sys
from datetime import date
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from yojitsu.cli import main as cli_main  # noqa: E402
from yojitsu.sample import create_sample_environment  # noqa: E402

BASE_DATE = date(2026, 10, 20)


@pytest.fixture
def sample_env(tmp_path, monkeypatch):
    dest = tmp_path / 'env'
    create_sample_environment(dest, base=BASE_DATE)
    monkeypatch.setenv('YOJITSU_SITE_SETTINGS', str(dest / 'site-settings.json'))
    monkeypatch.setenv('YOJITSU_SUBMIT_WRITER', 'xml')
    monkeypatch.delenv('YOJITSU_STORAGE_BACKEND', raising=False)
    return dest


@pytest.fixture
def db_file(sample_env):
    return str(sample_env / 'db' / 'management.sqlite')


@pytest.fixture
def cli(capsys):
    def run(*args):
        code = cli_main(list(args))
        out = capsys.readouterr().out
        result = json.loads(out)
        return code, result
    return run
