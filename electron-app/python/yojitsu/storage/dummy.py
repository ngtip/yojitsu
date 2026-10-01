"""ローカルフォルダを SharePoint に見立てるダミー実装"""

import shutil
from pathlib import Path, PurePosixPath
from typing import List, Set

from .base import RemoteFile


class DummyStorage:
    def __init__(self, root: Path):
        self.root = Path(root)

    def _local(self, remote_path: str) -> Path:
        parts = PurePosixPath('/' + remote_path.strip().lstrip('/')).parts[1:]
        if '..' in parts:
            raise ValueError(f"不正なリモートパスです: {remote_path}")
        return self.root.joinpath(*parts)

    def list_files(self, remote_dir: str, skip_top_level: Set[str] = frozenset()) -> List[RemoteFile]:
        base = self._local(remote_dir)
        if not base.is_dir():
            raise FileNotFoundError(f"リモートフォルダが見つかりません: {remote_dir}（ダミー: {base}）")
        prefix = '/' + remote_dir.strip().strip('/')
        files = []
        for path in sorted(base.rglob('*')):
            if not path.is_file():
                continue
            relative = path.relative_to(base)
            if relative.parts[0] in skip_top_level or path.name in skip_top_level or path.stem in skip_top_level:
                continue
            files.append(RemoteFile(f"{prefix}/{relative.as_posix()}"))
        return files

    def download(self, remote: RemoteFile, local_path: Path) -> None:
        source = self._local(remote.path)
        if not source.is_file():
            raise FileNotFoundError(f"リモートファイルが見つかりません: {remote.path}")
        shutil.copy2(source, local_path)

    def close(self) -> None:
        pass
