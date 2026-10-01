"""ストレージ共通のインターフェースと同期処理"""

import logging
import re
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from typing import Iterable, List, Protocol, Set

logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class RemoteFile:
    path: str   # server-relative path（例: /sites/xxx/Shared Documents/実績/01_<姓>.xlsx）

    @property
    def name(self) -> str:
        return PurePosixPath(self.path).name


class RemoteStorage(Protocol):
    def list_files(self, remote_dir: str, skip_top_level: Set[str] = frozenset()) -> List[RemoteFile]:
        """remote_dir 配下を再帰的に列挙する。
        skip_top_level に一致する直下のフォルダと、名前（拡張子あり/なし）が一致するファイルは除外"""
        ...

    def download(self, remote: RemoteFile, local_path: Path) -> None: ...

    def close(self) -> None: ...


def sanitize(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', '_', name).strip() or 'downloaded_file'


def local_path_for(remote_dir: str, remote: RemoteFile, local_root: Path) -> Path:
    root = PurePosixPath(remote_dir.rstrip('/'))
    path = PurePosixPath(remote.path)
    try:
        parts = path.relative_to(root).parts
    except ValueError:
        parts = (path.name,)
    return local_root.joinpath(*(sanitize(p) for p in parts))


def sync_directory(storage: RemoteStorage, remote_dir: str, local_dir: Path,
                   exclude: Iterable[str] = ()) -> List[Path]:
    """remote_dir 配下を local_dir に丸ごと取得する（同名ファイルは上書き）"""
    excluded = {e for e in exclude if e}
    files = storage.list_files(remote_dir, skip_top_level=excluded)
    logger.info(f"取得対象: {len(files)}件 (除外キー {len(excluded)}件)")
    saved = []
    for remote in files:
        target = local_path_for(remote_dir, remote, Path(local_dir))
        target.parent.mkdir(parents=True, exist_ok=True)
        storage.download(remote, target)
        logger.info(f"取得: {target}")
        saved.append(target)
    return saved
