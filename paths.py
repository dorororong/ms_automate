"""소스 실행과 실행 파일(exe) 실행의 경로 차이를 한곳에서 흡수한다.

PyInstaller 로 묶으면 두 위치가 갈라진다.

- **번들 자원**: onefile exe 는 실행할 때마다 임시 폴더에 풀린다.
  아이콘 같은 읽기 전용 자원은 거기에 있다. 종료하면 사라진다.
- **사용자 데이터**: DB·`.env`·`profile.json` 은 다음 실행에도 남아야 하므로
  임시 폴더가 아니라 **exe 가 놓인 폴더**에 둔다.

이 구분을 모듈마다 하면 하나만 빠뜨려도 "등록은 되는데 다음에 켜면 기록이
사라지는" 식으로 조용히 망가진다. 그래서 여기서만 정한다.
"""

from __future__ import annotations

from pathlib import Path
import sys

#: PyInstaller 로 묶인 실행 파일에서 도는 중인가.
FROZEN = bool(getattr(sys, "frozen", False))

SOURCE_ROOT = Path(__file__).resolve().parent


def resource_dir() -> Path:
    """읽기 전용 번들 자원(아이콘 등)이 있는 폴더."""

    if FROZEN:
        # onefile 은 _MEIPASS 에 풀고, onedir 은 exe 옆에 둔다.
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
    return SOURCE_ROOT


def user_dir() -> Path:
    """실행 사이에 남아야 하는 사용자 파일이 있는 폴더."""

    if FROZEN:
        return Path(sys.executable).resolve().parent
    return SOURCE_ROOT


def resource(*parts: str) -> Path:
    return resource_dir().joinpath(*parts)


def user_file(*parts: str) -> Path:
    return user_dir().joinpath(*parts)
