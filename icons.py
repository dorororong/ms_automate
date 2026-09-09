"""앱 아이콘. 창과 트레이가 같은 그림 하나를 쓴다.

원본은 `assets/tray_icon.png` 입니다. Windows 창과 알림 영역은 `.ico`
를 요구하므로 첫 실행 때 한 번 변환해 임시 폴더에 캐시합니다. 원본을 바꾸면
수정 시각이 달라져 자동으로 다시 만듭니다.

Pillow 나 원본이 없으면 `None` 을 돌려주고, 호출한 쪽은 예전처럼 기본 아이콘을
씁니다. 아이콘 때문에 앱이 뜨지 않는 일은 없어야 합니다.
"""

from __future__ import annotations

import ctypes
from pathlib import Path
import tempfile

import paths

try:
    from PIL import Image
except ImportError:  # pragma: no cover - Pillow 미설치 환경
    Image = None


ICON_SOURCE = paths.resource("assets", "tray_icon.png")

# 원본은 반투명한 회색 배경 위에 아이콘이 얹혀 있다. 이 값보다 옅은 픽셀은
# 배경으로 보고 완전히 지운다.
ALPHA_FLOOR = 40
ICON_SIZES = (16, 20, 24, 32, 40, 48, 64, 128, 256)
_cache: Path | None | str = "unset"


def _trimmed_square(image: "Image.Image") -> "Image.Image":
    """배경 안개를 지우고 아이콘만 정사각형으로 잘라 낸다."""

    image = image.convert("RGBA")
    alpha = image.getchannel("A")
    # 옅은 배경을 0으로 눌러야 bbox 가 그림 영역만 잡는다.
    solid = alpha.point(lambda value: 255 if value >= ALPHA_FLOOR else 0)
    image.putalpha(alpha.point(lambda value: value if value >= ALPHA_FLOOR else 0))

    box = solid.getbbox()
    if box:
        image = image.crop(box)

    side = max(image.width, image.height)
    canvas = Image.new("RGBA", (side, side), (0, 0, 0, 0))
    canvas.paste(
        image, ((side - image.width) // 2, (side - image.height) // 2), image
    )
    return canvas


def build_ico(source: Path, target: Path) -> Path:
    """PNG 한 장을 여러 크기가 담긴 .ico 로 만든다."""

    with Image.open(source) as raw:
        icon = _trimmed_square(raw)
    sizes = [(size, size) for size in ICON_SIZES if size <= max(icon.size)]
    target.parent.mkdir(parents=True, exist_ok=True)
    icon.save(target, format="ICO", sizes=sizes or [(32, 32)])
    return target


def ico_path() -> Path | None:
    """창·트레이가 쓸 .ico 경로. 만들지 못하면 None."""

    global _cache
    if _cache != "unset":
        return _cache  # type: ignore[return-value]

    _cache = None
    if Image is None or not ICON_SOURCE.exists():
        return None
    try:
        stamp = int(ICON_SOURCE.stat().st_mtime)
        target = Path(tempfile.gettempdir()) / f"ms_automate_icon_{stamp}.ico"
        if not target.exists():
            build_ico(ICON_SOURCE, target)
        _cache = target
    except (OSError, ValueError):
        _cache = None
    return _cache  # type: ignore[return-value]


APP_MODEL_ID = "dorororong.ms_automate"


def claim_taskbar_identity() -> None:
    """작업 표시줄에서 python.exe 가 아니라 이 앱으로 보이게 한다.

    이걸 부르지 않으면 Windows 가 창 아이콘 대신 파이썬 실행 파일 아이콘을
    씁니다. 창을 만들기 전에 한 번만 부르면 됩니다.
    """

    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_MODEL_ID)
    except (AttributeError, OSError):
        pass


def apply_window_icon(window) -> None:
    """Tk 창의 제목 표시줄·작업 표시줄 아이콘을 바꾼다."""

    path = ico_path()
    if path is None:
        return
    try:
        window.iconbitmap(default=str(path))
    except Exception:
        # 아이콘을 못 붙여도 창은 그대로 떠야 한다.
        pass
