# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 빌드 정의.

    python -m PyInstaller ms_automate.spec

`dist/업무정리.exe` 하나가 나온다. 실행하면 exe 가 놓인 폴더에 `data/` 를 만들고,
같은 폴더의 `.env` 와 `profile.json` 을 읽는다 (paths.py 참고).
"""

from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files

PROJECT = Path(SPECPATH)

# 실행 파일 아이콘은 앱이 쓰는 그림에서 그대로 만든다.
import icons  # noqa: E402  (SPECPATH 가 sys.path 에 들어간 뒤에 임포트)

build_dir = PROJECT / "build"
build_dir.mkdir(exist_ok=True)
exe_icon = icons.build_ico(icons.ICON_SOURCE, build_dir / "app.ico")

datas = [
    # 창·트레이 아이콘 원본. paths.resource() 가 이 경로로 찾는다.
    (str(icons.ICON_SOURCE), "assets"),
]
# tkinterdnd2 는 tcl 확장(tkdnd)을 데이터로 들고 다닌다. 빠지면 드래그앤드롭이
# 조용히 죽고 앱은 네이티브 드롭으로만 동작한다.
datas += collect_data_files("tkinterdnd2")

a = Analysis(
    ["app.py"],
    pathex=[str(PROJECT)],
    binaries=[],
    datas=datas,
    hiddenimports=[
        "tkinterdnd2",
        "win32timezone",  # pywin32 COM 이 늦게 부른다
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # 앱이 쓰지 않는데 딸려 오는 것들. 빼면 exe 가 82MB → 절반 아래로 준다.
        # Pillow 의 ImageQt 가 Qt 바인딩을 찾아 들어오는 게 가장 크다(45MB).
        "PySide6", "PySide2", "PyQt5", "PyQt6", "shiboken6", "shiboken2",
        "PIL.ImageQt",
        "pygame",
        "Pythonwin",        # pywin32 의 IDE
        "pygments", "jedi", "parso",
        "IPython", "notebook",
        "matplotlib", "numpy", "pandas", "scipy",
        "pytest", "setuptools", "pip",
        "tkinter.test", "test",
    ],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="업무정리",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=False,          # 콘솔 창 없이 뜬다
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(exe_icon),
)
