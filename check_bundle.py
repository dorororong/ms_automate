"""빌드된 exe 에 필요한 모듈이 다 들어갔는지 검사한다.

    python -m PyInstaller ms_automate.spec
    python check_bundle.py

`file_input.py` 의 PDF·HWP·OCR 의존성은 함수 안에서 지연 임포트된다. exe 를
직접 실행해 보는 것만으로는 파일을 넣기 전까지 빠진 게 드러나지 않으므로,
PyInstaller 가 남긴 의존성 목록(TOC)을 직접 확인한다.
"""

from __future__ import annotations

from pathlib import Path
import re
import sys

TOC = Path(__file__).resolve().parent / "build" / "ms_automate" / "Analysis-00.toc"

#: app.py 에서 출발해 반드시 묶여야 하는 앱 모듈.
APP_MODULES = (
    "app", "review", "cards", "service", "models", "storage", "theme", "paths",
    "icons", "tray", "hotkey", "selection", "windows_drop", "file_input",
    "classifier", "compact_extraction", "upstage_classifier",
    "outlook_adapter", "outlook_mirror",
)

#: 지연 임포트라 빠지기 쉬운 것들. 없으면 해당 기능만 조용히 죽는다.
REQUIRED_PACKAGES = (
    "pdfplumber",   # PDF 텍스트
    "pdfminer",     # pdfplumber 내부
    "olefile",      # HWP
    "pytesseract",  # 이미지 OCR
    "PIL",          # 아이콘·이미지
    "openai",       # Upstage 호출
    "win32com",     # Outlook COM
    "tkinterdnd2",  # 드래그앤드롭
)

#: 확장 모듈이라 이름이 `win32/xxx.pyd` 형태로 들어간다.
REQUIRED_BINARIES = ("win32gui", "win32api", "pythoncom")

#: 앱이 쓰지 않는데 딸려 오면 exe 만 커지는 것들.
UNWANTED = ("PySide6", "PyQt5", "pygame", "pygments", "jedi", "numpy", "Pythonwin")


def main() -> int:
    if not TOC.is_file():
        print(f"TOC 를 찾을 수 없습니다: {TOC}")
        print("먼저 `python -m PyInstaller ms_automate.spec` 을 실행하세요.")
        return 2

    text = TOC.read_text(encoding="utf-8", errors="replace")
    names = set(re.findall(r"\('([^']+)',", text))
    flat = " ".join(names)
    ok = True

    def report(label: str, good: bool, detail: str = "") -> None:
        nonlocal ok
        print(f"  {'OK  ' if good else 'FAIL'} {label}" + (f" - {detail}" if detail else ""))
        ok = ok and good

    print("[1] 앱 모듈")
    missing = [name for name in APP_MODULES if name not in names]
    report(f"{len(APP_MODULES)}개 모듈", not missing, "빠짐: " + ", ".join(missing) if missing else "")

    print("\n[2] 지연 임포트 의존성")
    for package in REQUIRED_PACKAGES:
        hits = [n for n in names if n == package or n.startswith(package + ".")]
        report(package, bool(hits), f"{len(hits)}개 항목" if hits else "번들에 없음")

    print("\n[3] 확장 모듈")
    for binary in REQUIRED_BINARIES:
        report(binary, binary in flat)

    print("\n[4] 제외되어야 하는 것")
    for package in UNWANTED:
        hits = [n for n in names if n == package or n.startswith(package + ".")]
        report(package, not hits, f"{len(hits)}개 항목이 남아 있음" if hits else "제외됨")

    print("\n결과:", "전부 통과" if ok else "실패 있음")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
