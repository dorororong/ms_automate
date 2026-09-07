"""Local file-to-text input for the CATMOA-style registration flow.

This module deliberately handles only files the user drops or selects.  It
does not search CoolMessenger UDB files and it does not run a PII detector.
"""

from __future__ import annotations

from dataclasses import dataclass
import re
import struct
from pathlib import Path
import zipfile


TEXT_EXTENSIONS = {".txt", ".md", ".csv", ".tsv", ".log"}
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".webp", ".tiff"}
MAX_TEXT_CHARS = 120_000


class FileLoadError(RuntimeError):
    """Raised when a dropped file cannot be converted into text."""


@dataclass(frozen=True)
class FileLoadResult:
    text: str
    names: tuple[str, ...]
    warnings: tuple[str, ...] = ()


def load_files(paths: list[str] | tuple[str, ...]) -> FileLoadResult:
    """Read dropped files and combine them into one analysis input."""

    chunks: list[str] = []
    names: list[str] = []
    warnings: list[str] = []
    seen: set[str] = set()

    for raw_path in paths:
        path = Path(raw_path)
        key = str(path.resolve()).casefold()
        if key in seen:
            continue
        seen.add(key)

        if not path.is_file():
            warnings.append(f"{path.name}: 파일을 찾을 수 없습니다.")
            continue
        try:
            text = parse_file(path).strip()
        except FileLoadError as exc:
            warnings.append(f"{path.name}: {exc}")
            continue
        except Exception as exc:  # Keep one bad attachment from hiding others.
            warnings.append(f"{path.name}: 읽기 실패 ({exc})")
            continue

        if not text:
            warnings.append(f"{path.name}: 일정으로 읽을 본문이 없습니다.")
            continue
        if len(text) > MAX_TEXT_CHARS:
            text = text[:MAX_TEXT_CHARS].rstrip()
            warnings.append(f"{path.name}: 본문이 길어 {MAX_TEXT_CHARS:,}자까지만 사용했습니다.")
        names.append(path.name)
        chunks.append(f"[파일: {path.name}]\n{text}")

    if not chunks:
        detail = "\n".join(warnings) or "선택한 파일에서 읽을 본문을 찾지 못했습니다."
        raise FileLoadError(detail)
    return FileLoadResult("\n\n".join(chunks), tuple(names), tuple(warnings))


def parse_file(path: Path) -> str:
    """Parse one supported file into plain text."""

    suffix = path.suffix.lower()
    if suffix in TEXT_EXTENSIONS:
        return _read_text_file(path)
    if suffix == ".pdf":
        return _parse_pdf(path)
    if suffix == ".hwpx":
        return _parse_hwpx(path)
    if suffix == ".hwp":
        return _parse_hwp(path)
    if suffix == ".docx":
        return _parse_docx(path)
    if suffix in IMAGE_EXTENSIONS:
        return _parse_image(path)

    # Text-like exports often have a non-standard extension.  A best-effort
    # decode is friendlier than silently ignoring a file the user dropped.
    try:
        return _read_text_file(path)
    except Exception as exc:
        raise FileLoadError(f"지원하지 않는 파일 형식입니다: {suffix or '(확장자 없음)'}") from exc


def _read_text_file(path: Path) -> str:
    data = path.read_bytes()
    for encoding in ("utf-8-sig", "utf-8", "cp949", "euc-kr"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="ignore")


def _parse_pdf(path: Path) -> str:
    try:
        import pdfplumber
    except ImportError as exc:
        raise FileLoadError("PDF를 읽으려면 pdfplumber를 설치하세요.") from exc

    texts: list[str] = []
    try:
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    texts.append(page_text)
                for table in page.extract_tables() or []:
                    rows = [" | ".join(str(cell or "") for cell in row) for row in table]
                    texts.extend(row for row in rows if row.strip())
    except Exception as exc:
        raise FileLoadError(f"PDF 파싱 오류: {exc}") from exc
    return "\n".join(texts)


def _parse_hwpx(path: Path) -> str:
    sections: list[str] = []
    try:
        with zipfile.ZipFile(path, "r") as archive:
            names = sorted(
                name for name in archive.namelist()
                if re.fullmatch(r"Contents/section\d+\.xml", name, re.IGNORECASE)
            )
            if not names:
                names = [name for name in archive.namelist() if name.endswith("document.xml")]
            for name in names:
                xml = archive.read(name).decode("utf-8", errors="ignore")
                text = re.sub(r"<[^>]+>", " ", xml)
                text = re.sub(r"\s+", " ", text).strip()
                if text:
                    sections.append(text)
    except zipfile.BadZipFile as exc:
        raise FileLoadError("HWPX 파일이 손상되었거나 올바르지 않습니다.") from exc
    except Exception as exc:
        raise FileLoadError(f"HWPX 파싱 오류: {exc}") from exc
    return "\n".join(sections)


def _parse_docx(path: Path) -> str:
    try:
        with zipfile.ZipFile(path, "r") as archive:
            xml = archive.read("word/document.xml").decode("utf-8", errors="ignore")
    except (KeyError, zipfile.BadZipFile) as exc:
        raise FileLoadError("DOCX 본문을 읽을 수 없습니다.") from exc
    except Exception as exc:
        raise FileLoadError(f"DOCX 파싱 오류: {exc}") from exc
    text = re.sub(r"</w:p\s*>", "\n", xml, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", "", text)
    return re.sub(r"\s+", " ", text).strip()


def _parse_hwp(path: Path) -> str:
    try:
        import olefile
    except ImportError as exc:
        raise FileLoadError("HWP를 읽으려면 olefile을 설치하세요.") from exc

    try:
        if not olefile.isOleFile(path):
            raise FileLoadError("올바른 HWP 파일이 아닙니다.")
        with olefile.OleFileIO(path) as ole:
            streams = [entry for entry in ole.listdir() if "BodyText" in entry]
            if not streams:
                raise FileLoadError("HWP 본문을 찾을 수 없습니다.")
            texts: list[str] = []
            for stream in streams:
                try:
                    raw = ole.openstream(stream).read()
                    try:
                        import zlib
                        raw = zlib.decompress(raw, -15)
                    except Exception:
                        pass
                    text = _extract_hwp_text(raw)
                    if text:
                        texts.append(text)
                except Exception:
                    continue
    except FileLoadError:
        raise
    except Exception as exc:
        raise FileLoadError(f"HWP 파싱 오류: {exc}") from exc

    if not texts:
        raise FileLoadError("HWP 텍스트 추출에 실패했습니다. HWPX 또는 텍스트 붙여넣기를 사용하세요.")
    return "\n".join(texts)


def _extract_hwp_text(data: bytes) -> str:
    texts: list[str] = []
    offset = 0
    while offset + 4 <= len(data):
        header = struct.unpack_from("<I", data, offset)[0]
        record_type = header & 0x3FF
        record_size = (header >> 20) & 0xFFF
        offset += 4
        if record_size == 0xFFF and offset + 4 <= len(data):
            record_size = struct.unpack_from("<I", data, offset)[0]
            offset += 4
        if record_type == 67 and offset + record_size <= len(data):
            text = data[offset:offset + record_size].decode("utf-16-le", errors="ignore").strip()
            if text and not text.startswith("\x00"):
                texts.append(text)
        offset += record_size
    return " ".join(texts)


def _parse_image(path: Path) -> str:
    """Use optional local OCR; no image is sent anywhere by this module."""

    try:
        from PIL import Image
        import pytesseract
    except ImportError as exc:
        raise FileLoadError(
            "이미지는 현재 로컬 OCR이 필요합니다. Pillow와 pytesseract, "
            "Tesseract(한국어 언어팩)를 설치하거나 텍스트를 붙여넣어 주세요."
        ) from exc
    try:
        text = pytesseract.image_to_string(Image.open(path), lang="kor+eng")
    except Exception as exc:
        raise FileLoadError(f"이미지 OCR에 실패했습니다: {exc}") from exc
    if not text.strip():
        raise FileLoadError("이미지에서 읽은 텍스트가 없습니다.")
    return text.strip()
