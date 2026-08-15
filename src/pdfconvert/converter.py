from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Callable, Iterable

from pypdf import PdfWriter

SUPPORTED_EXTENSIONS = {".pptx", ".docx"}
ProgressCallback = Callable[[int, int, str], None]


class ConversionError(RuntimeError):
    """使用者可理解的轉檔錯誤。"""


def validate_inputs(files: Iterable[Path]) -> list[Path]:
    paths = [Path(file).resolve() for file in files]
    if not paths:
        raise ConversionError("請先加入至少一個 PPTX 或 DOCX 檔案。")
    for path in paths:
        if not path.is_file():
            raise ConversionError(f"找不到檔案：{path}")
        if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            raise ConversionError(f"不支援的格式：{path.name}")
    return paths


def find_soffice() -> Path | None:
    executable = shutil.which("soffice") or shutil.which("libreoffice")
    if executable:
        return Path(executable)
    if os.name == "nt":
        candidates = [
            Path(os.environ.get("PROGRAMFILES", "")) / "LibreOffice/program/soffice.exe",
            Path(os.environ.get("PROGRAMFILES(X86)", "")) / "LibreOffice/program/soffice.exe",
        ]
        return next((path for path in candidates if path.is_file()), None)
    return None


def _convert_with_libreoffice(source: Path, output_dir: Path, soffice: Path) -> Path:
    profile = output_dir / f"lo-profile-{source.stem}"
    command = [
        str(soffice),
        "--headless",
        f"-env:UserInstallation={profile.resolve().as_uri()}",
        "--convert-to",
        "pdf",
        "--outdir",
        str(output_dir),
        str(source),
    ]
    result = subprocess.run(command, capture_output=True, text=True, timeout=180)
    output = output_dir / f"{source.stem}.pdf"
    if result.returncode != 0 or not output.is_file():
        detail = (result.stderr or result.stdout).strip()
        raise ConversionError(f"LibreOffice 無法轉換 {source.name}。{detail}")
    return output


def _convert_with_office(source: Path, output_dir: Path) -> Path:
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise ConversionError("找不到 LibreOffice，且無法啟用 Microsoft Office 轉檔元件。") from exc

    output = output_dir / f"{source.stem}.pdf"
    pythoncom.CoInitialize()
    app = None
    document = None
    try:
        if source.suffix.lower() == ".docx":
            app = win32com.client.DispatchEx("Word.Application")
            app.Visible = False
            document = app.Documents.Open(str(source), ReadOnly=True)
            document.ExportAsFixedFormat(str(output), 17)
        else:
            app = win32com.client.DispatchEx("PowerPoint.Application")
            document = app.Presentations.Open(str(source), WithWindow=False)
            document.SaveAs(str(output), 32)
    except Exception as exc:
        raise ConversionError(f"Microsoft Office 無法轉換 {source.name}：{exc}") from exc
    finally:
        if document is not None:
            document.Close()
        if app is not None:
            app.Quit()
        pythoncom.CoUninitialize()
    if not output.is_file():
        raise ConversionError(f"轉檔完成後找不到輸出：{source.name}")
    return output


def convert_and_merge(
    files: Iterable[Path], output: Path, progress: ProgressCallback | None = None
) -> Path:
    sources = validate_inputs(files)
    output = Path(output).resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    soffice = find_soffice()
    if not soffice and os.name != "nt":
        raise ConversionError("找不到 LibreOffice。請先安裝 LibreOffice，再重新啟動程式。")

    with tempfile.TemporaryDirectory(prefix="pdfconvert-") as temp_name:
        temp_dir = Path(temp_name)
        converted: list[Path] = []
        for index, source in enumerate(sources, start=1):
            if progress:
                progress(index - 1, len(sources), f"正在轉換：{source.name}")
            item_dir = temp_dir / str(index)
            item_dir.mkdir()
            converted.append(
                _convert_with_libreoffice(source, item_dir, soffice)
                if soffice
                else _convert_with_office(source, item_dir)
            )

        if progress:
            progress(len(sources), len(sources), "正在合併 PDF…")
        writer = PdfWriter()
        try:
            for pdf in converted:
                writer.append(str(pdf))
            with output.open("wb") as stream:
                writer.write(stream)
        finally:
            writer.close()
    return output

