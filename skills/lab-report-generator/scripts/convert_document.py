#!/usr/bin/env python3
"""Convert a working copy using installed Word or LibreOffice; never overwrite inputs."""

import argparse
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
from zipfile import ZipFile


def find_soffice(explicit=None):
    if explicit:
        if not Path(explicit).is_file():
            raise ValueError(f"LibreOffice executable does not exist: {explicit}")
        return str(Path(explicit).resolve())
    candidate = shutil.which("soffice") or shutil.which("soffice.exe")
    if candidate:
        return candidate
    defaults = ([Path(os.environ.get("ProgramFiles", "C:/Program Files")) / "LibreOffice/program/soffice.exe"]
                if sys.platform == "win32" else [Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")])
    return next((str(p) for p in defaults if p.is_file()), None)


def word_available():
    if sys.platform != "win32":
        return False
    import winreg
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Word.Application\\CLSID"):
            return True
    except OSError:
        return False


def validate_output(path):
    if not path.is_file() or path.stat().st_size == 0:
        raise ValueError("Converter did not produce a non-empty output.")
    if path.suffix.lower() == ".docx":
        with ZipFile(path) as archive:
            if "word/document.xml" not in archive.namelist() or archive.testzip():
                raise ValueError("Converter produced an invalid DOCX package.")
    elif not path.read_bytes().startswith(b"%PDF-"):
        raise ValueError("Converter did not produce a PDF.")


def convert(source, output, engine="auto", soffice=None, timeout=60):
    source, output = Path(source).resolve(), Path(output).resolve()
    if not source.is_file():
        raise ValueError("Input document does not exist.")
    if output == source or output.exists():
        raise ValueError("Output must be a new path different from the input.")
    if (source.suffix.lower(), output.suffix.lower()) not in {
        (".doc", ".docx"), (".docx", ".pdf"), (".doc", ".pdf")
    }:
        raise ValueError("Supported conversions: DOC to DOCX, DOC/DOCX to PDF.")
    if timeout <= 0:
        raise ValueError("Timeout must be positive.")
    if engine == "auto":
        engine = "word" if word_available() else "libreoffice"
    if engine == "word" and not word_available():
        raise ValueError("Microsoft Word COM is unavailable. Install Word or choose LibreOffice.")
    executable = find_soffice(soffice) if engine == "libreoffice" else None
    if engine == "libreoffice" and not executable:
        raise ValueError("LibreOffice is unavailable. Install it or provide --soffice; no file was converted.")
    before = hashlib.sha256(source.read_bytes()).hexdigest()
    with tempfile.TemporaryDirectory(prefix="lab-report-convert-") as directory:
        task = Path(directory)
        copied = task / ("source" + source.suffix.lower())
        shutil.copy2(source, copied)
        if engine == "word":
            converted = task / ("converted" + output.suffix.lower())
            command = ["powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                       "-File", str(Path(__file__).with_name("word_convert.ps1")),
                       "-Source", str(copied), "-Destination", str(converted)]
        else:
            target = task / "converted"
            target.mkdir()
            profile = (task / "profile").as_uri()
            format_filter = "docx:Office Open XML Text" if output.suffix.lower() == ".docx" else "pdf:writer_pdf_Export"
            command = [executable, f"-env:UserInstallation={profile}", "--headless", "--nologo",
                       "--nodefault", "--norestore", "--convert-to", format_filter,
                       "--outdir", str(target), str(copied)]
            converted = target / ("source" + output.suffix.lower())
        try:
            process = subprocess.run(command, capture_output=True, text=True, timeout=timeout)
        except subprocess.TimeoutExpired as error:
            raise ValueError("Conversion timed out; no final output was written. Check the converter instance before retrying.") from error
        if process.returncode:
            detail = (process.stderr or process.stdout).strip()
            raise ValueError(f"{engine} conversion failed ({process.returncode}): {detail[-1500:]}")
        validate_output(converted)
        if hashlib.sha256(source.read_bytes()).hexdigest() != before:
            raise ValueError("Input changed during conversion; final output was not written.")
        output.parent.mkdir(parents=True, exist_ok=True)
        created = False
        try:
            with output.open("xb") as stream:
                created = True
                stream.write(converted.read_bytes())
        except BaseException:
            if created:
                output.unlink(missing_ok=True)
            raise
    return {"engine": engine, "source_sha256": before, "output": str(output),
            "note": "Conversion succeeded. Visual fidelity still requires review."}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--engine", choices=["auto", "word", "libreoffice"], default="auto")
    parser.add_argument("--soffice")
    parser.add_argument("--timeout", type=int, default=60)
    args = parser.parse_args()
    try:
        print(json.dumps(convert(args.source, args.output, args.engine, args.soffice, args.timeout),
                         ensure_ascii=False, indent=2))
        return 0
    except (OSError, ValueError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
