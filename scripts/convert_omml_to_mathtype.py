"""Convert Word OMML equations to MathType equations after Pandoc export.

Author: Luo Qiang, Jimei University

Default workflow for the parent skill remains native Word OMML. Run this script
only after the user explicitly confirms they need MathType output.

Requirements:
  - Windows
  - Microsoft Word
  - MathType installed as Word add-in
  - pywin32 (`py -m pip install pywin32`)

Usage:
  py convert_omml_to_mathtype.py --check
  py convert_omml_to_mathtype.py input.docx
  py convert_omml_to_mathtype.py input.docx output_mathtype.docx
"""

from __future__ import annotations

import argparse
import shutil
import sys
import time
import traceback
import zipfile
from pathlib import Path


def check_mathtype_installed() -> bool:
    """Return True when MathType appears registered on this machine."""
    if sys.platform != "win32":
        return False

    import winreg

    registry_keys = [
        r"Equation.DSMT4",
        r"CLSID\{0002CE03-0000-0000-C000-000000000046}",
    ]
    for key_path in registry_keys:
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, key_path):
                return True
        except OSError:
            continue
    return False


def _import_com_modules():
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError(
            "pywin32 not installed. Run in PowerShell: py -m pip install pywin32"
        ) from exc
    return pythoncom, win32com.client


def convert_to_mathtype(input_path: str, output_path: str | None = None) -> bool:
    """Convert all OMML equations in input_path and save output_path."""
    if sys.platform != "win32":
        print("MathType COM conversion only works on Windows.")
        return False

    input_file = Path(input_path).resolve()
    if not input_file.exists():
        print(f"File not found: {input_file}")
        return False

    if not check_mathtype_installed():
        print("MathType not detected. Keep native OMML output.")
        return False

    output_file = (
        Path(output_path).resolve()
        if output_path
        else input_file.with_name(f"{input_file.stem}_mathtype{input_file.suffix}")
    )
    if input_file != output_file:
        shutil.copy2(input_file, output_file)

    pythoncom, win32_client = _import_com_modules()
    pythoncom.CoInitialize()
    word = None

    try:
        word = win32_client.Dispatch(
            "Word.Application", clsctx=pythoncom.CLSCTX_LOCAL_SERVER
        )
        word.Visible = True
        word.DisplayAlerts = 0
        print(f"Word started: v{word.Version}")

        doc = word.Documents.Open(str(output_file))
        total = int(doc.OMaths.Count)
        print(f"Opened: {output_file.name}")
        print(f"OMML equations found: {total}")

        if total == 0:
            doc.Save()
            doc.Close()
            print(f"No OMML equations. Saved unchanged: {output_file}")
            return True

        before_stats = inspect_open_doc_math(doc)
        converted = _run_convert_equations_macro(word, win32_client)
        time.sleep(1.0)
        active_doc = word.ActiveDocument
        remaining = int(active_doc.OMaths.Count)
        active_doc.Save()
        active_doc.Close()

        output_stats = inspect_docx_math(output_file)
        has_real_conversion = (
            output_stats["ole_objects"] > 0
            or output_stats["equation_dsmt4"] > 0
            or output_stats["omath"] < total
        )

        print("")
        print("MathType conversion complete.")
        print(f"  Input:     {input_file}")
        print(f"  Output:    {output_file}")
        print(f"  Conversion macro ran: {converted}")
        print(f"  Initial OMML: {before_stats['omath']}")
        print(f"  Remaining OMML: {remaining}")
        print(f"  Output OLE objects: {output_stats['ole_objects']}")
        print(f"  Output Equation.DSMT4 markers: {output_stats['equation_dsmt4']}")

        if not has_real_conversion:
            print("")
            print("Conversion did not produce MathType OLE objects.")
            print("Keep native OMML output or convert manually in Word/MathType.")
            return False

        return True

    except Exception as exc:
        print(f"Conversion error: {exc}")
        traceback.print_exc()
        return False
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def _run_convert_equations_macro(word, win32_client) -> bool:
    """Run MathType's full Convert Equations command.

    This is the same conversion dialog users open manually. MathType persists the
    user's checkbox settings, so configure the dialog once with all equation
    types checked and target "MathType equation (OLE object)".
    """
    for macro in ("MTCommand_ConvertEqns", "MTCommand_OnConvertEquationsD"):
        try:
            _schedule_enter_key(win32_client, delay_seconds=1.5)
            word.Run(macro)
            time.sleep(2.0)
            print(f"  OK via Word.Run({macro})")
            return True
        except Exception as exc:
            print(f"  Word.Run({macro}) failed: {exc}")
    return False


def _schedule_enter_key(win32_client, delay_seconds: float = 1.5) -> None:
    """Send Enter shortly after MathType opens the modal Convert dialog.

    Word.Run is blocking while the dialog is open, so this uses WScript.Shell
    asynchronously via a separate process.
    """
    import subprocess

    first_delay = int(delay_seconds * 1000)
    ps = (
        "$ws = New-Object -ComObject WScript.Shell; "
        f"Start-Sleep -Milliseconds {first_delay}; "
        "$ws.SendKeys('{ENTER}'); "
        "Start-Sleep -Milliseconds 2500; "
        "$ws.SendKeys('{ENTER}'); "
        "Start-Sleep -Milliseconds 1500; "
        "$ws.SendKeys('{ENTER}')"
    )
    subprocess.Popen(
        ["powershell", "-NoProfile", "-WindowStyle", "Hidden", "-Command", ps],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


def inspect_docx_math(docx_path: str | Path) -> dict[str, int]:
    """Count OMML and MathType/OLE markers in a docx package."""
    path = Path(docx_path)
    stats = {"omath": 0, "equation_dsmt4": 0, "ole_objects": 0}

    try:
        with zipfile.ZipFile(path) as zf:
            names = zf.namelist()
            stats["ole_objects"] = sum(
                1 for name in names if name.startswith("word/embeddings/")
            )
            xml = zf.read("word/document.xml").decode("utf-8", errors="replace")
    except Exception as exc:
        print(f"Could not inspect output docx: {exc}")
        return stats

    stats["omath"] = xml.count("<m:oMath")
    stats["equation_dsmt4"] = xml.count("Equation.DSMT4")
    return stats


def inspect_open_doc_math(doc) -> dict[str, int]:
    try:
        omath = int(doc.OMaths.Count)
    except Exception:
        omath = 0
    try:
        inline_shapes = int(doc.InlineShapes.Count)
    except Exception:
        inline_shapes = 0
    return {"omath": omath, "inline_shapes": inline_shapes}


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Convert OMML equations in a .docx to MathType equations."
    )
    parser.add_argument("input", nargs="?", help="Input .docx with native OMML equations")
    parser.add_argument("output", nargs="?", help="Output .docx path")
    parser.add_argument("--check", action="store_true", help="Only check MathType install")
    args = parser.parse_args()

    if args.check:
        ok = check_mathtype_installed()
        print("MathType detected" if ok else "MathType not detected")
        return 0 if ok else 1

    if not args.input:
        parser.error("input is required unless --check is used")

    return 0 if convert_to_mathtype(args.input, args.output) else 1


if __name__ == "__main__":
    raise SystemExit(main())
