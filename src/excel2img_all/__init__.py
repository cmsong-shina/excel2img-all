# /// script
# requires-python = ">=3.12"
# dependencies = ["pywin32>=306; sys_platform == 'win32'", "Pillow>=11.3,<13"]
# ///
# Based on excel2img by Alexey Gaydyukov (2016), Apache License 2.0.
"""Standalone Excel batch exporter. Run: uv run excel2img-all "**/*.xlsx".

No imports from this repository are needed. Each file gets one private Excel
instance; a session-wide named mutex protects CopyPicture and clipboard reads.
"""

from __future__ import annotations

import argparse
from contextlib import contextmanager, suppress
from dataclasses import dataclass
import glob
import hashlib
import math
import os
from pathlib import Path
import posixpath
import re
import sys
import tempfile
import time
import xml.etree.ElementTree as ET
from zipfile import BadZipFile, ZipFile


NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
EXTENSIONS = {".xlsx", ".xlsm", ".xlsb", ".xls"}
CLIPBOARD_MUTEX = r"Local\excel2img.capture.v1"
ANSI_COLORS = {"info": 36, "success": 32, "warning": 33, "error": 31}


def stdout_supports_color() -> bool:
    if not sys.stdout.isatty():
        return False
    if sys.platform != "win32":
        return True

    # Windows consoles need virtual-terminal processing to interpret ANSI colors.
    import ctypes
    from ctypes import wintypes
    import msvcrt

    try:
        handle = msvcrt.get_osfhandle(sys.stdout.fileno())
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel32.GetConsoleMode.argtypes = [wintypes.HANDLE, ctypes.POINTER(wintypes.DWORD)]
        kernel32.GetConsoleMode.restype = wintypes.BOOL
        kernel32.SetConsoleMode.argtypes = [wintypes.HANDLE, wintypes.DWORD]
        kernel32.SetConsoleMode.restype = wintypes.BOOL
        mode = wintypes.DWORD()
        if not kernel32.GetConsoleMode(handle, ctypes.byref(mode)):
            return False
        enable_virtual_terminal_processing = 0x0004
        return bool(mode.value & enable_virtual_terminal_processing or kernel32.SetConsoleMode(
            handle, mode.value | enable_virtual_terminal_processing))
    except (AttributeError, OSError, ValueError):
        return False


def print_status(message: str, style: str = "info", *, flush: bool = False):
    if stdout_supports_color():
        message = f"\033[{ANSI_COLORS[style]}m{message}\033[0m"
    print(message, flush=flush)


@dataclass(frozen=True)
class Job:
    source: str
    output: str
    gridlines: bool = False
    zoom: int = 100
    retries: int = 3
    clipboard_timeout: float = 8.0


def safe_name(value: str) -> str:
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value)[:100].rstrip(" .") or "untitled"
    reserved = {"CON", "PRN", "AUX", "NUL"} | {f"{prefix}{n}" for prefix in ("COM", "LPT") for n in range(1, 10)}
    return "_" + value if value.upper().split(".")[0] in reserved else value


def expand_inputs(patterns: list[str]) -> list[Path]:
    found = {}
    for pattern in patterns:
        expanded = os.path.expandvars(os.path.expanduser(pattern))
        matches = [expanded] if Path(expanded).is_file() else glob.iglob(expanded, recursive=True)
        accepted = []
        for match in matches:
            path = Path(match)
            if path.is_file() and path.suffix.lower() in EXTENSIONS and not path.name.startswith("~$"):
                accepted.append(path.resolve())
        if not accepted:
            raise ValueError(f"No Excel files matched: {pattern}")
        for path in accepted:
            found[os.path.normcase(str(path))] = path
    return sorted(found.values(), key=lambda p: str(p).casefold())


def output_paths(files: list[Path], root: Path | None) -> dict[Path, Path]:
    proposed = {
        path: root / safe_name(path.stem) if root else path.with_name(f"{safe_name(path.stem)}_png")
        for path in files
    }
    counts = {}
    for output in proposed.values():
        key = str(output).casefold()
        counts[key] = counts.get(key, 0) + 1
    for path, output in list(proposed.items()):
        if counts[str(output).casefold()] > 1:
            digest = hashlib.sha256(os.path.normcase(str(path)).encode("utf-8")).hexdigest()[:12]
            proposed[path] = output.with_name(f"{output.name}_{digest}")
    if len({str(p).casefold() for p in proposed.values()}) != len(files):
        raise ValueError("Output directory collision; choose a different --output-dir.")
    return proposed


def cell_position(address: str) -> tuple[int, int]:
    match = re.fullmatch(r"\$?([A-Za-z]+)\$?([1-9][0-9]*)", address)
    if not match:
        raise ValueError(f"Invalid cell address: {address}")
    column = 0
    for char in match[1].upper():
        column = column * 26 + ord(char) - 64
    return int(match[2]), column


def cell_address(row: int, column: int) -> str:
    letters = ""
    while column:
        column, remainder = divmod(column - 1, 26)
        letters = chr(65 + remainder) + letters
    return f"{letters}{row}"


def content_range(sheet: ET.Element) -> str:
    positions = []
    for cell in sheet.findall(f"{NS}sheetData/{NS}row/{NS}c"):
        value = cell.find(f"{NS}v")
        if (cell.find(f"{NS}f") is not None or (value is not None and value.text is not None)
                or any(t.text for t in cell.findall(f"{NS}is//{NS}t"))):
            positions.append(cell_position(cell.attrib["r"]))
    if not positions:
        return "A1"
    top, bottom = min(p[0] for p in positions), max(p[0] for p in positions)
    left, right = min(p[1] for p in positions), max(p[1] for p in positions)
    merges = []
    for merge in sheet.findall(f"{NS}mergeCells/{NS}mergeCell"):
        first, _, last = merge.attrib["ref"].partition(":")
        a, b = cell_position(first)
        c, d = cell_position(last or first)
        merges.append((a, b, c, d))
    # Expanding one edge can make another merged region intersect the rectangle.
    while True:
        previous = top, left, bottom, right
        for a, b, c, d in merges:
            if a <= bottom and c >= top and b <= right and d >= left:
                top, left, bottom, right = min(top, a), min(left, b), max(bottom, c), max(right, d)
        if previous == (top, left, bottom, right):
            break
    return f"{cell_address(top, left)}:{cell_address(bottom, right)}"


def ooxml_ranges(source: Path) -> dict[str, str]:
    if source.suffix.lower() not in {".xlsx", ".xlsm"}:
        return {}
    with ZipFile(source) as archive:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        targets = {r.attrib["Id"]: r.attrib["Target"] for r in rels if r.get("TargetMode") != "External"}
        result = {}
        for entry in workbook.findall(f"{NS}sheets/{NS}sheet"):
            target = targets[entry.attrib[f"{REL}id"]]
            part = target.lstrip("/") if target.startswith("/") else posixpath.normpath(f"xl/{target}")
            sheet = ET.fromstring(archive.read(part))
            if sheet.tag == f"{NS}worksheet":
                result[entry.attrib["name"]] = content_range(sheet)
        return result


@contextmanager
def clipboard_lock(timeout: float = 120.0):
    """Coordinate all exporter processes in this Windows desktop session."""
    import pythoncom
    import win32event

    handle = win32event.CreateMutex(None, False, CLIPBOARD_MUTEX)
    acquired = False
    try:
        deadline = time.monotonic() + timeout
        while True:
            # Poll without blocking COM message delivery to the STA thread.
            status = win32event.WaitForSingleObject(handle, 0)
            if status in (win32event.WAIT_OBJECT_0, win32event.WAIT_ABANDONED):
                acquired = True
                break
            if status != win32event.WAIT_TIMEOUT:
                raise OSError(f"Clipboard mutex wait failed: {status}")
            if time.monotonic() >= deadline:
                raise TimeoutError("Timed out waiting for another exporter's clipboard capture")
            pythoncom.PumpWaitingMessages()
            time.sleep(0.05)
        yield
    finally:
        try:
            if acquired:
                win32event.ReleaseMutex(handle)
        finally:
            handle.Close()


def clear_clipboard(timeout: float):
    import pythoncom
    import pywintypes
    import win32clipboard

    deadline = time.monotonic() + timeout
    while True:
        try:
            win32clipboard.OpenClipboard()
            try:
                win32clipboard.EmptyClipboard()
            finally:
                win32clipboard.CloseClipboard()
            return
        except pywintypes.error:
            if time.monotonic() >= deadline:
                raise
            pythoncom.PumpWaitingMessages()
            time.sleep(0.1)


def clipboard_image(timeout: float):
    import pythoncom
    from PIL import Image, ImageGrab

    deadline = time.monotonic() + timeout
    last_error = "Excel has not supplied an image yet"
    while time.monotonic() < deadline:
        time.sleep(0.15)
        pythoncom.PumpWaitingMessages()
        try:
            image = ImageGrab.grabclipboard()
            if isinstance(image, Image.Image):
                try:
                    image.load()  # Decode fully before another exporter can replace the clipboard.
                except OSError:
                    image.close()
                    raise
                return image
        except OSError as exc:
            last_error = str(exc)
    raise TimeoutError(f"Clipboard image unavailable after {timeout:g}s: {last_error}")


def wait_for_excel(app, timeout: float, settle: float = 0.25):
    """Let Excel finish calculation and process painting messages before copying.

    Ready alone does not guarantee CopyPicture will succeed, so keep the bounded
    capture retries as well. Reset the settling period if Excel becomes busy.
    """
    import pythoncom
    import pywintypes

    deadline = time.monotonic() + timeout
    settle = min(settle, timeout / 2)
    ready_since = None
    last_error = "Excel is busy or still calculating"
    while True:
        pythoncom.PumpWaitingMessages()
        try:
            ready = bool(app.Ready)
            if ready:
                state = app.CalculationState
                # Manual-calculation workbooks may remain xlPending indefinitely.
                ready = state == 0 or (state == 2 and app.Calculation == -4135)
        except pywintypes.com_error as exc:
            ready = False
            last_error = str(exc)
        now = time.monotonic()
        if ready:
            if ready_since is None:
                ready_since = now
            if now - ready_since >= settle:
                return
        else:
            ready_since = None
        if now >= deadline:
            raise TimeoutError(f"Excel was not ready to capture after {timeout:g}s: {last_error}")
        time.sleep(min(0.05, deadline - now))


def capture_range(app, sheet, address: str, job: Job):
    import pywintypes

    for attempt in range(1, job.retries + 1):
        try:
            with clipboard_lock():
                try:
                    app.CutCopyMode = False
                    app.ScreenUpdating = True
                    sheet.Activate()
                    window = app.ActiveWindow
                    # Avoid triggering another layout pass when retrying the same sheet.
                    if bool(window.DisplayGridlines) != job.gridlines:
                        window.DisplayGridlines = job.gridlines
                    if window.Zoom != job.zoom:
                        window.Zoom = job.zoom
                    region = sheet.Range(address)
                    # Activating the sheet alone can leave hidden Excel unable to
                    # copy again. Select and scroll to the actual capture range.
                    app.Goto(region, True)
                    # Preserve excel2img's workaround for Excel's lazy drawing initialization.
                    for _ in sheet.Shapes:
                        pass
                    wait_for_excel(app, job.clipboard_timeout, settle=0.25 * attempt)
                    clear_clipboard(job.clipboard_timeout)
                    region.CopyPicture(1, 2)  # xlScreen, xlBitmap
                    return clipboard_image(job.clipboard_timeout)
                finally:
                    # A failed CopyPicture can leave Excel in copy mode too.
                    with suppress(pywintypes.com_error):
                        app.CutCopyMode = False
        except (OSError, pywintypes.error, pywintypes.com_error) as exc:
            if attempt == job.retries:
                raise
            print_status(f"Retry [{Path(job.source).name}/{sheet.Name}] {attempt + 1}/{job.retries}: {exc}", "warning", flush=True)
            time.sleep(0.5 * attempt)


def save_png(image, destination: Path):
    """Keep any previous output intact if encoding or writing the new image fails."""
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(dir=destination.parent, suffix=".png", delete=False) as stream:
            temporary = Path(stream.name)
        image.save(temporary, format="PNG")
        temporary.replace(destination)
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)


def export_workbook(job: Job) -> dict:
    """Export one workbook in a private Excel instance, then close it."""
    import pythoncom
    import win32com.client

    source, output = Path(job.source), Path(job.output)
    report = {"source": str(source), "saved": 0, "errors": []}
    app = workbook = None
    pythoncom.CoInitialize()
    try:
        ranges = ooxml_ranges(source)
        output.mkdir(parents=True, exist_ok=True)
        app = win32com.client.DispatchEx("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        app.EnableEvents = False
        app.AskToUpdateLinks = False
        app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        workbook = app.Workbooks.Open(str(source), UpdateLinks=0, ReadOnly=True, IgnoreReadOnlyRecommended=True)
        # A chart sheet's Type is its chart style (e.g. xlXYScatter), not xlChart.
        charts = workbook.Charts
        chart_names = {str(charts.Item(index).Name) for index in range(1, charts.Count + 1)}
        for index in range(1, workbook.Sheets.Count + 1):
            sheet = workbook.Sheets.Item(index)
            name = str(sheet.Name)
            destination = output / f"{index:02d}_{safe_name(name)}.png"
            try:
                if name in chart_names:  # Chart sheets can export without the clipboard.
                    if not sheet.Export(str(destination), "PNG"):
                        raise OSError("Excel could not export the chart sheet")
                    detail = "chart"
                elif sheet.Type == -4167:  # xlWorksheet
                    sheet.Visible = -1
                    address = ranges.get(name) or str(sheet.UsedRange.Address)
                    print_status(f"Exporting [{source.name}/{name}!{address}] ...", flush=True)
                    image = capture_range(app, sheet, address, job)
                    try:
                        detail = f"{address}, {image.width} x {image.height}"
                        save_png(image, destination)
                    finally:
                        image.close()
                else:
                    raise ValueError(f"Unsupported Excel sheet type: {sheet.Type}")
                report["saved"] += 1
                print_status(f"Saved [{source.name}]: {destination} ({detail})", "success", flush=True)
            except Exception as exc:
                report["errors"].append(f"{name}: {exc}")
                print_status(f"Failed [{source.name}/{name}]: {exc}", "error", flush=True)
    except Exception as exc:
        report["errors"].append(str(exc))
        print_status(f"Failed [{source.name}]: {exc}", "error", flush=True)
    finally:
        try:
            if workbook is not None:
                workbook.Close(SaveChanges=False)
        except Exception as exc:
            report["errors"].append(f"Closing workbook: {exc}")
        finally:
            try:
                if app is not None:
                    app.Quit()
            except Exception as exc:
                report["errors"].append(f"Closing Excel: {exc}")
            finally:
                workbook = app = None
                pythoncom.CoUninitialize()
    return report


def positive_int(value: str) -> int:
    number = int(value)
    if number < 1:
        raise argparse.ArgumentTypeError("must be at least 1")
    return number


def positive_float(value: str) -> float:
    number = float(value)
    if not math.isfinite(number) or number <= 0:
        raise argparse.ArgumentTypeError("must be finite and greater than 0")
    return number


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    parser.add_argument("patterns", nargs="+", help='Files or quoted globs, e.g. "**/*.xlsx" "reports/*.xlsm"')
    parser.add_argument("-o", "--output-dir", type=Path, help="Group output folders under this directory (default: beside each source)")
    parser.add_argument("--gridlines", action="store_true", help="Show cell gridlines (default: hidden; styled borders remain)")
    parser.add_argument("--zoom", type=int, default=100, help="Excel view zoom, 10 to 400 (default: 100)")
    parser.add_argument("--retries", type=positive_int, default=3, help="Capture attempts per sheet (default: 3)")
    parser.add_argument("--clipboard-timeout", type=positive_float, default=8.0, help="Seconds to wait for Excel readiness or clipboard access/image (default: 8)")
    parser.add_argument("--dry-run", action="store_true", help="List files, content ranges and output directories without starting Excel")
    args = parser.parse_args(argv)
    if not 10 <= args.zoom <= 400:
        parser.error("--zoom must be between 10 and 400")
    try:
        files = expand_inputs(args.patterns)
        root = args.output_dir.expanduser().resolve() if args.output_dir else None
        outputs = output_paths(files, root)
        if args.dry_run:
            for source in files:
                print_status(f"{source} -> {outputs[source]}")
                ranges = ooxml_ranges(source)
                for name, address in ranges.items():
                    print_status(f"  {name}: {address}")
                if not ranges:
                    print_status("  Excel UsedRange / chart sheet export")
            return 0
        if sys.platform != "win32":
            parser.error("Export requires Windows and installed desktop Microsoft Excel; --dry-run works elsewhere")
        try:
            import pythoncom  # noqa: F401
            import win32com.client  # noqa: F401
            from PIL import Image  # noqa: F401
        except ImportError as exc:
            parser.error(f"Missing dependency ({exc}); run with uv run or uvx as described in the README")
        jobs = [Job(str(p), str(outputs[p]), args.gridlines, args.zoom, args.retries, args.clipboard_timeout) for p in files]
        print_status(f"Exporting {len(jobs)} workbook(s) sequentially.", flush=True)
        reports = [export_workbook(job) for job in jobs]
        saved = sum(report["saved"] for report in reports)
        failures = sum(len(report["errors"]) for report in reports)
        print_status(f"Finished: {len(reports)} workbook(s), {saved} image(s) saved, {failures} error(s).",
                     "error" if failures else "success")
        return 1 if failures else 0
    except (OSError, ValueError, BadZipFile, ET.ParseError, KeyError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        print("Interrupted.", file=sys.stderr)
        return 130


if __name__ == "__main__":
    raise SystemExit(main())
