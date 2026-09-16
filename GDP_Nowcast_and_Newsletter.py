"""Refresh the existing Nigeria model and newsletter without a PowerShell script."""
from __future__ import annotations

import argparse
import calendar
import hashlib
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
from datetime import date


def shared_hash(path):
    """Read only; permit Excel's shared access on Windows."""
    if os.name == "nt":
        import ctypes
        from ctypes import wintypes
        import msvcrt
        kernel = ctypes.WinDLL("kernel32", use_last_error=True)
        create = kernel.CreateFileW
        create.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
                           wintypes.LPVOID, wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
        create.restype = wintypes.HANDLE
        handle = create(str(path), 0x80000000, 3, None, 3, 0x80, None)
        if handle == wintypes.HANDLE(-1).value:
            raise ctypes.WinError(ctypes.get_last_error())
        try:
            descriptor = msvcrt.open_osfhandle(handle, os.O_RDONLY | os.O_BINARY)
        except Exception:
            close = kernel.CloseHandle
            close.argtypes = [wintypes.HANDLE]
            close(handle)
            raise
        stream = os.fdopen(descriptor, "rb")
    else:
        stream = open(path, "rb")
    with stream:
        digest = hashlib.sha256()
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
        return digest.hexdigest()


def month_value(value):
    if not re.fullmatch(r"\d{4}-(0[1-9]|1[0-2])", value):
        raise argparse.ArgumentTypeError("Use YYYY-MM, for example 2026-10.")
    try:
        date.fromisoformat(value + "-01")
    except ValueError as exc:
        raise argparse.ArgumentTypeError(str(exc)) from exc
    return value


def report_paths(root, month, template=None):
    year, number = map(int, month.split("-"))
    name = calendar.month_name[number]
    folder = root / str(year) / f"{name} Newsletter"
    stem = f"CAPE Economic Performance and Prospect Bulletin {name} {year}"
    output = folder / f"{stem}.docx"
    previous_number = number - 1 or 12
    previous_year = year if number > 1 else year - 1
    previous_name = calendar.month_name[previous_number]
    previous = root / str(previous_year) / f"{previous_name} Newsletter" / (
        f"CAPE Economic Performance and Prospect Bulletin {previous_name} {previous_year}.docx")
    candidates = [Path(template).resolve()] if template else [output, folder / f"{stem} - Template.docx", previous]
    selected = next((p for p in candidates if p.is_file()), None)
    if selected is None:
        raise FileNotFoundError("No newsletter template found. Supply --newsletter-template PATH. Checked: "
                                + "; ".join(map(str, candidates)))
    return output, selected


def stamp(path):
    return path.stat().st_mtime_ns if path.exists() else None


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("report_month", nargs="?", type=month_value, default=date.today().strftime("%Y-%m"),
                        help="Newsletter month YYYY-MM (default: current month); model quarter remains data-driven.")
    parser.add_argument("--newsletter-template", type=Path)
    parser.add_argument("--plots-folder", type=Path)
    parser.add_argument("--dry-run", action="store_true", help="Check paths and show commands; do not change any files.")
    args = parser.parse_args(argv)
    project_root = Path(__file__).resolve().parent
    nowcast_root = project_root / "Nigeria GDP Nowcasting"
    model_root = project_root.parent
    automation = nowcast_root / ".automation"
    master = model_root / "Main Data" / "MultipleFrequencyModelData.xlsx"
    newsletter_root = project_root / "CAPE ERC Economic Newsletter"
    output, template = report_paths(newsletter_root, args.report_month, args.newsletter_template)
    plots = args.plots_folder.resolve() if args.plots_folder else output.parent / "Plots"
    bundled_node = Path.home() / ".cache/codex-runtimes/codex-primary-runtime/dependencies/node/bin/node.exe"
    node = str(bundled_node) if bundled_node.is_file() else shutil.which("node")
    if not node:
        raise FileNotFoundError("Node.js is required by the existing workbook renderer but was not found.")
    for path in [master, *(automation / f for f in (
            "analyze_nowcast.py", "build_workbooks.mjs", "verify_refresh.py", "publish_newsletter.py"))]:
        if not path.is_file():
            raise FileNotFoundError(path)
    python = [sys.executable, "-X", "utf8"]
    workbook = nowcast_root / "Nigeria Quarterly GDP Nowcast.xlsx"
    commands = [
        python + [str(automation / "analyze_nowcast.py")],
        [node, str(automation / "build_workbooks.mjs")],
        python + [str(automation / "verify_refresh.py"), str(nowcast_root)],
        python + [str(automation / "publish_newsletter.py"), "--workbook", str(workbook),
                  "--template", str(template), "--output", str(output),
                  "--chart", str(nowcast_root / "Newsletter Assets/Nigeria GDP Nowcast Chart.png"),
                  "--pmi-chart", str(nowcast_root / "Newsletter Assets/Nigeria PMI Trend Chart.png"),
                  "--report-month", args.report_month, "--plots-folder", str(plots)],
    ]
    if args.dry_run:
        print("DRY RUN: no model estimation, files changed, or report published.")
        print(f"Read-only master: {master}\nReport: {output}")
        for command in commands:
            print(subprocess.list2cmdline(command))
        return 0
    env = dict(os.environ, NIGERIA_NOWCAST_OUTPUT_DIR=str(nowcast_root), NIGERIA_SKIP_SOURCE_UPDATE="1")

    def run(index, check=True):
        return subprocess.run(commands[index], cwd=nowcast_root, env=env, check=check).returncode

    before = shared_hash(master)
    results = automation / "nowcast_results.json"
    results_before = stamp(results)
    run(0)
    if stamp(results) is None or stamp(results) == results_before:
        raise RuntimeError("Estimation did not refresh nowcast_results.json; stopping.")
    workbook_before = stamp(workbook)
    builder_status = run(1, check=False)
    if stamp(workbook) is None or stamp(workbook) == workbook_before:
        raise RuntimeError(f"Builder did not save a fresh workbook (status {builder_status}); stopping.")
    run(2)
    if shared_hash(master) != before:
        raise RuntimeError("Master workbook changed during estimation/build; publication stopped. Do not edit during refresh.")
    if builder_status:
        print(f"Warning: renderer exited with status {builder_status}; freshly saved workbook passed validation.")
    output.parent.mkdir(parents=True, exist_ok=True)
    plots.mkdir(parents=True, exist_ok=True)
    run(3)
    if shared_hash(master) != before:
        raise RuntimeError("Master workbook changed during publication. Review the report and rerun after saving inputs.")
    print(f"Nigeria GDP nowcast refreshed: {workbook}")
    print(f"Economic newsletter refreshed: {output}")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except (OSError, RuntimeError, subprocess.CalledProcessError) as exc:
        print(f"Refresh failed: {exc}", file=sys.stderr)
        sys.exit(1)
