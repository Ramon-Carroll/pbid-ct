#!/usr/bin/env python3
"""
pbid-ct.py  --  Power BI Desktop Control Tool
Generic pbi-cli report wrapper for any PBIR (.pbip) project.

Subcommands:
  open   -- Open a .pbip file in Power BI Desktop
  close  -- Close Power BI Desktop
  run    -- Full cycle: close Desktop, run pbi-cli script, reopen Desktop

Usage:
  python pbid-ct.py open [--pbip path/to/Report.pbip]
  python pbid-ct.py close
  python pbid-ct.py run <script-file> [--pbip path/to/Report.pbip]

The .pbip file is auto-detected from CWD (and one level of subdirs) if --pbip is omitted.
Script files contain pbi-cli commands (one per line). Blank lines and lines
starting with # or // are ignored. Use --no-sync on write commands.
"""

import argparse
import os
import shlex
import subprocess
import sys
import time
from pathlib import Path


# ---------------------------------------------------------------------------
# Utilities
# ---------------------------------------------------------------------------

def find_pbip(start: Path) -> Path | None:
    """Search CWD (including one level of subdirs), then walk up the tree."""
    matches = list(start.glob("*.pbip")) + list(start.glob("*/*.pbip"))
    if matches:
        return matches[0]
    for directory in start.parents:
        matches = list(directory.glob("*.pbip"))
        if matches:
            return matches[0]
    return None


def resolve_pbip(pbip_arg: str | None) -> Path:
    """Return resolved .pbip path, or exit with a clear error."""
    if pbip_arg:
        path = Path(pbip_arg).resolve()
    else:
        path = find_pbip(Path.cwd())
    if not path or not path.exists():
        print(
            "Error: No .pbip file found. Pass --pbip or run from inside a .pbip project directory.",
            flush=True,
        )
        sys.exit(1)
    return path


def find_pbi_pid() -> int | None:
    """Return the PID of a running PBIDesktop.exe, or None."""
    result = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq PBIDesktop.exe", "/FO", "CSV", "/NH"],
        capture_output=True, text=True,
    )
    for line in result.stdout.splitlines():
        if "PBIDesktop.exe" in line:
            try:
                return int(line.split(",")[1].strip('"'))
            except (IndexError, ValueError):
                pass
    return None


def kill_pbi(pid: int) -> None:
    """Force-kill PBI Desktop and wait for it to fully exit."""
    subprocess.run(["taskkill", "/F", "/PID", str(pid)], capture_output=True)
    for _ in range(50):  # poll up to 5s
        check = subprocess.run(
            ["tasklist", "/FI", f"PID eq {pid}", "/FO", "CSV", "/NH"],
            capture_output=True, text=True,
        )
        if str(pid) not in check.stdout:
            break
        time.sleep(0.1)
    time.sleep(2.0)  # OneDrive sync settling delay


def clear_readonly_dirs(path: Path) -> None:
    """Strip ReadOnly attribute from all directories under path.

    OneDrive marks placeholder directories ReadOnly. Python's shutil.rmtree
    treats this as Access Denied. Clearing with attrib before pbi-cli runs
    avoids the error on delete operations.
    """
    subprocess.run(
        ["attrib", "-R", str(path / "*"), "/S", "/D"],
        capture_output=True,
    )


def parse_script(script_path: Path) -> list[str]:
    """Return non-blank, non-comment lines from script file."""
    commands = []
    with open(script_path) as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and not line.startswith("//"):
                commands.append(line)
    return commands


# ---------------------------------------------------------------------------
# Subcommand handlers
# ---------------------------------------------------------------------------

def cmd_open(args: argparse.Namespace) -> None:
    """Open a .pbip file in Power BI Desktop."""
    pbip_path = resolve_pbip(args.pbip)
    print(f"Opening {pbip_path.name}...", flush=True)
    os.startfile(str(pbip_path))
    print("Done.", flush=True)


def cmd_close(args: argparse.Namespace) -> None:
    """Close Power BI Desktop if it is running."""
    pid = find_pbi_pid()
    if pid:
        print(f"Closing PBI Desktop (PID {pid})...", flush=True)
        kill_pbi(pid)
        print("Done.", flush=True)
    else:
        print("PBI Desktop is not open.", flush=True)


def cmd_run(args: argparse.Namespace) -> None:
    """Full cycle: close Desktop, run pbi-cli script, reopen Desktop."""
    script_path = Path(args.script).resolve()
    if not script_path.exists():
        print(f"Error: Script not found: {script_path}", flush=True)
        sys.exit(1)

    pbip_path = resolve_pbip(args.pbip)
    report_path = pbip_path.parent / (pbip_path.stem + ".Report")
    if not report_path.exists():
        print(f"Error: Report folder not found: {report_path}", flush=True)
        sys.exit(1)

    print(f"Project: {pbip_path.name}", flush=True)

    pid = find_pbi_pid()
    if pid:
        print(f"PBI Desktop is open (PID {pid}). Closing to apply report changes...", flush=True)
        kill_pbi(pid)
        was_open = True
    else:
        print("PBI Desktop is not open. Applying report changes to disk...", flush=True)
        was_open = False

    clear_readonly_dirs(report_path)

    os.chdir(report_path)
    for cmd in parse_script(script_path):
        print(f"  {cmd}", flush=True)
        result = subprocess.run(shlex.split(cmd))
        if result.returncode != 0:
            print(f"Error: command failed with exit code {result.returncode}", flush=True)
            sys.exit(result.returncode)

    if was_open:
        print("Reopening PBI Desktop...", flush=True)
        os.startfile(str(pbip_path))

    print("Done.", flush=True)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description="pbi-cli report wrapper — open, close, or run a full write cycle.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    # open
    p_open = subparsers.add_parser("open", help="Open a .pbip file in Power BI Desktop")
    p_open.add_argument("--pbip", help="Path to the .pbip file (auto-detected if omitted)")

    # close
    subparsers.add_parser("close", help="Close Power BI Desktop")

    # run
    p_run = subparsers.add_parser(
        "run", help="Close Desktop, run pbi-cli script, reopen Desktop"
    )
    p_run.add_argument("script", help="Script file containing pbi-cli commands (one per line)")
    p_run.add_argument("--pbip", help="Path to the .pbip file (auto-detected if omitted)")

    args = parser.parse_args()

    if args.command == "open":
        cmd_open(args)
    elif args.command == "close":
        cmd_close(args)
    elif args.command == "run":
        cmd_run(args)


if __name__ == "__main__":
    main()
