#!/usr/bin/env python3
"""
pbid-ct.py  --  Power BI Desktop Control Tool

Subcommands:
  open   -- Open a .pbip file in Power BI Desktop
  close  -- Close Power BI Desktop
  save   -- Close Desktop if open, then reopen to pick up file changes

Usage:
  python pbid-ct.py open [--pbip path/to/Report.pbip]
  python pbid-ct.py close
  python pbid-ct.py save [--pbip path/to/Report.pbip]

The .pbip file is auto-detected from CWD (and one level of subdirs) if --pbip is omitted.
"""

import argparse
import os
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


def cmd_save(args: argparse.Namespace) -> None:
    """Close Desktop if open, then reopen to pick up file changes on disk."""
    pid = find_pbi_pid()
    if pid:
        pbip_path = resolve_pbip(args.pbip)
        print(f"PBI Desktop is open (PID {pid}). Closing to pick up changes...", flush=True)
        kill_pbi(pid)
        print(f"Reopening {pbip_path.name}...", flush=True)
        os.startfile(str(pbip_path))
        print("Done.", flush=True)
    else:
        print("PBI Desktop is not open. Changes are on disk.", flush=True)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Power BI Desktop lifecycle manager — open, close, or reload after file changes.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    # open
    p_open = subparsers.add_parser("open", help="Open a .pbip file in Power BI Desktop")
    p_open.add_argument("--pbip", help="Path to the .pbip file (auto-detected if omitted)")

    # close
    subparsers.add_parser("close", help="Close Power BI Desktop")

    # save
    p_save = subparsers.add_parser(
        "save", help="Close Desktop if open, reopen to pick up file changes"
    )
    p_save.add_argument("--pbip", help="Path to the .pbip file (auto-detected if omitted)")

    args = parser.parse_args()

    if args.command == "open":
        cmd_open(args)
    elif args.command == "close":
        cmd_close(args)
    elif args.command == "save":
        cmd_save(args)


if __name__ == "__main__":
    main()
