# pbid-ct — Power BI Desktop Control Tool

A Python CLI for controlling Power BI Desktop when working with PBIR report files
via [pbi-cli](https://github.com/MinaSaad1/pbi-cli).

## The Problem

Power BI Desktop holds file locks on PBIR JSON files while open. Writing report
changes (adding pages, visuals, filters) while Desktop is running causes access
errors. The built-in `pbi report reload` keyboard-shortcut approach is fragile
on Windows when Desktop isn't the foreground window.

## The Solution

`pbid-ct` owns the Desktop process lifecycle — it closes Desktop, applies your
pbi-cli commands cleanly to disk, then reopens Desktop. It also handles the
OneDrive `ReadOnly` directory attribute that blocks Python's file deletion.

## Requirements

- Windows
- Python 3.10+
- [pbi-cli](https://github.com/MinaSaad1/pbi-cli) (`pipx install pbi-cli-tool`)
- Power BI Desktop

## Usage

```bash
# Close Power BI Desktop
python pbid-ct.py close

# Open a .pbip file (auto-detected from CWD)
python pbid-ct.py open
python pbid-ct.py open --pbip path/to/Report.pbip

# Full cycle: close → apply script → reopen
python pbid-ct.py run tools/my-changes.ps1
python pbid-ct.py run tools/my-changes.ps1 --pbip path/to/Report.pbip
```

## Script Format

Script files contain pbi-cli commands, one per line. Use `--no-sync` on every
write command — `pbid-ct` handles the single Desktop reload at the end.

```bash
# tools/add-revenue-page.ps1
pbi report --no-sync add-page --display-name "Revenue Overview" --name revenue
pbi visual --no-sync add --page revenue --type card --name rev_card
pbi visual --no-sync bind rev_card --page revenue --field "_Measures[Total Revenue]"
pbi visual --no-sync add --page revenue --type bar --name rev_bar
pbi visual --no-sync bind rev_bar --page revenue --category "Dim_Site[Name]" --value "_Measures[Total Revenue]"
```

```bash
python pbid-ct.py run tools/add-revenue-page.ps1
```

## Auto-Detection

`pbid-ct` walks up from the current directory (checking subdirectories one level
deep) to find a `.pbip` file. Run it from anywhere inside your project folder
and it will find the right file automatically.

## Path Setup

Add the `pbid-ct` folder to your user PATH so it's callable from any directory:

```powershell
$path = [Environment]::GetEnvironmentVariable('PATH', 'User')
[Environment]::SetEnvironmentVariable('PATH', "$path;C:\Users\ramon\tools\pbid-ct", 'User')
```

Then call it as `python pbid-ct.py` from any project directory.
