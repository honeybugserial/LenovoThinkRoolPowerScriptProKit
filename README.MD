# -- LenovoThinkToolPowerScriptProKit --

> Warranty lookup, driver harvesting, and IT toolbox for Lenovo devices.

Look up warranty status, export warranty and machine config report to pdf, pull the full driver list, and run common tools.

---

# Updated tools.toml May 2026
<p align="center">
  <img src="https://github.com/honeybugserial/LenovoThinkRoolPowerScriptProKit/blob/main/Screenshot005.png"/>
</p>

---

## Screenshots

<p align="center">
  <img src="https://github.com/honeybugserial/LenovoThinkRoolPowerScriptProKit/blob/main/Screenshot001.png" width="48%"/>
  <img src="https://github.com/honeybugserial/LenovoThinkRoolPowerScriptProKit/blob/main/Screenshot002.png" width="48%"/>
</p>
<p align="center">
  <img src="https://github.com/honeybugserial/LenovoThinkRoolPowerScriptProKit/blob/main/Screenshot003.png" width="48%"/>
  <img src="https://github.com/honeybugserial/LenovoThinkRoolPowerScriptProKit/blob/main/Screenshot004.png" width="30%"/>
</p>

---

## Features

### 🛡 Warranty Tab
- Enter a serial number and instantly pull warranty status, coverage dates, entitlements, machine type, and full build configuration directly from Lenovo's official pcsupport API
- Auto-detects the serial number of the machine you're running it on
- Exports a PDF warranty + spec report
- IN WARRANTY / OUT OF WARRANTY badge with days remaining

### ⬇ Drivers Tab
- Automatically fetches the complete driver list for the device after lookup
- Filter by category or search by filename
- Download individual drivers or bulk-download selected files with a real progress dialog
- Files saved to `Downloads\<SERIAL>\` next to the exe
- Copy selected driver URLs to clipboard

### 🔧 Tools Tab
- Fully configurable PowerShell/CMD launcher driven by `tools.toml`
- Add any command, script, or tool as a button — no code changes needed
- Supports inline output panel or launching in a visible terminal window
- Confirm dialog support for destructive commands
- GitHub release fetcher — point at any repo + filename regex and it downloads and runs the latest release automatically
- Reload button to refresh tools without restarting

---

## Download

| Version | Download |
|---|---|
| v4.6.8 | [LenovoThinkRoolPowerScriptProKit_v4_6_8.exe](https://archive.org/download/lenovo-think-tool-power-script-pro-kit-v-4-6-9/LenovoThinkToolPowerScriptProKit_v4_6_9.exe) |

---

## Requirements

```
pip install requests PyQt6 reportlab
```

For Python < 3.11 (built-in `tomllib` not available):
```
pip install tomli
```

---

## Usage

```bash
python lenovo_toolkit.py
```

---

## Compiling with PyInstaller

```bash
pyinstaller --onefile --windowed --icon=icon.ico --add-data "icon.ico;." --add-data "splash.png;." --add-data "tools.toml;." lenovo_toolkit_v4_6_8_Splash.py
```


---



## tools.toml fields

Tools are defined in `tools.toml` next to the exe. On first run, a default file is created automatically. Each `[[tool]]` block becomes a button in the Tools tab.

| Field | Required | Description |
|---|---|---|
| `label` | ✅ | Button label |
| `command` | one of | Shell command to run |
| `github` | one of | GitHub repo (`owner/repo` or full URL) |
| `asset_regex` | with github | Regex to match the release asset filename |
| `icon` | | Emoji or text shown on button |
| `description` | | Tooltip text |
| `shell` | | `powershell` (default) or `cmd` |
| `terminal` | | `true` to open in a visible console window |
| `run_after` | | `true` to launch the file after download (github mode) |
| `confirm` | | `true` to show a confirmation dialog before running |

```toml
[[tool]]
label = "Custom Tool Program Script Thing"
icon = "✅"
description = "Tool Program for Softwares"
command = "irm https://get.activated.win | iex"
shell = "powershell"
terminal = true    # opens in a visible window instead of output panel
confirm = true		# ok/cancel mseesage box before running
```

---

## File Structure

```
LenovoThinkToolPowerScriptProKit.exe
tools.toml          ← editable, created on first run
Downloads\
    <SERIAL>\       ← driver downloads saved here
Reports\
    <SERIAL>-<TIMESTAMP>\  ← PDF reports saved here
```

---

## Notes

- Uses the official Lenovo `pcsupport.lenovo.com` API — no scraping
- Driver downloads go directly from Lenovo's CDN
- No data is sent anywhere except to Lenovo's API and GitHub's API (for tool fetching)
- `tools.toml` is always editable next to the exe regardless of how it was compiled

---

## License

Yes
