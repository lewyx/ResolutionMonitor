# ResolutionMonitor - Resolution Monitor

A lightweight Windows system tray tool that monitors your display resolution and DPI scaling, and lets you quickly switch to a preferred configuration.

## Features

- System tray icon shows current state at a glance (monitor icon = on target, warning icon = mismatch)
- Balloon notification when resolution or scaling changes
- One-click apply to restore your target resolution and scaling
- Configurable target resolution and scaling via context menu
- Custom resolution and scaling input
- Optional auto-start with Windows
- Responds instantly to display change events, with 30-second polling as backup

## Download

Download **[ResolutionMonitor.vbs](ResolutionMonitor.vbs)** and double-click to run. No installation required.

Requires Windows with PowerShell 5.1+ (included in Windows 10/11).

## Usage

- **Right-click** or **left-click** the tray icon to open the menu
- **Apply Target** - immediately sets your display to the configured resolution and scaling
- **Options > Target resolution** - choose from standard resolutions or enter a custom one
- **Options > Target scaling** - choose from standard DPI scaling values or enter a custom one
- **Options > Auto start** - toggle auto-start with Windows (registers via `HKCU\...\Run`)

Settings are persisted in the Windows Registry under `HKCU\Software\PanSoft\Resolution Monitor`.

## Development

The main source is `ResolutionMonitor.ps1`. To run directly:

```powershell
powershell -ExecutionPolicy Bypass -NoProfile -File ResolutionMonitor.ps1
```

The distributable `ResolutionMonitor.vbs` is a VBScript wrapper that launches same embedded PowerShell script hidden (no console window). It is automatically rebuilt via a pre-commit hook whenever `ResolutionMonitor.ps1` changes.

To rebuild manually:

```powershell
powershell -ExecutionPolicy Bypass -NoProfile -File build.ps1
```

## License

[MIT](LICENSE)