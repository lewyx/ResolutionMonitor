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

Two builds are attached to each [release](../../releases):

| Build | File | Size | Requirements |
|-------|------|------|-------------|
| **Framework** | `ResMon-framework.exe` | ~20 KB | .NET Framework 4.x (preinstalled on Windows 10/11) |
| **Standalone** | `ResMon-standalone.exe` | ~70 MB | None (self-contained) |

**Architecture coverage:**
- The **Framework** build is compiled as **AnyCPU** — it runs natively on both 32-bit and 64-bit Windows.
- The **Standalone** build targets **x64** only. For the rare case of 32-bit Windows, use the Framework build instead.

Download the appropriate exe and run it. No installation required.

## Usage

- **Right-click** or **left-click** the tray icon to open the menu
- **Apply Target** - immediately sets your display to the configured resolution and scaling
- **Options > Target resolution** - choose from standard resolutions or enter a custom one
- **Options > Target scaling** - choose from standard DPI scaling values or enter a custom one
- **Options > Auto start** - toggle auto-start with Windows (registers via `HKCU\...\Run`)

Settings are persisted in the Windows Registry under `HKCU\Software\PanSoft\Resolution Monitor`.

## Development

The source is a single C# file (`ResolutionMonitor.cs`).

### Build with .NET Framework (AnyCPU, no SDK required)

```cmd
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe /target:winexe /optimize /r:System.Windows.Forms.dll /r:System.Drawing.dll /out:ResMon.exe ResolutionMonitor.cs
```

### Build with .NET 8

```cmd
dotnet build -c Release
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -o publish
```

### CI

A GitHub Actions workflow automatically builds both variants and attaches them to each release.

## License

[MIT](LICENSE)
