# OOF Manager

A Windows desktop app for managing **Out-of-Office (Automatic Reply)** settings on Microsoft 365 / Exchange Online mailboxes — without ever opening Outlook or a browser.

Built with WPF (.NET Framework 4.8), [ModernWpfUI](https://github.com/Kinnara/ModernWpf), and [CommunityToolkit.Mvvm](https://learn.microsoft.com/dotnet/communitytoolkit/mvvm/). Talks to Exchange Online by hosting the [`ExchangeOnlineManagement`](https://www.powershellgallery.com/packages/ExchangeOnlineManagement) PowerShell module in-process (no `powershell.exe` window pops up).

## Features

- **One-click sign-in** with your Microsoft 365 account (modern auth / device code via `Connect-ExchangeOnline`)
- **View and edit auto-reply state** (Disabled / Scheduled / Enabled) for the signed-in mailbox
- **Separate internal vs. external replies** with rich-text bodies
- **Template library** — save and reuse common auto-reply messages
- **Per-day work-schedule automation** — each weekday gets its own start/end time and an "Off Work" toggle; the app auto-flips OOF on/off at the boundaries
- **Sync to Outlook (works while the app is closed)** — the next off-hours window is pushed to the mailbox as a Scheduled OOF, so Exchange itself toggles auto-replies even if OofManager isn't running
- **Auto-sync** option pushes a fresh window on every schedule check + at sign-in, so a single launch keeps the server in step for the whole week
- **Tray notifications** for background OOF flips and auto-syncs while the window is hidden
- **Start with Windows** option (per-user, no admin) launches the app hidden in the tray at logon
- **System tray support** — minimize to tray, single-click to restore, customizable close-button behavior (ask / always-exit / always-tray)
- **Bundled `ExchangeOnlineManagement` module** — works on machines that have never installed the module from PSGallery
- **English UI**

## Requirements

- Windows 10 or later (x64)
- .NET Framework 4.8 (preinstalled on all supported Windows versions)
- A Microsoft 365 / Exchange Online mailbox you can sign in to

## Installing

Grab the latest **`OofManagerSetup.exe`** from the [Releases page](../../releases) and run it. The installer is per-user by default (no admin rights required).

## Building from source

```powershell
# Restore + build
dotnet build OofManager.Wpf.csproj -c Release

# Publish (produces a self-contained folder under .\publish\ ready for the installer)
dotnet publish OofManager.Wpf.csproj -c Release -o publish

# Build the installer (requires Inno Setup 6: winget install JRSoftware.InnoSetup)
& "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe" .\Installer\OofManager.iss
# -> .\Installer\Output\OofManagerSetup.exe
```

## Project layout

| Path | What it does |
| --- | --- |
| `Views/`, `ViewModels/` | XAML pages + MVVM view-models (Login, Main) |
| `Services/ExchangeService.cs` | In-process PowerShell runspace; talks to Exchange Online |
| `Services/PreferencesService.cs` | Persists settings to `%LocalAppData%\OofManager\preferences.json` |
| `Services/TemplateService.cs` | Template library (SQLite via `sqlite-net-pcl`) |
| `Services/TrayIconService.cs` | `NotifyIcon`-based system tray support + balloon notifications |
| `Services/StartupService.cs` | Per-user autostart via HKCU `Run` key (writes `--minimized` so the app launches into the tray) |
| `Modules/ExchangeOnlineManagement/` | Bundled PowerShell module shipped with the app |
| `Installer/OofManager.iss` | Inno Setup script |

## License

MIT — see [LICENSE](LICENSE).
