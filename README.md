# OOF Manager

A Windows desktop app for managing **Out-of-Office (Automatic Reply)** settings on Microsoft 365 / Exchange Online mailboxes — without ever opening Outlook or a browser.

Built with WPF (.NET Framework 4.8), [ModernWpfUI](https://github.com/Kinnara/ModernWpf), and [CommunityToolkit.Mvvm](https://learn.microsoft.com/dotnet/communitytoolkit/mvvm/). Talks to Exchange Online by hosting the [`ExchangeOnlineManagement`](https://www.powershellgallery.com/packages/ExchangeOnlineManagement) PowerShell module in-process (no `powershell.exe` window pops up).

## Features

- **One-click sign-in** with your Microsoft 365 account (modern auth / device code via `Connect-ExchangeOnline`)
- **View and edit auto-reply state** (Disabled / Scheduled / Enabled) for the signed-in mailbox
- **Separate internal vs. external replies** with rich-text bodies
- **Template library** — save and reuse common auto-reply messages
- **Work-schedule automation** — pick weekdays + start/end times (in 30-minute increments) and the app will toggle OOF on/off automatically every 5 minutes
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
| `Services/TrayIconService.cs` | `NotifyIcon`-based system tray support |
| `Modules/ExchangeOnlineManagement/` | Bundled PowerShell module shipped with the app |
| `Installer/OofManager.iss` | Inno Setup script |

## License

MIT — see [LICENSE](LICENSE).
