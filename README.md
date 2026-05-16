# OOF Manager

OOF Manager is a Windows desktop app that helps Microsoft 365 users automate **Out-of-Office / Automatic Replies** without building Power Automate flows by hand.

Its standout feature is the ability to generate a ready-to-import **Power Automate Solution package** (`OofManager-CloudSchedule.zip`). Import the package into the Power Automate environment named after you via **Power Automate > Solutions > Import solution**, connect your Outlook account, and the generated cloud flow will keep scheduling your automatic replies from Microsoft 365 itself. Users do not need to write complex Power Automate date/time expressions, calculate weekday-specific start and end windows, or hand-format HTML reply bodies.

The desktop app is still useful on its own: it can sign in to Exchange Online, edit your reply messages, manage weekday schedules, and sync your next OOF window directly to Outlook. The generated cloud package takes that schedule one step further by keeping it alive even when your PC is off.

Built with WPF (.NET Framework 4.8), [ModernWpfUI](https://github.com/Kinnara/ModernWpf), and [CommunityToolkit.Mvvm](https://learn.microsoft.com/dotnet/communitytoolkit/mvvm/). The local Exchange sync path hosts the [`ExchangeOnlineManagement`](https://www.powershellgallery.com/packages/ExchangeOnlineManagement) PowerShell module in-process, so no `powershell.exe` window pops up.

## Features

- **Generate a Power Automate Solution package** — creates `OofManager-CloudSchedule.zip` on your Desktop and opens the Power Automate Solutions page; import it into the environment named after you.
- **No manual Power Automate expressions** — the app writes the weekday-aware recurrence logic, local time-zone conversion, start/end window calculations, and Outlook automatic-reply action for you.
- **No manual HTML reply formatting** — internal and external reply text from the app is converted into HTML that Outlook renders correctly, preserving line breaks in the cloud flow.
- **Always-on cloud schedule** — after import, the flow runs inside Microsoft 365 and keeps your OOF schedule rolling forward even when OOF Manager and your computer are closed.
- **Personalized package output** — the generated solution includes your mailbox, local Windows time zone, work schedule, internal reply, external reply, and external-audience setting.
- **One-click sign-in** with your Microsoft 365 account (modern auth / device code via `Connect-ExchangeOnline`)
- **View and edit auto-reply state** (Disabled / Scheduled / Enabled) for the signed-in mailbox
- **Separate internal vs. external replies** with rich-text bodies
- **Template library** — save and reuse common auto-reply messages
- **Per-day work-schedule automation** — each weekday gets its own start/end time and an "Off Work" toggle; the app auto-flips OOF on/off at the boundaries
- **Sync to Outlook (works while the app is closed)** — the next off-hours window is pushed to the mailbox as a Scheduled OOF, so Exchange itself toggles auto-replies even if OofManager isn't running
- **Auto-sync** option pushes a fresh window on every schedule check + at sign-in, so a single launch keeps the server in step for the whole week
- **Manual setup guide fallback** — generates a personalized HTML walkthrough for tenants or users who prefer to create the cloud flow manually.
- **Tray notifications** for background OOF flips and auto-syncs while the window is hidden
- **Start with Windows** option (per-user, no admin) launches the app hidden in the tray at logon
- **System tray support** — minimize to tray, single-click to restore, customizable close-button behavior (ask / always-exit / always-tray)
- **Bundled `ExchangeOnlineManagement` module** — works on machines that have never installed the module from PSGallery
- **English UI**

## Why the Power Automate package matters

Creating this flow manually is tedious. A user normally has to:

- Build a scheduled cloud flow from scratch.
- Write Power Automate expressions for weekday-aware dates, time zones, and next-workday calculations.
- Copy the correct internal and external reply bodies into Outlook's automatic-reply action.
- Preserve HTML formatting so Outlook does not collapse the reply into one unreadable paragraph.
- Keep the flow aligned with their actual OOF Manager schedule.

OOF Manager turns that setup into one generated solution zip. The user imports the zip, selects the Outlook connection, and gets a cloud-hosted OOF schedule that mirrors the desktop configuration.

## Requirements

- Windows 10 or later (x64)
- .NET Framework 4.8 (preinstalled on all supported Windows versions)
- A Microsoft 365 / Exchange Online mailbox you can sign in to
- Power Automate access if you want to import the generated cloud-sync Solution package

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
| `Services/CloudSchedulePackageGenerator.cs` | Builds the Power Automate Solution zip for always-on cloud OOF scheduling |
| `Services/CloudScheduleGuideGenerator.cs` | Builds the personalized manual setup guide for the cloud flow |
| `Services/PreferencesService.cs` | Persists settings to `%LocalAppData%\OofManager\preferences.json` |
| `Services/TemplateService.cs` | Template library (SQLite via `sqlite-net-pcl`) |
| `Services/TrayIconService.cs` | `NotifyIcon`-based system tray support + balloon notifications |
| `Services/StartupService.cs` | Per-user autostart via HKCU `Run` key (writes `--minimized` so the app launches into the tray) |
| `Modules/ExchangeOnlineManagement/` | Bundled PowerShell module shipped with the app |
| `Installer/OofManager.iss` | Inno Setup script |

## License

MIT — see [LICENSE](LICENSE).
