# OOF Manager

OOF Manager is a Windows desktop app that turns your Microsoft 365 **Out-of-Office / Automatic Replies** schedule into a cloud-hosted Power Automate flow.

Its standout feature is the end-to-end Power Automate setup: OOF Manager generates a personalized **Solution package** (`OofManager-CloudSchedule.zip`), imports it into your Power Automate environment automatically when possible, and can then find the generated flow so you can turn it **on or off directly from the desktop app**. The zip is kept under the app's local data folder as a manual fallback for locked-down tenants.

Once connected, the generated cloud flow keeps your automatic replies scheduled from Microsoft 365 itself, even when OOF Manager and your computer are closed. You do not need to write Power Automate date/time expressions, calculate weekday-specific start and end windows, or hand-format HTML reply bodies.

Built with WPF (.NET Framework 4.8), [ModernWpfUI](https://github.com/Kinnara/ModernWpf), and [CommunityToolkit.Mvvm](https://learn.microsoft.com/dotnet/communitytoolkit/mvvm/). The local Exchange sync path hosts the [`ExchangeOnlineManagement`](https://www.powershellgallery.com/packages/ExchangeOnlineManagement) PowerShell module in-process, so no `powershell.exe` window pops up.

## Features

- **Generate & auto-import a Power Automate Solution** — creates `OofManager-CloudSchedule.zip` under the app's local data folder and imports it automatically when possible; if automatic import is blocked, the app opens Power Automate for manual import.
- **Turn the cloud flow on/off from OOF Manager** — after the solution is imported, the app locates the generated Power Automate flow and lets you enable or disable it without hunting through the Maker portal.
- **Always-on cloud schedule** — after import, the flow runs inside Microsoft 365 and keeps your OOF schedule rolling forward even when OOF Manager and your computer are closed.
- **No manual Power Automate expressions** — the app writes the weekday-aware recurrence logic, local time-zone conversion, start/end window calculations, and Outlook automatic-reply action for you.
- **No manual HTML reply formatting** — internal and external reply text from the app is converted into HTML that Outlook renders correctly, preserving line breaks in the cloud flow.
- **Personalized package output** — the generated solution includes your mailbox, local Windows time zone, work schedule, internal reply, external reply, and external-audience setting.
- **One-click sign-in** with your Microsoft 365 account (modern auth / device code via `Connect-ExchangeOnline`)
- **View and edit auto-reply state** (Disabled / Scheduled / Enabled) for the signed-in mailbox
- **Separate internal vs. external replies** with rich-text bodies
- **Template library** — save and reuse common auto-reply messages
- **Per-day work-schedule planning** — each weekday gets its own start/end time and an "Off Work" toggle
- **Sync to Outlook (works while the app is closed)** — the next off-hours window is pushed to the mailbox as a Scheduled OOF, so Exchange itself toggles auto-replies even if OofManager isn't running
- **Power Automate cloud schedule** keeps the work schedule advancing even when all local computers are off
- **Manual setup guide fallback** — generates a personalized HTML walkthrough for tenants or users who prefer to create the cloud flow manually.
- **Bundled `ExchangeOnlineManagement` module** — works on machines that have never installed the module from PSGallery
- **English UI**

## Why the Power Automate package matters

Creating this flow manually is tedious. A user normally has to:

- Build a scheduled cloud flow from scratch.
- Write Power Automate expressions for weekday-aware dates, time zones, and next-workday calculations.
- Copy the correct internal and external reply bodies into Outlook's automatic-reply action.
- Preserve HTML formatting so Outlook does not collapse the reply into one unreadable paragraph.
- Keep the flow aligned with their actual OOF Manager schedule.

OOF Manager turns that setup into one generated solution zip, imports it automatically when the tenant allows it, and keeps track of the generated flow afterward. From the desktop app you can refresh the flow state, turn the cloud schedule on or off, and still fall back to the Maker portal when a tenant policy blocks automation.

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

# Build the installer (requires Inno Setup 6: winget install --id JRSoftware.InnoSetup -e)
.\Tools\build-installer.ps1
# -> .\Installer\Output\OofManagerSetup.exe
```

## Project layout

| Path | What it does |
| --- | --- |
| `Views/`, `ViewModels/` | XAML pages + MVVM view-models (Login, Main) |
| `Services/ExchangeService.cs` | In-process PowerShell runspace; talks to Exchange Online |
| `Services/PowerAutomateService.cs` | Imports the generated Power Automate Solution and manages the cloud flow state |
| `Services/CloudSchedulePackageGenerator.cs` | Builds the Power Automate Solution zip for always-on cloud OOF scheduling |
| `Services/CloudScheduleGuideGenerator.cs` | Builds the personalized manual setup guide for the cloud flow |
| `Services/PreferencesService.cs` | Persists settings to `%LocalAppData%\OofManager\preferences.json` |
| `Services/TemplateService.cs` | Template library (SQLite via `sqlite-net-pcl`) |
| `Services/StartupService.cs` | Legacy HKCU `Run` cleanup for older local-background builds |
| `Modules/ExchangeOnlineManagement/` | Bundled PowerShell module shipped with the app |
| `Installer/OofManager.iss` | Inno Setup script |

## License

MIT — see [LICENSE](LICENSE).
