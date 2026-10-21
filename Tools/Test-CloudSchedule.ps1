#Requires -Version 5.1
param(
    [string]$AssemblyPath = (Join-Path $PSScriptRoot '..\bin\Release\net48\OofManager.exe')
)

$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSEdition -eq 'Core') {
    throw 'Run this script with Windows PowerShell (powershell.exe) for .NET Framework 4.8.'
}

$assembly = [Reflection.Assembly]::LoadFrom((Resolve-Path $AssemblyPath).Path)
$assemblyDirectory = Split-Path $assembly.Location
# The PowerShell host does not apply OofManager.exe.config binding redirects.
$resolver = [ResolveEventHandler] {
    param($sender, $eventArgs)
    $name = New-Object Reflection.AssemblyName($eventArgs.Name)
    $path = Join-Path $assemblyDirectory ($name.Name + '.dll')
    if (Test-Path -LiteralPath $path) {
        return [Reflection.Assembly]::LoadFrom($path)
    }
    return $null
}
[AppDomain]::CurrentDomain.add_AssemblyResolve($resolver)
$compiler = New-Object CodeDom.Compiler.CompilerParameters
$null = $compiler.ReferencedAssemblies.Add($assembly.Location)
Add-Type -CompilerParameters $compiler -TypeDefinition @'
using System;
using System.Collections.Generic;
using OofManager.Wpf.Services;

public sealed class CloudScheduleTestPreferences : IPreferencesService
{
    private readonly Dictionary<string, object> values = new Dictionary<string, object>();
    public bool GetBool(string key, bool fallback)
    {
        object value;
        return values.TryGetValue(key, out value) ? (bool)value : fallback;
    }
    public int GetInt(string key, int fallback)
    {
        object value;
        return values.TryGetValue(key, out value) ? (int)value : fallback;
    }
    public string GetString(string key, string fallback = null)
    {
        object value;
        return values.TryGetValue(key, out value) ? (string)value : fallback;
    }
    public void Set(string key, bool value) { values[key] = value; }
    public void Set(string key, int value) { values[key] = value; }
    public void Set(string key, string value) { values[key] = value; }
    public IDisposable BeginBatch() { return new Batch(); }
    private sealed class Batch : IDisposable { public void Dispose() { } }
}
'@

function Assert-Equal($Actual, $Expected, [string]$Message) {
    if ($Actual -cne $Expected) {
        throw "${Message}: expected '$Expected', got '$Actual'."
    }
}

function Read-Definition($Definition) {
    $json = @{ Outcome = 'Success'; Definition = $Definition } | ConvertTo-Json -Depth 100
    $method = [OofManager.Wpf.Services.PowerAutomateService].GetMethod('ParseDefinitionResult', [Reflection.BindingFlags]'NonPublic,Static')
    $result = $method.Invoke($null, @([string]$json, 0, 'OOF regression test'))
    Assert-Equal $result.Outcome.ToString() 'Success' 'Definition parsing'
    return $result
}

$prefs = New-Object CloudScheduleTestPreferences
foreach ($day in [Enum]::GetValues([DayOfWeek])) {
    $prefs.Set("WorkSchedule.$day", $day -notin @([DayOfWeek]::Saturday, [DayOfWeek]::Sunday))
}
$prefs.Set('WorkSchedule.StartMinutes', 9 * 60)
$prefs.Set('WorkSchedule.EndMinutes', 18 * 60)
$prefs.Set('WorkSchedule.Wednesday.EndMinutes', 17 * 60 + 30)
$schedule = New-Object OofManager.Wpf.Services.WorkScheduleSnapshot($prefs)
$zipPath = Join-Path ([IO.Path]::GetTempPath()) ("OofManager-Test-" + [Guid]::NewGuid() + '.zip')

try {
    $null = [OofManager.Wpf.Services.CloudSchedulePackageGenerator]::Generate(
        $schedule, 'test@example.com', 'Internal reply', 'External reply', $true, $true, $zipPath)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($zipPath)
    try {
        $entry = @($zip.Entries | Where-Object { $_.FullName -like 'Workflows/*.json' })
        Assert-Equal $entry.Count 1 'Workflow count'
        $reader = New-Object IO.StreamReader($entry[0].Open())
        try { $flow = $reader.ReadToEnd() | ConvertFrom-Json } finally { $reader.Dispose() }
    }
    finally { $zip.Dispose() }

    $definition = $flow.properties.definition
    $recurrence = $definition.triggers.Recurrence.recurrence
    Assert-Equal $recurrence.frequency 'Day' 'Daily frequency'
    Assert-Equal $recurrence.interval 1 'Daily interval'
    Assert-Equal (@($recurrence.schedule.hours) -join ',') '17' 'Earliest end hour'
    Assert-Equal (@($recurrence.schedule.minutes) -join ',') '30' 'Earliest end minute'
    Assert-Equal ($recurrence.schedule.PSObject.Properties.Name -contains 'weekDays') $false 'No workday filter'
    Assert-Equal $recurrence.timeZone ([TimeZoneInfo]::Local.Id) 'Local time zone'
    Assert-Equal $definition.parameters._oofmgr_source.defaultValue.days.Saturday.workday $false 'Saturday remains off'
    Assert-Equal $definition.parameters._oofmgr_source.defaultValue.days.Sunday.workday $false 'Sunday remains off'

    $parsed = Read-Definition $definition
    Assert-Equal (($parsed.TriggerDays | Sort-Object) -join ',') 'Sunday,Monday,Tuesday,Wednesday,Thursday,Friday,Saturday' 'Daily trigger coverage'
    Assert-Equal $parsed.PerDaySchedule[[DayOfWeek]::Saturday].IsWorkday $false 'Parsed workdays independent of trigger'

    # Check the generated action's lookup tables, not just the desktop schedule.
    $timestampBuilder = [OofManager.Wpf.Services.CloudSchedulePackageGenerator].GetMethod('BuildPerDowTimestampExpression', [Reflection.BindingFlags]'NonPublic,Static')
    $expectedStart = $timestampBuilder.Invoke($null, @(
        [int[]]@(-2, 0, 0, 0, 0, 0, -1),
        [int[]]@(18, 18, 18, 17, 18, 18, 18),
        [int[]]@(0, 0, 0, 30, 0, 0, 0), [TimeZoneInfo]::Local.Id))
    $expectedEnd = $timestampBuilder.Invoke($null, @(
        [int[]]@(1, 1, 1, 1, 1, 3, 2),
        [int[]]@(9, 9, 9, 9, 9, 9, 9),
        [int[]]@(0, 0, 0, 0, 0, 0, 0), [TimeZoneInfo]::Local.Id))
    $action = $definition.actions.'Set up automatic replies (V2)'.inputs.parameters
    Assert-Equal $action.'body/automaticRepliesSetting/status' 'scheduled' 'OOF restoration'
    Assert-Equal $action.'body/automaticRepliesSetting/scheduledStartDateTime/dateTime' ("@{$expectedStart}") 'Weekend start anchored to Friday end'
    Assert-Equal $action.'body/automaticRepliesSetting/scheduledEndDateTime/dateTime' ("@{$expectedEnd}") 'Weekend end anchored to Monday start'

    foreach ($date in @([datetime]'2026-10-24T17:30:00', [datetime]'2026-10-25T17:30:00')) {
        $window = $schedule.ComputeNextOffHoursWindow($date)
        Assert-Equal $window.Item1.DateTime ([datetime]'2026-10-23T18:00:00') 'Weekend desktop start'
        Assert-Equal $window.Item2.DateTime ([datetime]'2026-10-26T09:00:00') 'Weekend desktop end'
    }

    $recurrence.frequency = 'Week'
    $recurrence.schedule | Add-Member NoteProperty weekDays @('Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday')
    $legacy = Read-Definition $definition
    Assert-Equal $legacy.TriggerDays.Count 5 'Legacy workday-only trigger remains detectable'
    $recurrence.schedule.weekDays = @('Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday')
    Assert-Equal (Read-Definition $definition).TriggerDays.Count 7 'Legacy seven-day weekly trigger'
    $recurrence.schedule.PSObject.Properties.Remove('weekDays')
    $recurrence.frequency = 'Day'
    $recurrence.interval = 2
    Assert-Equal (Read-Definition $definition).TriggerDays.Count 0 'Every-other-day trigger is not daily'

    $guideBuilder = [OofManager.Wpf.Services.CloudScheduleGuideGenerator].GetMethod('Build', [Reflection.BindingFlags]'NonPublic,Static')
    $guide = $guideBuilder.Invoke($null, @($schedule.PSObject.BaseObject, 'test@example.com', 'Internal', 'External', $true))
    Assert-Equal ($guide -match 'runs every day at the earliest configured workday end time') $true 'Manual guide matches recurrence'
    Write-Host 'Cloud schedule regression checks passed.'
}
finally {
    [AppDomain]::CurrentDomain.remove_AssemblyResolve($resolver)
    if (Test-Path $zipPath) { Remove-Item -LiteralPath $zipPath }
}
