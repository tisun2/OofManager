# Feasibility: Manual-mode vacation → Power Automate flows

Branch: `research/manual-vacation-cloud-flows`
Date: 2026-05-25
Status: Research only — no code changes yet.

## What the user wants

Today, Manual mode in `MainViewModel` lets the user pick a vacation window
(`VacationStartDate/Time`, `VacationEndDate/Time`) and click the OOF toggle;
`StartVacationAsync` / `EndVacationAsync` push an EWS Scheduled block. The
local app must be running (or be re-launched within the window) for
`EndVacationAsync` to clean up. Schedule mode separately uses a cloud flow —
`OofManager Cloud Sync ({alias})` — imported as a Dataverse solution via
`PowerAutomateService.ImportCloudScheduleSolutionAsync`.

The four asks are:

1. **Manual-vacation flow.** When Manual mode has a start+end window, generate
   a cloud flow that *enacts* the OOF window from the cloud (no local PC
   required).
2. **Pause-schedule flow.** A second flow that, at vacation start, disables the
   existing Schedule-mode flow.
3. **Resume-schedule flow.** A third flow that, at vacation end, re-enables
   the Schedule-mode flow.
4. **Persist + clean up.** Save the vacation window so the user can cancel it
   later, which must delete all three flows above.

## Existing assets we can reuse

| Asset | Where | Reuse for |
|---|---|---|
| Solution-zip generator (Dataverse format, deterministic GUIDs) | `Services/CloudSchedulePackageGenerator.cs` | Need a sibling generator per new flow, or one solution with 3 workflows. |
| Dataverse `ImportSolutionAsync` + poll + `PublishAllXml` path | `Services/PowerAutomateService.cs` (see repo memory `oofmanager-powerautomate.md`) | Reuse verbatim — same env discovery, same Get-JwtToken, same tenant-blocked fallback (pac CLI). |
| Flow on/off via ProcessSimple (`Stop-Flow` / `Start-Flow`) | `EnableOofManagerFlowsAsync` in `PowerAutomateService.cs` | Already proves we can drive flow state from PowerShell — but here we want the *cloud flow itself* to call that API, which is a different surface (Power Automate Management connector). |
| Preferences store (`Vacation.Start`, `Vacation.End` already round-trip) | `Services/PreferencesService.cs`, `MainViewModel` lines ~2046–2055 | Ask #4 partly works already — we just need to add `Vacation.PauseFlowId`, `Vacation.OofFlowId`, `Vacation.ResumeFlowId` alongside. |

## Feasibility per ask

### Ask #1 — Manual vacation flow (cloud-side OOF on/off)

**Verdict: Feasible. Medium effort.**

Two scheduled flows packaged in one solution:
- Trigger: "Recurrence" → one-shot at `VacationStart` UTC. Action:
  `Set-Mailbox -AutoReplyConfiguration` via the **Office 365 Outlook**
  connector action `Set up automatic replies (V2)` — already in the same
  shared_office365 connection we use for the Schedule flow, so **no new
  connection reference** is required.
- Same pattern at `VacationEnd` to clear AutoReply.

Implementation work:
- Add `ManualVacationPackageGenerator` modelled on
  `CloudSchedulePackageGenerator`. Stamp deterministic `WorkflowId` per
  (alias, vacation-window-hash) so re-import = upgrade.
- Solution unique name: `OofManagerManualVacation_{alias}`. Display name:
  `OofManager Manual Vacation ({alias})`.
- Reuse `PowerAutomateService.ImportCloudScheduleSolutionAsync` verbatim by
  parameterising the solution path & expected display name.

Risks / gotchas:
- One-shot Recurrence flows with start/end in the past are immediately
  "Succeeded" with zero runs. Need to validate the window is still in the
  future before import, otherwise we get a silently-no-op flow.
- Time-zone: trigger times use the Recurrence trigger's `timeZone` parameter.
  Vacation date pickers are local — must serialise as the user's TZID, or
  convert to UTC and set `timeZone=UTC`.

### Ask #2 + #3 — Pause/resume the Schedule flow from a flow

**Verdict: Feasible but auth-coupled. Medium-high effort.**

The two control flows need to call **Power Automate Management**
connector → actions `Turn off flow` / `Turn on flow`, targeting the existing
Schedule flow by `Flow Name` (the runtime GUID we already cache as
`FlowName` after the first toggle — see memory file lines 60–62) and
`Environment Name`.

Pros:
- We already capture both ids in cache after the Schedule flow runs once. We
  can write them into the new flows at solution-generation time.
- Same trigger pattern as Ask #1 (one-shot Recurrence at start/end). Pause
  and Resume can ship as two workflows in the same solution as Ask #1, so
  we end up with **one** solution import covering #1+#2+#3.

Cons / risks:
- Power Automate Management connector requires its own connection reference
  → user gets a second consent prompt during import. Not blockable, but it's
  another sign-in dialog.
- On Microsoft mega-tenant the M365 PA Management connector is sometimes
  restricted by DLP policy ("Business" vs "Non-business" data group). Need
  to verify on `tisun@microsoft.com`'s default env before committing.
- If the Schedule flow runtime `FlowName` GUID is **not yet cached** (user
  hasn't toggled the Schedule flow at least once), we don't know what id to
  hard-code into the pause/resume flows. Mitigations:
  - Force a `Get-Flow` lookup at vacation-plan time and fail with a clear
    "Toggle the Schedule flow once first" error.
  - Or use the *display-name* fallback inside the flow: a "List Flows as
    Admin" action filtered by display name, then "Turn off". Adds a second
    Management connector action; runtime is bounded by env size.
- If the user reimports the Schedule flow after planning the vacation, the
  workflow id is stable (deterministic v5 GUID by design — see
  `CloudSchedulePackageGenerator` lines 51–53), so cached `FlowName` remains
  valid. But runtime `FlowName` is **not the same** as the deterministic
  WorkflowId after import — see memory file line 3. The control flows must
  store the *runtime* `FlowName`, refreshed at plan-time.

### Ask #4 — Persist vacation + cancellation deletes the flows

**Verdict: Feasible. Low-medium effort.**

Persistence is half-done. `MainViewModel` already round-trips
`Vacation.Start` / `Vacation.End` through `PreferencesService`
(see lines 2046–2055).

New work:
- Extend `PreferencesService` keys: `Vacation.SolutionUniqueName`,
  `Vacation.EnvironmentId`, `Vacation.OofWorkflowId`,
  `Vacation.PauseWorkflowId`, `Vacation.ResumeWorkflowId`.
- Add `IPowerAutomateService.DeleteVacationSolutionAsync(envId, solutionUniqueName)`:
  Dataverse `DELETE /solutions({id})` after lookup by `uniquename`.
  Deleting the solution removes all 3 contained workflows in one call —
  cheaper and atomic vs. 3 individual flow deletes.
- UI: a "Cancel planned vacation" button in the Manual panel, enabled iff
  `Vacation.SolutionUniqueName` is present in prefs. On click → confirm →
  call delete → clear prefs.
- Edge case: if the user already passed `VacationStart` (flow already ran &
  AutoReply is on), cancellation must *also* turn AutoReply off via the
  existing local EWS path; otherwise the user appears still on vacation.

## Recommended packaging

Ship the three new flows as a **single** Dataverse solution
`OofManagerManualVacation_{alias}` containing:

```
Workflows/
  OofManager-Vacation-Start-{guid}.json   # set AutoReply, turn off Schedule flow
  OofManager-Vacation-End-{guid}.json     # clear AutoReply, turn on Schedule flow
```

Two workflows beats three: each workflow already needs both a recurrence
trigger and a "turn off / turn on Schedule flow" action, so merging
{AutoReply-on, Schedule-off} into one start-time workflow and {AutoReply-off,
Schedule-on} into one end-time workflow is cleaner, halves the trigger count,
and still gives us the "delete one solution" cleanup story for #4.

## Open questions to validate before coding

1. Does the **Power Automate Management** connector load in `tisun@microsoft.com`'s
   default env without DLP block? (Manual test in maker portal — 5 min.)
2. Confirm runtime `FlowName` for the Schedule flow is present in prefs
   after a single successful `EnableOofManagerFlowsAsync` (it should be —
   memory file line 60). If absent, the plan-vacation UX must surface a
   "toggle Schedule flow once first" gate.
3. For #1 the AutoReply connector action wants `InternalReply` / `ExternalReply`
   strings — we already have these in `OofSettings`. Confirm the action
   supports start/end time fields (it does: `StartTime`, `EndTime`,
   `ExternalAudience`).
4. Solution-version bumping: same dynamic `1.yy.Mdd.HHmm` scheme as the
   Schedule package (see memory file line 50) so re-plans upgrade in place.

## Effort estimate

| Piece | Rough size |
|---|---|
| `ManualVacationPackageGenerator` (mirror of existing) | ~400 LOC |
| `PowerAutomateService.ImportManualVacationSolutionAsync` + `DeleteVacationSolutionAsync` | ~200 LOC (most paths reused) |
| `MainViewModel` plan/cancel commands + plumbing | ~150 LOC |
| XAML: "Plan in cloud" + "Cancel planned vacation" buttons | small |
| Tests / manual validation on mega-tenant | the long pole |

## Verdict

All four asks are feasible and align with the existing Cloud-Sync
architecture. Recommended next step: spike Open Question #1 (DLP policy on
Power Automate Management connector) before writing the generator — that's
the one finding that could change the design from "two flows in one
solution" to "Azure Logic App with a service-principal Dataverse
connection", which is a much bigger lift.
