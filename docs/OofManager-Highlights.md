# OOF Manager

**OOF Manager** is a lightweight Windows application for managing Microsoft 365 and Exchange Online automatic replies. It combines flexible work schedules, vacation automation, and cloud-based execution in one focused interface.

## Highlights

### Flexible Weekly Scheduling

- Configure working days independently.
- Set different start and end times for each day.
- Automatically schedule OOF from the end of one workday until the next working period begins.
- Trigger cloud updates once every day, including non-workdays, at the earliest configured workday end time.
- Restore OOF after temporary non-workday overrides until the next configured workday starts, without marking non-workdays as working days.

### Works Even When Your PC Is Off

OOF Manager can generate and import a Power Automate solution that runs entirely in Microsoft 365. Once configured, scheduled replies continue working without Outlook, OOF Manager, or your computer being open.

### Vacation Automation

Schedule a vacation window in advance and let OOF Manager:

1. Apply the vacation reply at the selected start time.
2. Pause the regular weekly schedule.
3. Clear the vacation period automatically at the end.
4. Restore the normal off-hours reply and resume the weekly flow.

### Rich-Text Replies

Create polished internal and external replies with:

- Font family, size, and color
- **Bold**, *italic*, and underline
- Hyperlinks
- Bulleted and numbered lists

Formatting is preserved when synchronizing with Exchange Online and Power Automate.

### Internal and External Audiences

Maintain separate messages for colleagues and external senders, making it easy to share detailed internal context while keeping customer-facing replies concise.

### Cloud Drift Detection

Use **Compare with cloud** to identify differences between the local configuration and the deployed Power Automate flow, including:

- Working days and per-day hours
- Daily trigger coverage, independently of working days
- Trigger time
- Reply content
- External audience settings

### Direct Exchange Online Integration

OOF Manager reads and writes mailbox automatic-reply settings directly through Exchange Online. It does not depend on the Outlook desktop application.

### Simple Windows Experience

- Guided Microsoft 365 sign-in
- Reusable reply templates
- Manual one-click synchronization
- Optional startup with Windows
- Installer bundles the required Exchange Online and Power Platform components

> OOF Manager is designed for people with non-standard schedules, hybrid work patterns, or recurring after-hours coverage who need reliable automatic replies without manually updating Outlook every day.
