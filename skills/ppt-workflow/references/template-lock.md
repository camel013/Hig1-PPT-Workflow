# Template Lock

Locked files include workspace instructions, this skill, the design system, project specifications, project scripts, final deliverables, and validation records.

Do not edit, move, delete, or re-export a locked file without an explicit user request.

Before an authorized change:

1. Name the affected files and expected result.
2. Create a Git tag and a local archive of generated or ignored material.
3. Change only the approved scope.
4. Regenerate and verify the lock manifest.

Run:

```powershell
pwsh -NoProfile -File .\skills\ppt-workflow\scripts\template_lock.ps1 -Mode Verify
pwsh -NoProfile -File .\skills\ppt-workflow\scripts\template_lock.ps1 -Mode Write
```
