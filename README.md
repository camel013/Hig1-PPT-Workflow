# Hig1 PPT Workflow

A compact, reproducible PPT workspace. The active tree contains the local skill, the design system, and project sources and deliverables. Generated previews, one-off patches, backups, and historical builds remain outside the active tree.

## Structure

```text
.
|-- AGENTS.md
|-- README.md
|-- skills/ppt-workflow/
|   |-- SKILL.md
|   |-- agents/
|   |-- references/
|   `-- scripts/
|-- design-system/
|   `-- examples/
|-- projects/
`-- .local/              # ignored local archive and runs
```

## Start a task

Read the compact project index instead of scanning the repository:

```powershell
pwsh -NoProfile -File .\skills\ppt-workflow\scripts\index_projects.ps1
pwsh -NoProfile -File .\skills\ppt-workflow\scripts\index_projects.ps1 -Project blues_shadows_ppt169_20261026
```

Create a run manifest only when starting a new deck:

```powershell
pwsh -NoProfile -File .\skills\ppt-workflow\scripts\preflight.ps1 -DeckName example -Pages 12
```

Verify locked content:

```powershell
pwsh -NoProfile -File .\skills\ppt-workflow\scripts\template_lock.ps1 -Mode Verify
```

## Repository use

For text-only work, skip LFS downloads:

```powershell
$env:GIT_LFS_SKIP_SMUDGE = "1"
git clone https://github.com/camel013/Hig1-PPT-Workflow.git
```

Fetch only the assets needed by the selected project:

```powershell
git lfs pull --include="projects/<project>/**"
```

## Operating rules

- Read the project index, then only relevant files.
- Reuse existing layout and generation scripts.
- Validate mechanical constraints first.
- Render changed pages only; check the final deck once.
- Stop after success or after the same failure remains through two repairs.
