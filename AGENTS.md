# Workspace Instructions

## Routing

- For PPTX, slides, posters, or visual-layout work, read and use `skills/ppt-workflow/SKILL.md`.
- Start with `skills/ppt-workflow/scripts/index_projects.ps1`. Read only the selected project.

## Retrieval

- Do not read `.local/`, `.git/`, `.venv/`, `vendor/`, `backups/`, `__pycache__/`, `node_modules/`, or archived generated previews.
- Do not scan all of `projects/`. Read the selected project index, specification, and named entrypoints.
- Limit large command output before processing it. Use targeted search and line ranges.

## Execution

- Reuse existing templates and scripts. Put working output under `.local/runs/`.
- Make local changes incrementally. Do not rebuild a whole deck after a local fix.
- Run mechanical checks before visual checks.
- Embed fonts before the final preview, inspect a contact sheet first, and run one full visual pass on the final artifact.
- Retry the same failure at most twice. Stop after success or when no new state is produced.

## Locked Content

- The skill, design system, project specifications, generation scripts, final deliverables, and validation records are locked.
- Without an explicit request, do not edit, move, delete, re-export, or auto-save locked content.
- For an authorized change, name the scope, create a backup or Git tag, apply the smallest change, then regenerate and verify the template lock.

Detailed design and validation guidance lives under `skills/ppt-workflow/references/` and is read only when relevant.
