---
name: ppt-workflow
description: Create, edit, inspect, or validate PPTX decks, slides, posters, and visual-layout documents in this workspace.
---

# PPT Workflow

Use the existing design system and project assets. Keep retrieval narrow and validation bounded.

## Start

1. Run the project index:

   ```powershell
   pwsh -NoProfile -File .\skills\ppt-workflow\scripts\index_projects.ps1 -Project "<project>"
   ```

2. Read only that project's `README.md`, design specification, and the entry scripts named by the index.
3. For new slide design, read [references/design-rules.md](references/design-rules.md).
4. Before validation, read [references/validation-protocol.md](references/validation-protocol.md).
5. Use `scripts/preflight.ps1` to create a compact run manifest for a new deck.

Do not scan `.local/`, `.git/`, job archives, previews, or all projects. Fetch only the LFS files named by the selected project index.

## Work

- Reuse project scripts and assets. Modify the smallest affected source or build unit.
- Put generated working files under `.local/runs/<project>_<timestamp>/`.
- Keep final deliverables under the project's `deliverables/` directory when the project already uses one.
- Preserve fixed project files unless the user explicitly authorizes a change.
- Read [references/template-lock.md](references/template-lock.md) before changing locked material.

## Validate

- Run mechanical checks first.
- Render changed pages only.
- Generate a contact sheet for decks over eight pages and inspect the final deck once.
- Stop after all required checks pass. Retry a failing repair at most twice.
- Return the artifact path, checks performed, and any remaining limitation.
