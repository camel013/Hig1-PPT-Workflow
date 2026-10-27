# Validation Protocol

## Retrieval

- Read the selected project index before opening project files.
- Read files by path and line range. Do not scan the entire workspace.
- Keep tool output under 8KB. Store long logs outside tracked directories and read only failure lines.
- Fetch LFS assets only for the selected project.

## Mechanical Validation

Check the expected slide count, slide size, fonts, text-box bounds, punctuation, image crop, package integrity, and whether the PPTX opens.

Run a mechanical check once per input state. Reuse the previous result when the command, inputs, artifact, and environment are unchanged.

## Visual Validation

- Embed fonts before generating the final preview.
- Render only changed pages during repair.
- For decks over eight pages, inspect a contact sheet first and zoom only suspect pages.
- Run one full visual pass on the final embedded artifact.

## Stop Conditions

Stop immediately when mechanical and final visual checks pass.

Stop and report when:

- the same failure signature remains after two repair attempts;
- a repair produces no state change;
- the tool-call or token budget is exhausted;
- the client blocks a repeated identical command.

Do not repeat an unchanged test, render, file read, or screenshot.
