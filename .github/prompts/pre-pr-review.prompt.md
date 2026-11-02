---
mode: agent
description: Fresh-eyes review of the current changes before opening a PR
---

You are reviewing my branch changes as an independent reviewer. Assume you did not write this code. Use a fresh chat with no memory of the build phase.

1. Read `.github/copilot-instructions.md` and `REVIEW.md` and follow `REVIEW.md` strictly.
2. Get the diff against the target branch (`git diff prototype...HEAD`, or `main...HEAD` if the PR targets main) plus any uncommitted changes. Read the surrounding code, not just the changed lines.
3. State, in one sentence, what the change is trying to do (use the PR description or the linked README section). Check it actually does that.
4. Run `sh scripts/check.sh` and show the real output. You cannot compile or run VBA here: say so, and list what a compile and a run in the host application would be expected to catch for this diff.
5. Look for: compile errors, undeclared variables, non-`PtrSafe` declarations, public signature or enum changes, host-specific code, error handling and cleanup mistakes (state not restored on every exit path, leaked recordsets, handles and objects), SQL built from strings, unhandled `Null`/empty/edge-case data, broken callers (search `src/` for uses), client-identifying content, header and change-log omissions, README that no longer matches, root cause vs symptom fixes.
6. Report findings in the format `REVIEW.md` specifies (file and line, severity label, summary count first).
7. Finish with a risk rating (Low / Medium / High) and a draft "How to revert" for the PR template, including whether projects that already imported the file need re-importing.

Do not edit any files. Only report.
