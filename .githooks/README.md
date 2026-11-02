# Git hooks

Local hooks that run the same checks as CI before a push.

## Enable (once per clone)

```
git config core.hooksPath .githooks
```

On macOS/Linux you may also need `chmod +x .githooks/pre-push`. On Windows, Git for Windows runs the hook with its bundled shell, so nothing extra is needed.

## What `pre-push` does

- Runs `scripts/check.sh` (static checks on the VBA source: `Option Explicit`, `VB_Name` matches the file name, `PtrSafe` on API declarations, no conflict markers, no mixed line endings, no credentials or client paths, no host files).
- Only runs when the push changes files under `src/` or `scripts/` (docs-only pushes are skipped).
- Cancels the push if a check fails.

It does **not** compile or run the VBA. That needs Access, Excel or Word, so it stays a manual step: import the changed files, run **Debug > Compile**, and exercise the change (see `REVIEW.md`).

## Skip it (emergencies only)

```
git push --no-verify
```

## Why this exists

CI runs the same script, but a hook catches problems before they are pushed. This repo is public, so GitHub Free can enforce branch protection: consider requiring the `static-checks` CI job to pass before merging into `main`.
