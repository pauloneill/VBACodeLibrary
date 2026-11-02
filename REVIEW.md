# REVIEW.md: what the reviewer should care about

<!-- Brief on purpose. Add a line whenever the reviewer misses something or nitpicks something. -->

## What counts as a failure in this project
- A change to the public surface of a component (names, parameters, return types, enum values, instancing `Attribute` lines) with no note in the area `README.md`, or one that would silently break projects that already imported the file
- Code that will not compile: missing `Option Explicit`, undeclared variables, non-`PtrSafe` `Declare` statements, `Long` used for a handle or pointer that must be `LongPtr`, a reference to a library the target project will not have
- Host assumptions: Access-only code (`CurrentDb`, `DoCmd`, `Forms!`) in a component documented or placed as host-neutral, or Excel/Word objects used without saying so
- Errors swallowed silently, `On Error Resume Next` left active beyond one statement, or an error path that skips cleanup
- State not restored on every exit path: mouse pointer, status bar, `ScreenUpdating`, `SetWarnings`, open recordsets, file handles, objects not set to `Nothing`, circular references that stop `Class_Terminate` running
- SQL built from raw strings (quote injection, broken dates and numbers), text from users or files used unvalidated, file paths built from input
- Client names, client data, credentials, connection strings, server or personal paths in code, comments, examples or README (the repo is public)
- Wrong `ByRef`/`ByVal`, `Variant` used where a type is known, implicit type conversions or `Null` handling that fails on real data
- Header block, copyright or change log removed or not updated; `VB_Name` not matching the file name
- Docs (area `README.md`, comments) that say something the code does not do

## Check that the change
1. Does what the PR description says
2. Fits the conventions in `.github/copilot-instructions.md`
3. Does not break other callers: search the repo for uses of anything whose signature or behaviour changed (for example `Assert` and its handlers are used across `src/`)
4. Updates the area `README.md` when public members change, and adds or changes a change-log line in the file header
5. For a bug fix: addresses the root cause, not just the symptom
6. Has been compiled (**Debug > Compile**) and exercised in the host application, and says which host and Office version. If not, that is stated and the risk rated accordingly

## Do NOT report
- Pure style or formatting
- Naming preferences, unless misleading
- Line-ending differences, unless a file is mixed CRLF/LF
- Issues that already existed before this change, unless you label them **Pre-existing**

## Output format
- Every finding needs a **file path and line number**. Never guess filenames.
- Label each finding: **Important**, **Nit**, or **Pre-existing**. Order by severity.
- Start with a one-line summary: how many findings, and by severity.
- If there are no issues, say clearly: "No issues found."
- When flagging a bug, name the class of bug so all instances can be fixed together.

## Recurring issues (promote to a check in `scripts/check.sh` or a line in copilot-instructions.md once it repeats)
- <date>: <issue the reviewer keeps flagging>
