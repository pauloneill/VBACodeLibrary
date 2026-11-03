# Project instructions for the AI agent

<!-- Keep this file SHORT. Every line is loaded into every request.
     When the reviewer or agent repeatedly gets something wrong, add ONE line here. -->

## Project
- VBA Code Library: a public, copy-and-import library of Microsoft Access and Word/Excel VBA classes and modules written over many years of business use. Other people (and my own projects) import these files, so the public surface is a contract.
- Layout: exported source only, under `src/<area>/` (`.bas` modules, `.cls` classes, `.frm` forms), with a `README.md` per area documenting each component (properties, methods, usage, limitations). The top-level `README.md` is the front page.
- Host: mostly Access VBA (`Option Compare Database`). Say so in the component's README when code only works in one host application (Access, Excel or Word).
- Goals: reusable, self-contained components that drop into any project with the fewest dependencies, and are documented well enough to use without reading the code.
- Non-goals: anything that needs a build step, an add-in or a third-party library; client-specific code.
- This repo is public. Never commit client names, client data, connection strings, credentials, server or network paths, or code that only makes sense inside one client's system.

## Conventions
- Every module and class starts with `Option Explicit`. Keep the existing header block (copyright, author, change log) intact. Add a change-log line in the existing style (date, initials, what changed) when changing a file.
- Keep the `Attribute VB_Name` equal to the file name. Do not change other `Attribute` lines (`VB_PredeclaredId`, `VB_Exposed`, `VB_GlobalNameSpace`) without asking: they change how the component is instantiated.
- Host-specific code (`CurrentDb`, `DoCmd`, `Application.Cursor`, `System.Cursor`) goes inside `#If InAccess` / `#If InExcel` / `#If InWord`, with a host-neutral `#Else` where possible. Do not add `Option Compare Database` (Access-only); use explicit `vbTextCompare`/`vbBinaryCompare`.
- Prefer built-in VBA and the object model of the host. Late-bind (`CreateObject`) rather than adding a reference, and say so in the README.
- API `Declare` statements must be `PtrSafe` and use `LongPtr` for handles and pointers, with `#If VBA7` / `Win64` guards where older Office matters.
- Handle errors deliberately: `On Error GoTo` with a handler that cleans up. Use `On Error Resume Next` only around one statement, then restore with `On Error GoTo 0`. Never swallow an error silently.
- Release what you open: recordsets, database and file handles, objects (`Set x = Nothing`), mouse pointer, status bar, screen updating, and `DoCmd.SetWarnings`. Restore state on every exit path including the error path.
- Never build SQL by concatenating raw values. Use parameters (`QueryDef`/`ADODB.Command`), or escape quotes and use proper delimiters for text, dates and numbers. Treat all text from users and files as untrusted.
- Dates in SQL are US `mm/dd/yyyy` or ISO `yyyy-mm-dd`, never a locale-formatted string.
- Name things consistently with the existing code in the same area (class prefix `c`, `Public`/`Private` stated explicitly, `ByVal`/`ByRef` stated explicitly).
- Prefer small, focused changes. Do not refactor unrelated code in the same change.
- Do not add a dependency or a new host requirement without asking me first.

## Verification (show the output before saying "done")
- Static checks: `sh scripts/check.sh` (also run by the pre-push hook and CI). It can only read the text; it cannot compile or run VBA.
- Compile and run in the host application: import the changed files into a clean project, run **Debug > Compile**, and exercise the change by hand or with the `Assert` handlers. You usually cannot do this in the agent environment. Say so explicitly, and say what it would have caught (compile errors, wrong host, runtime errors, behaviour changes).

Never report a task as complete without showing the result of the checks above.
If a check cannot be run, say so explicitly and say what it would have caught.
A change to a component's public members also needs its area `README.md` updated in the same change.

## Sensitive areas (stop and ask me before changing)
- The public surface of a component: names, parameters, return types and meaning of public properties and methods, and enum values. Consumers import these files by copy, so a change silently breaks projects I cannot see. Add a new member instead of changing an old one.
- `Attribute` lines and `Globals.bas`: they control instancing and the global `Assert` object.
- The `Assert` classes and handlers (`src/base/`): other components rely on their behaviour.
- Anything that deletes or overwrites data, files or registry settings on the user's machine.
- The header, copyright and licence text, and `LICENSE`.
- CI and release: `.github/workflows/`, `.githooks/`, `scripts/check.sh`.

## Review loop limits
- After 3 fix-and-review rounds on the same change, stop and re-evaluate the approach instead of continuing.
- On re-review, look only at the new diff unless I ask for a full pass.
- When fixing a flagged bug, fix the whole class of bug, not just the one line.
