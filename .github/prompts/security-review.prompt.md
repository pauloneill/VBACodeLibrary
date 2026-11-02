---
mode: agent
description: Security-focused review of the current changes
---

Review the current branch diff for security and integrity problems only. Do not comment on style.

This is a public library of VBA code that other people import into Access, Excel and Word projects that handle real business data, so check especially:
- Confidentiality of this repo: client names, client data, credentials, API keys, connection strings, server, share or personal file paths, email addresses, anything that identifies a client system. Check code, comments, examples, test data and README, and check that no host file (`.accdb`, `.mdb`, `.xlsm`, `.docm` and similar) or lock file is committed.
- Injection: SQL built from strings instead of parameters or correctly escaped and delimited values; text from users or files passed to `Shell`, `CreateObject`, `Eval`, `Application.Run` or macro execution; file paths built from input without validation (path traversal, UNC paths).
- Dangerous capabilities: anything that deletes, overwrites or moves files, writes the registry, calls `Shell` or `WScript.Shell`, makes network calls (`XMLHTTP`, `WinHttp`), or reads the clipboard or environment. Each must be obvious from the component's name and documented in its README, and never happen as a side effect of import or `Class_Initialize`.
- API declarations: wrong argument types or buffer sizes (buffer overruns), `Long` instead of `LongPtr`, handles that are not released.
- Error handling that hides failures from the caller or leaves the host in a changed state (warnings off, screen updating off, locked files).
- Anything resembling obfuscated code, auto-run macros (`Auto_Open`, `Document_Open`, `Workbook_Open`) or behaviour not described by the change.

For each finding give: file and line, severity (Important / Nit), a short explanation, and a suggested fix.
If there are none, say "No security findings." Do not edit files.
