---
mode: agent
description: Generate an up-to-date explainer for one part of the codebase before changing it
---

I'm about to change the area described below. Write a short, accurate explainer so I understand it before editing.

Area: ${input:area:file, folder or component name (for example the Assert classes, cBusyPointer, Globals)}

Cover:
1. Purpose: what this component does and why it exists (link the relevant section of the area `README.md`)
2. Key files and how they connect (a small text diagram is fine), including which host application each part needs
3. Public surface: properties, methods, enums and instancing attributes that callers depend on
4. Conventions and gotchas a new reader would trip on (for example `As New` auto-instancing, `VB_PredeclaredId`, `Class_Terminate` timing, error-handler state)
5. What depends on this code: search the repo for callers, and note what could break in projects that have already imported the file
6. How it is tested today (assertions, manual steps) and gaps

Base everything on the actual code. Mark anything you are unsure about as uncertain.
