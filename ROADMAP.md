# Roadmap

Ideas and planned work, roughly in priority order. Move an item to `README.md` or delete it when it ships.

## Planned

- **vba-blocks manifest**: load all components through `vba-block.toml`. First pass exists; still to do is a real `vba build`, confirming whether an Access (`accdb`) target is supported, and a separate `[project]` manifest for a test host.
- **Fix `.accdb` / `.accde` handling in the assert classes**: `CAssert` and `cAssertHandler` only recognise `.mdb` and `.mde`, so assertions are silently ignored in modern Access files.

## Under consideration

- **Runtime host detection instead of conditional compilation constants** (replaces `InAccess` / `InExcel` / `InWord`):
  - Detect the host with `Application.Name` and call host-specific members late-bound (for example via `CallByName`), so projects need no Conditional Compilation Arguments.
  - Fixes the failure where forgetting `InAccess = 1` in an Access project makes `cBusyPointer` silently do nothing.
  - Trade-offs: loses compile-time checking of the host calls, and late-bound code is harder to read.
  - Constraint: `#Const` is module-scoped in VBA, so putting the constants in `Globals.bas` does not work; the current options are project-level arguments or a per-file `#Const`.
