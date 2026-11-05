# Test workbook

An Excel workbook (`.xlsm`) used to compile and exercise the library modules. VBA cannot be compiled outside Office, so this is where the "compile and run in the host application" step from `REVIEW.md` happens for Excel. (Access-only behaviour, such as `.mdb` handling, needs an Access test database, assembled by hand: vba-blocks does not support an `accdb` target.)

## Contents

- `vba-block.toml`: project manifest. Lists the library modules under `[dependencies]` (add one as each module is worked through).
- `src/TestBase.bas`: smoke tests for `base`.
- the working `.xlsm` (once it exists): see "Committing the workbook" below.

## Build the workbook with vba-blocks (unverified)

1. Install [vba-blocks](https://www.vba-blocks.com/) and, in Excel, enable **Trust access to the VBA project object model** (File > Options > Trust Center > Trust Center Settings > Macro Settings).
2. From this folder run `vba build`.
3. Open the generated workbook.

## Or assemble it by hand

1. Create a new macro-enabled workbook (`.xlsm`) and open the VBA editor (`Alt+F11`).
2. Import the module files you are testing from `src/<module>/`, plus `test/src/TestBase.bas`.
3. Set the compiler directive (it is stored in the project, so repeat it for every new workbook): **Tools > VBAProject Properties > General > Conditional Compilation Arguments**, enter `InExcel = 1`. See [INSTALL.md](../INSTALL.md).
4. **Debug > Compile VBAProject** must be clean.

## Run the tests

In the Immediate window (`Ctrl+G`) run `RunBaseTests`. Expected output:

```
--- base tests ---
PASS  failed assertion is captured
PASS  true assertion does not fire
--- done ---
```

Record the Excel and Office version you used in the PR description.

## Committing the workbook

`.gitignore` blocks Office files everywhere except `test/*.xlsm`, because this repo is public. Before committing a workbook, check that it contains no client data, that **File > Info > Inspect Document** shows no personal information or file paths, and that the macros in it are only the library and the tests. Prefer committing the manifest and sources and treating the workbook as a build output when vba-blocks can build it.
