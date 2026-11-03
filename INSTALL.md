# Installing the library

The components are plain exported VBA source (`.bas` modules, `.cls` classes) under `src/`. Nothing here is compiled or installed for you: you import the files into your own project and tell the compiler which host application it is running in.

## 1. Import the files

1. Open the VBA editor (`Alt+F11`) in your Access database, Excel workbook or Word document.
2. **File > Import File...** and select the files you need from `src/<area>/`. Each area's `README.md` lists what its components need.
3. For `src/base/`, import `CAssert.cls`, `cAssertHandler.cls`, `cSilentAssertionHandler.cls`, `cBusyPointer.cls`, then either import `Globals.bas` or copy its contents into your own global module. Importing it creates a module named `Globals`, which will clash if you already have one.

## 2. Set the compiler directives (once per project)

Host-specific code is switched on by conditional compilation constants. VBA has no built-in constant for the host application, so you declare it yourself, and it is **stored in the project, not in the imported files**. Repeat this for every project you import the library into.

1. In the VBA editor choose **Tools > *YourProjectName* Properties...**
2. On the **General** tab, find **Conditional Compilation Arguments**.
3. Enter the constant for your host application:

| Host | Enter | What it enables |
|------|-------|-----------------|
| Access | `InAccess = 1` | `CurrentDb`-based assertion handling, `DoCmd.HourGlass` busy pointer |
| Excel | `InExcel = 1` | `Application.Cursor` busy pointer |
| Word | `InWord = 1` | `System.Cursor` busy pointer |

To set more than one, separate them with a colon: `InAccess = 1 : InExcel = 0`. Leave any you do not need unset.

4. Click **OK**, then run **Debug > Compile *YourProjectName***.

If you set none, the library still compiles and runs in any host, but takes the host-neutral path: assertions always offer to debug via a message box, and `cBusyPointer` does nothing.

## 3. Check it worked

Paste this into any module and run it from the Immediate window (`InstallCheck`):

```vb
Public Sub InstallCheck()
#If InAccess Then
    Debug.Print "InAccess is set"
#ElseIf InExcel Then
    Debug.Print "InExcel is set"
#ElseIf InWord Then
    Debug.Print "InWord is set"
#Else
    Debug.Print "No host constant is set: running host-neutral"
#End If
End Sub
```

Then fire an assertion: `Assert.Assert False, "install check"`. You should get a message box offering to debug it.

## Known limitations

- The constants are per project and do not travel when you import files or copy a database. A new project needs step 2 again.
- In Access, `CAssert` and `cAssertHandler` only recognise `.mdb` and `.mde` files. Assertions are ignored in `.accdb` / `.accde` databases (tracked in [ROADMAP.md](ROADMAP.md)).
- Whether the [vba-blocks](https://www.vba-blocks.com/) manifest (`vba-block.toml`) can set these arguments for a built project has not been verified.
