' ============================================================================
' build.vbs - creates InventoryTracker.accdb next to this script, imports the
' VBA modules from .\vba and runs BuildInventoryDatabase.
'
' Double-click it, or:   cscript //nologo build.vbs [C:\path\to\NewFile.accdb]
' Requires Microsoft Access (any 2010+ desktop version, 32 or 64 bit).
' ============================================================================
Option Explicit

Const acCmdCompileAndSaveAllModules = 126

Dim fso, here, dbPath, app, f, imported
Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)

dbPath = here & "\InventoryTracker.accdb"
If WScript.Arguments.Count > 0 Then dbPath = WScript.Arguments(0)

If fso.FileExists(dbPath) Then
    MsgBox "This file already exists:" & vbCrLf & dbPath & vbCrLf & vbCrLf & _
           "Delete or rename it first (or pass another path).", vbExclamation, "Inventory build"
    WScript.Quit 1
End If

Set app = CreateObject("Access.Application")
app.NewCurrentDatabase dbPath
app.Visible = True

imported = 0
For Each f In fso.GetFolder(here & "\vba").Files
    If LCase(fso.GetExtensionName(f.Name)) = "bas" Then
        app.VBE.ActiveVBProject.VBComponents.Import f.Path
        imported = imported + 1
    End If
Next

If imported = 0 Then
    MsgBox "No .bas files found in " & here & "\vba", vbCritical, "Inventory build"
    app.Quit
    WScript.Quit 1
End If

On Error Resume Next
app.RunCommand acCmdCompileAndSaveAllModules
If Err.Number <> 0 Then
    MsgBox "The VBA modules did not compile: " & Err.Description & vbCrLf & _
           "Open the VBA editor (Alt+F11), choose Debug > Compile to see the line.", vbCritical, "Inventory build"
    Err.Clear
End If
On Error GoTo 0

app.Run "BuildInventoryDatabase"

' Leave Access open for the user.
app.UserControl = True
