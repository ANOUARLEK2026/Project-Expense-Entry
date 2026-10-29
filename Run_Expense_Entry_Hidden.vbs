Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
base = fso.GetParentFolderName(WScript.ScriptFullName)
app = base & "\app.py"

If Not fso.FileExists(app) Then
    MsgBox "The application file was not found:" & vbCrLf & app, vbExclamation, "Project Expense Entry"
    WScript.Quit 1
End If

shell.CurrentDirectory = base

On Error Resume Next
Err.Clear
Set check = shell.Exec("py -3 -c ""import PySide6""")
If Err.Number <> 0 Then
    MsgBox "Python 3 was not found. Install Python 3.11 or newer and ensure the ""py"" launcher is available.", vbExclamation, "Project Expense Entry"
    WScript.Quit 1
End If
On Error GoTo 0

Do While check.Status = 0
    WScript.Sleep 100
Loop

If check.ExitCode <> 0 Then
    MsgBox "PySide6 is not installed for Python 3. Install the app requirements with:" & vbCrLf & vbCrLf & "py -3 -m pip install --user -r requirements.txt", vbExclamation, "Project Expense Entry"
    WScript.Quit 1
End If

On Error Resume Next
exitCode = shell.Run("pyw -3 """ & app & """", 0, True)
If Err.Number <> 0 Then
    MsgBox "Could not start Python without a console. Ensure the Python launcher (pyw) is installed.", vbExclamation, "Project Expense Entry"
    WScript.Quit 1
End If
On Error GoTo 0

If exitCode <> 0 Then
    MsgBox "Project Expense Entry exited with an error (code " & exitCode & ").", vbExclamation, "Project Expense Entry"
End If
