' run.vbs
' Usage:
'   cscript run.vbs                           ' default: run scriptRun.jsx in same folder as run.vbs
'   cscript run.vbs "C:\path\to\script.jsx" ' absolute path to jsx file
'   cscript run.vbs "myscript.jsx"            ' relative path (relative to current working directory)

Set fso = CreateObject("Scripting.FileSystemObject")
Set app = CreateObject("InDesign.Application")
Set shell = CreateObject("WScript.Shell")

' Check if an argument was passed
If WScript.Arguments.Count > 0 Then
    inputPath = WScript.Arguments(0)

    ' Check if inputPath is an absolute path:
    '   - Windows drive-letter path:  C:\...
    '   - UNC path:               \\server\share\...
    '   - Unix-style absolute:      /path/to/...
    If Mid(inputPath, 2, 2) = ":\" Or Left(inputPath, 2) = "\\" Or Left(inputPath, 1) = "/" Then
        scriptPath = inputPath
    Else
        ' Relative path: resolve against current working directory
        scriptPath = fso.BuildPath(shell.CurrentDirectory, inputPath)
    End If
Else
    ' No argument: use scriptRun.jsx in the same folder as run.vbs
    scriptPath = fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "scriptRun.jsx")
End If

' Check that the script file exists
If Not fso.FileExists(scriptPath) Then
    WScript.Echo "Error: script file not found: " & scriptPath
    WScript.Quit 1
End If

app.DoScript scriptPath, 1246973031