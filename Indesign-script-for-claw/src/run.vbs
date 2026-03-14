Set app = CreateObject("InDesign.Application")
Set fso = CreateObject("Scripting.FileSystemObject")
scriptPath = fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "scriptRun.jsx")

app.DoScript scriptPath, 1246973031