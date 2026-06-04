Set WshShell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")
scriptDir = FSO.GetParentFolderName(WScript.ScriptFullName)
appPath = FSO.BuildPath(scriptDir, "app.pyw")
WshShell.Run """" & appPath & """", 0, False
