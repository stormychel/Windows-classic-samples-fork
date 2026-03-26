Dim xlApp, wb, addinPath, basPath, fso

' Automatically resolve SendToEmail.bas relative to this script's location
Set fso = CreateObject("Scripting.FileSystemObject")
basPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\SendToEmail.bas"
Set fso = Nothing
addinPath = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\AddIns\SendToEmail.xlam"

Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False
xlApp.DisplayAlerts = False

Set wb = xlApp.Workbooks.Add

' Import VBA module
wb.VBProject.VBComponents.Import basPath

' Save as xlAddIn format (55)
wb.SaveAs addinPath, 55

wb.Close False
xlApp.Quit

Set wb = Nothing
Set xlApp = Nothing

WScript.Echo "Add-in created at: " & addinPath
