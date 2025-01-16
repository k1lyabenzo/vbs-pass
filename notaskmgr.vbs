Option Explicit
If AppPrevInstance() Then 
    MsgBox "There is an existing proceeding !" & VbCrLF &_
    CommandLineLike(WScript.ScriptName),VbExclamation,"There is an existing proceeding !"    
    WScript.Quit   
Else
Do
    Call KillProcess("TaSkMgR.ExE")
Loop
End If
'*************************************************************************************
Sub KillProcess(MyProcess)
Dim colProcessList,objProcess
Set colProcessList = GetObject("Winmgmts:").ExecQuery ("Select * from Win32_Process")
For Each objProcess in colProcessList
	On Error Resume Next
    If LCase(objProcess.name) = LCase(MyProcess) then
        objProcess.Terminate()
    End if
	On Error GoTo 0
Next
End Sub
'*********************************************************************************************
Function AppPrevInstance()   
    With GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")   
        With .ExecQuery("SELECT * FROM Win32_Process WHERE CommandLine LIKE " & CommandLineLike(WScript.ScriptFullName) & _
            " AND CommandLine LIKE '%WScript%' OR CommandLine LIKE '%cscript%'")   
            AppPrevInstance = (.Count > 1)   
        End With   
    End With   
End Function    
'*********************************************************************************************
Function CommandLineLike(ProcessPath)   
    ProcessPath = Replace(ProcessPath, "\", "\\")   
    CommandLineLike = "'%" & ProcessPath & "%'"   
End Function