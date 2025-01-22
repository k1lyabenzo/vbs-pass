Dim userInput
Set shell = CreateObject("WScript.Shell")
set FSO=createobject("scripting.filesystemobject")

Do
    userInput = InputBox("Введите пароль для продолжения:", "Требуется пароль")
    If userInput = "123321" Then
		shell.run"taskkill /f /im wscript.exe",0 
		shell.run """explorer.exe" /TV""
        Exit Do
    Else
        MsgBox "Неверный пароль. Пожалуйста, попробуйте снова."
    End If
Loop
