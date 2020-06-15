Option Explicit

Dim language
Dim messages(4)

Dim backupfile
Dim record
Dim myshell
Dim appmyshell
Dim myresult
Dim myline
Dim makeactive
Dim makepassive
Dim reboot

Set myshell = WScript.CreateObject("WScript.Shell")
record = ""
messages(0) = "Hypervisor status is passive, do you want set active?"
messages(1) = "Hypervisor status is active, do you want set to passive?"
messages(2) = "Not Changed"
messages(3) = "Error: record can't find"
language = "es"
' language = MsgBox("Select Your Language?", vbDefaultButton3, "Hypervisor")
Select Case language
    Case "es"
        messages(0) = "HyperV esta desactivado, desea activarlo?"
        messages(1) = "HyperV esta activado, desea desactivarlo?"
        messages(2) = "No hubieron cambios!!"
        messages(3) = "No se encontro el estado de HyperV"
        End Select


If WScript.Arguments.Length = 0 Then
    Set appmyshell  = CreateObject("Shell.Application")
    appmyshell.ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ RunAsAdministrator", , "runas", 1
    WScript.Quit
End if

Set backupfile = CreateObject("Scripting.FileSystemObject")
If Not (backupfile.FileExists("C:\bcdedit.bak")) Then
    Set myresult = myshell.Exec("cmd /c bcdedit /export c:\bcdedit.bak")
End If

Set myresult = myshell.Exec("cmd /c bcdedit")
Do While Not myresult.StdOut.AtEndOfStream
    myline = myresult.StdOut.ReadLine()
    If myline="The boot configuration data store could not be opened." Then
        record=""
        exit do
    End If

    If Instr(myline, "identifier") Or Instr(myline, "Identificador") > 0 Then
        record=""
        If Instr(myline, "{current}") > 0 Then
            record="current"
        End If
    End If

    If Instr(myline, "hypervisorlaunchtype") > 0 And record = "current" Then
        If Instr(myline, "Auto") > 0 Then
            record="1"
            Exit Do
        End If

        If Instr(myline, "On") > 0 Then
            record="1"
            Exit Do
        End If

        If Instr(myline, "Off") > 0 Then
            record="0"
            Exit Do
        End If
    End If
Loop

If record="1" Then
    makepassive = MsgBox (messages(1), vbYesNo, "Hypervisor")
    Select Case makepassive
    Case vbYes
        myshell.run "cmd.exe /C  bcdedit /set hypervisorlaunchtype off"
        reboot = MsgBox ("Hypervisor chenged to passive; Computer must reboot. Reboot now? ", vbYesNo, "Hypervisor")
        Select Case reboot
            Case vbYes
                myshell.run "cmd.exe /C  shutdown /r /t 0"
        End Select
    Case vbNo
        record = MsgBox (messages(2), vbOkOnly, "Hypervisor")
        End Select
End If

If record="0" Then
    makeactive = MsgBox (messages(0), vbYesNo, "Hypervisor")
    Select Case makeactive
    Case vbYes
        myshell.run "cmd.exe /C  bcdedit /set hypervisorlaunchtype auto"
        reboot = MsgBox ("Hypervisor changed to active;  Computer must reboot. Reboot now?", vbYesNo, "Hypervisor")
        Select Case reboot
            Case vbYes
                myshell.run "cmd.exe /C  shutdown /r /t 0"
        End Select
    Case vbNo
        record = MsgBox (messages(2), vbOkOnly, "Hypervisor")
    End Select
End If

If record="" Then
    record = MsgBox (messages(3), vbOkOnly, "Hypervisor")
End If