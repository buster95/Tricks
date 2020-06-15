Option Explicit
Dim myshell
Set myshell = WScript.CreateObject("WScript.Shell")

' myshell.run "regedit.exe /S putty_config.reg"
' myshell.run "putty.exe -ssh root@192.168.52.130 -pw Temporal%123"
myshell.run "putty.exe -ssh dgtiadmin@192.168.52.130 -pw Temporal%123"
' myshell.run "putty.exe -ssh wcorrales@172.16.2.60"