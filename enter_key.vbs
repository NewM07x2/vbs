Option Explicit

Dim WSHShell
Set WSHShell = WScript.CreateObject("WScript.Shell")
Dim strCmd
strCmd = "cmd /c Get-AzVM -Status > C:\Users\DUKeDev0044\mnitta\devlop\vbs\Azure.txt"
WshShell.Run strCmd, 0, false
' WSHShell.Run "cmd /c ipconfig /all > C:\Users\DUKeDev0044\mnitta\devlop\vbs\ipconfig.txt", 0, false
'スタートアップからの起動時なので、待機を10秒と長めに設定

' WScript.Sleep 100

WSHShell.SendKeys("{Enter}")             
Set WSHShell = Nothing


