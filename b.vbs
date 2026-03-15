Set WshShell = CreateObject("WScript.Shell")
Set WshNetwork = CreateObject("WScript.Network")

' Bilgisayarı yeniden başlat
WshShell.Run "shutdown /r /t 5", 0, False

' Belleği temizle
Set WshShell = Nothing
Set WshNetwork = Nothing
Set fso = Nothing
