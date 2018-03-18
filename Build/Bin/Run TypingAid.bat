REM Starting TypingAid without compilation. I can try to attach VS later.
cd C:\GitHub\TypingAid\Build\Bin
REM If we want cmd window to stay open after AutoHotkeyU64.exe exits
C:\Windows\System32\cmd.exe /k C:\Users\10114976\Downloads\Apps\AutoHotKey\AutoHotkey_1.1.28.00\AutoHotkeyU64.exe^
 ..\..\Source\TypingAid.ahk

REM Console window will close after AutoHotkeyU64.exe is killed; closing console will kill the app
C:\Users\10114976\Downloads\Apps\AutoHotKey\AutoHotkey_1.1.28.00\AutoHotkeyU64.exe^
 ..\..\Source\TypingAid.ahk