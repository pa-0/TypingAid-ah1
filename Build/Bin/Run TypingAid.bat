REM Kill previous task, if running
taskkill /F /IM AutoHotkeyU64.exe

REM Starting TypingAid without compilation. I can try to attach VS later.
cd C:\GitHub\TypingAid\Build\Bin

REM If we want cmd window to stay open after AutoHotkeyU64.exe exits: 

REM C:\Windows\System32\cmd.exe /k C:\Users\10114976\Downloads\Apps\AutoHotKey\AutoHotkey_1.1.28.00\AutoHotkeyU64.exe^
REM  ..\..\Source\TypingAid.ahk

REM Console window will close after AutoHotkeyU64.exe is killed; with start, console window will close right away.

REM start^
REM  C:\Users\10114976\Downloads\Apps\AutoHotKey\AutoHotkey_1.1.28.00\AutoHotkeyU64.exe^
REM  ..\..\Source\TypingAid.ahk

REM Run AutoHotkeyU64.exe with limited affinity: only 2 CPUs out of 8, 0 and 1, mask 0000 0011 = 3 = 0x3
REM
start /AFFINITY 0x3^
 C:\Users\10114976\Downloads\Apps\AutoHotKey\AutoHotkey_1.1.28.00\AutoHotkeyU64.exe^
 ..\..\Source\TypingAid.ahk

REM  pause