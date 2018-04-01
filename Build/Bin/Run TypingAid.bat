REM Kill previous task, if running
taskkill /F /IM AutoHotkeyU64.exe

REM Starting TypingAid without compilation. I can try to attach VS later.
cd D:\GitHub\TypingAid\Build\Bin

REM If we want cmd window to stay open after AutoHotkeyU64.exe exits: 

REM C:\Windows\System32\cmd.exe /k D:\Downloads\Apps\AutoHotKey\AutoHotkey_1.1.28.00\AutoHotkeyU64.exe^
REM  ..\..\Source\TypingAid.ahk

REM Without 'start' the console window would close after AutoHotkeyU64.exe is killed; 
REM With 'start', console window will close right away.

REM start^
REM  D:\Downloads\Apps\AutoHotKey\AutoHotkey_1.1.28.00\AutoHotkeyU64.exe^
REM  ..\..\Source\TypingAid.ahk

REM Run AutoHotkeyU64.exe with limited affinity: only 1 CPUs out of 8, #8: mask 1000 0000 = 256 = 0x100 (mask 0000 0011 = 3 = 0x3)
REM
start /AFFINITY 0x100^
 D:\Downloads\Apps\AutoHotKey\AutoHotkey_1.1.28.00\AutoHotkeyU64.exe^
 ..\..\Source\TypingAid.ahk

REM  pause