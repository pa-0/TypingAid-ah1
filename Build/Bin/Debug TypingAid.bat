REM Starting in Debugging mode - use it to debug with Notepad++ (or another DBGp client)
cd C:\GitHub\TypingAid\Build\Bin

REM If we want cmd window to stay open after AutoHotkeyU64.exe exits:

REM C:\Windows\System32\cmd.exe /k C:\Users\10114976\Downloads\Apps\AutoHotKey\AutoHotkey_1.1.28.00\AutoHotkeyU64.exe^
REM  /Debug  ..\..\Source\TypingAid.ahk	
 
REM Console window will close after AutoHotkeyU64.exe is killed; closing console will NOT kill the app

C:\Users\10114976\Downloads\Apps\AutoHotKey\AutoHotkey_1.1.28.00\AutoHotkeyU64.exe^
   /Debug ..\..\Source\TypingAid.ahk
