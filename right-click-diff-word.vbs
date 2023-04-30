'in order to hide the cmd terminals, invoking command with wscript
set ws=CreateObject("WScript.Shell")
set args=WScript.Arguments
'cmdline="cmd.exe /c C:/Users/" & ws.ExpandEnvironmentStrings("%username%") & "/AppData/Local/Microsoft/WindowsApps/right-click-diff-word.cmd " _
'& chr(34)  & args(0) & chr(34) 
if args.length = 0 then
    cmdline="powershell.exe -F C:/Users/" & ws.ExpandEnvironmentStrings("%username%") & "/AppData/Local/Microsoft/WindowsApps/right-click-diff-word.ps1"
else
    cmdline="powershell.exe -F C:/Users/" & ws.ExpandEnvironmentStrings("%username%") & "/AppData/Local/Microsoft/WindowsApps/right-click-diff-word.ps1 " & chr(34)  & args(0) & chr(34) 
end if
ws.Run cmdline,0