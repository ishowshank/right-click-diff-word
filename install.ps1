Param(
    [string]$regfile=".\default-right-menu-setting.reg"
)

$regfile = Resolve-Path $regfile
$content = Get-Content $regfile
$result = $content.Replace("username", $env:USERNAME)
$result | Out-File ".\your-right-menu-setting.reg"

cd .\
Copy-Item *.vbs C:\Users\$env:USERNAME\AppData\Local\Microsoft\WindowsApps\
Copy-Item *.ps1 C:\Users\$env:USERNAME\AppData\Local\Microsoft\WindowsApps\
Copy-Item *.cmd C:\Users\$env:USERNAME\AppData\Local\Microsoft\WindowsApps\

#it is unable to reg, so it need to reg it by double-clicking reg file yourself
#reg import .\your-right-menu-setting.reg
