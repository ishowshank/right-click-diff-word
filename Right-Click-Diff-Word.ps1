param(
    [string] $TargetFileName=""
)

$_ext = ".doc"
$_tmp = "$env:TMP\_tmp.txt"
$BaseFileName = ""
$ChangedFileName = ""

if($TargetFileName -eq ""){
    echo "" > $_tmp
}else{
    if($TargetFileName -ne ""){
        if($TargetFileName -imatch $_ext){
            echo $TargetFileName >> $_tmp
        }
    }
}
Get-Content $_tmp | % {
    if($_.ToString().Trim() -ne ""){
        if($BaseFileName -eq "") {
            $BaseFileName = $_
        }elseif($ChangedFileName -eq ""){
            $ChangedFileName = $_
        }
    }
}
if($BaseFileName -ne "" -and $ChangedFileName -ne ""){
    #Write-Host "base=$BaseFileName, changed=$ChangedFileName"        
    echo "" > $_tmp
    powershell.exe -F C:\Users\$env:USERNAME\AppData\Local\Microsoft\WindowsApps\Diff-Word.ps1 "$BaseFileName" "$ChangedFileName"
}
