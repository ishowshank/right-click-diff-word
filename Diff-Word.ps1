#based on https://github.com/ForNeVeR/ExtDiff/Diff-Word.ps1
#improve some codes including:
#add finally block to clear program safely;
#add $word.Documents.Application.CompareDocuments() function to compare with options customized

param(
    [string] $BaseFileName,
    [string] $ChangedFileName
)

$ErrorActionPreference = 'Stop'

function resolve($relativePath) {
    (Resolve-Path $relativePath).Path
}

$BaseFileName = resolve $BaseFileName
$ChangedFileName = resolve $ChangedFileName

# Remove the readonly attribute because Word is unable to compare readonly
# files:
$baseFile = Get-ChildItem $BaseFileName
if ($baseFile.IsReadOnly) {
    $baseFile.IsReadOnly = $false
}

# Constants
$wdDoNotSaveChanges = 0
$wdGranularityWordLevel = 0
$wdCompareTargetNew = 2

try {
    $word = New-Object -ComObject Word.Application
    #$word.Visible = $true
    $document = $word.Documents.Open($BaseFileName, $false, $false)
    $document2 = $word.Documents.Open($ChangedFileName, $false, $false)

    #the below function can compare docments with turning on all options.
    #$document.Compare($ChangedFileName, [ref]$word.UserName, [ref]$wdCompareTargetNew, [ref]$true)

    #the below function can compare documents with options customized as you need.
    $result = $word.Documents.Application.CompareDocuments([ref]$document, [ref]$document2,[ref]$wdCompareTargetNew, [ref]$wdGranularityWordLevel, [ref]$false, [ref]$false, [ref]$false, [ref]$true, [ref]$false, [ref]$false, [ref]$true, [ref]$false, [ref]$false, [ref]$true, [ref]$word.UserName, [ref]$false)
    
    #set to 0 as you want to show the message box for saving document.
    $word.ActiveDocument.Saved = 1

    # Moving the codes to finally block, and close them safely.
    #$document.Close([ref]$wdDoNotSaveChanges)
    #$document2.Close([ref]$wdDoNotSaveChanges)

    #Turn on Word window here, due to disable the menus and function region of word. So moving the codes to finally block, then it works.
    #$word.Visible = $true
} catch {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show($_.Exception)
}finally{
    $document.Close([ref]$wdDoNotSaveChanges)
    $document2.Close([ref]$wdDoNotSaveChanges)
    $word.Visible = $true
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
