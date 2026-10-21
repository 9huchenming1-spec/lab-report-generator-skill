param(
    [Parameter(Mandatory=$true)][string]$Source,
    [Parameter(Mandatory=$true)][string]$Destination
)
$ErrorActionPreference = 'Stop'
$wordApp = $null
$wordDocument = $null
$originalLinkSetting = $null
try {
    $sourcePath = [System.IO.Path]::GetFullPath($Source)
    $destinationPath = [System.IO.Path]::GetFullPath($Destination)
    if ($sourcePath -eq $destinationPath -or (Test-Path -LiteralPath $destinationPath)) {
        throw 'Destination must be a new file.'
    }
    $wordApp = New-Object -ComObject Word.Application
    $wordApp.Visible = $false
    $wordApp.DisplayAlerts = 0
    $wordApp.AutomationSecurity = 3
    $originalLinkSetting = $wordApp.Options.UpdateLinksAtOpen
    $wordApp.Options.UpdateLinksAtOpen = $false
    $confirmConversions = $false
    $readOnly = $true
    $addToRecentFiles = $false
    $wordDocument = $wordApp.Documents.Open([ref]$sourcePath, [ref]$confirmConversions, [ref]$readOnly, [ref]$addToRecentFiles)
    $extension = [System.IO.Path]::GetExtension($destinationPath).ToLowerInvariant()
    if ($extension -eq '.pdf') {
        $wordDocument.ExportAsFixedFormat($destinationPath, 17)
    } elseif ($extension -eq '.docx') {
        $formatCode = 16
        $wordDocument.SaveAs2([ref]$destinationPath, [ref]$formatCode)
    } else {
        throw 'Destination format must be DOCX or PDF.'
    }
} finally {
    if ($null -ne $wordDocument) {
        try {
            $discard = 0
            $wordDocument.Close([ref]$discard)
        } catch {
            Write-Warning ('Could not close converter document: ' + $_.Exception.Message)
        } finally {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wordDocument)
        }
    }
    if ($null -ne $wordApp) {
        try {
            if ($null -ne $originalLinkSetting) {
                $wordApp.Options.UpdateLinksAtOpen = $originalLinkSetting
            }
            $discard = 0
            $wordApp.Quit([ref]$discard)
        } catch {
            Write-Warning ('Could not quit converter instance: ' + $_.Exception.Message)
        } finally {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wordApp)
        }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
