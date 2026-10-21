# Development-only fixture maker; the public workflow converts DOC to DOCX.
param(
    [Parameter(Mandatory=$true)][string]$Source,
    [Parameter(Mandatory=$true)][string]$Destination
)
$ErrorActionPreference = 'Stop'
$wordApp = $null
$document = $null
try {
    $sourcePath = [System.IO.Path]::GetFullPath($Source)
    $destinationPath = [System.IO.Path]::GetFullPath($Destination)
    if ((Test-Path -LiteralPath $destinationPath) -or $sourcePath -eq $destinationPath) {
        throw 'Fixture destination already exists or equals source.'
    }
    if ([System.IO.Path]::GetExtension($destinationPath) -ne '.doc') {
        throw 'Fixture destination must be .doc.'
    }
    $wordApp = New-Object -ComObject Word.Application
    $wordApp.Visible = $false
    $wordApp.DisplayAlerts = 0
    $wordApp.AutomationSecurity = 3
    $confirm = $false
    $readOnly = $true
    $recent = $false
    $document = $wordApp.Documents.Open([ref]$sourcePath, [ref]$confirm, [ref]$readOnly, [ref]$recent)
    $formatCode = 0
    $document.SaveAs2([ref]$destinationPath, [ref]$formatCode)
} finally {
    $discard = 0
    if ($null -ne $document) {
        try { $document.Close([ref]$discard) } finally {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
        }
    }
    if ($null -ne $wordApp) {
        try { $wordApp.Quit([ref]$discard) } finally {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($wordApp)
        }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
