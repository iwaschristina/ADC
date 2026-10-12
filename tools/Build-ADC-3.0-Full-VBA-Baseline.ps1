#Requires -Version 5.1
<#
  ADC 3.0: build the authoritative FULL VBA baseline from xlib2.accdb.
  The ZIP contains ONLY complete, unmodified VBA exports named x__UTIL_ADC
  or starting x__UTIL_ADC_. No manifest, database, unrelated module or script
  is placed inside the ZIP. This script runs on Windows with Microsoft Access.

  Requires: Microsoft Access desktop; Trust Center > Macro Settings >
  "Trust access to the VBA project object model". Run against the unchanged
  xlib2.accdb extracted from the ADC 3.0 native baseline package.

  The database opened in Access is always a disposable COPY; the original
  xlib2.accdb is never opened for editing or otherwise modified by the script.
#>
[CmdletBinding()]
param(
    [string]$DatabasePath = (Join-Path $PSScriptRoot 'xlib2.accdb'),
    [string]$OutputDirectory = $PSScriptRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$expectedSha256 = 'dfb116bb18e6e0556067584ca13eae9bd259f07382a1a6c877dc61a159f8b659'
$release = 'ADC 3.0'
$allowedNamePattern = '^x__UTIL_ADC(?:_|$)'
$protectedNames = @('x__UTIL_ADC_MODULE_GUARD', 'x__UTIL_ADC_NO_EDIT_REFERENCE')

function Get-StreamSha256([System.IO.Stream]$Stream) {
    $algorithm = [System.Security.Cryptography.SHA256]::Create()
    try {
        $sum = $algorithm.ComputeHash($Stream)
        return ([BitConverter]::ToString($sum)).Replace('-', '').ToLowerInvariant()
    }
    finally { $algorithm.Dispose() }
}

if (-not (Test-Path -LiteralPath $DatabasePath -PathType Leaf)) {
    throw "Source database not found: $DatabasePath. Extract xlib2.accdb from the authoritative ADC 3.0 package first."
}
$sourcePath = (Resolve-Path -LiteralPath $DatabasePath).Path
if (-not [string]::Equals([IO.Path]::GetExtension($sourcePath), '.accdb', [StringComparison]::OrdinalIgnoreCase)) {
    throw 'The authoritative source must be an .accdb file.'
}
$sourceSha256 = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash.ToLowerInvariant()
if ($sourceSha256 -ne $expectedSha256) {
    throw "Database SHA-256 mismatch. This is not the exact xlib2.accdb provided for ADC 3.0. Actual: $sourceSha256; expected: $expectedSha256. No baseline was created."
}
if (-not (Test-Path -LiteralPath $OutputDirectory -PathType Container)) {
    throw "Output directory does not exist: $OutputDirectory"
}
$outputRoot = (Resolve-Path -LiteralPath $OutputDirectory).Path
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$base = "VBA_Export_x_UTIL_ADC_${stamp}_ADC_3.0_CGPT_FINAL_FULL_BASELINE"
$stagePath = Join-Path $outputRoot ($base + '_source')
$zipPath = Join-Path $outputRoot ($base + '.zip')
$manifestPath = Join-Path $outputRoot ($base + '_MANIFEST.json')
foreach ($target in @($stagePath, $zipPath, $manifestPath)) {
    if (Test-Path -LiteralPath $target) { throw "Target already exists; refusing overwrite: $target" }
}
$null = New-Item -ItemType Directory -Path $stagePath
$tempDir = Join-Path ([IO.Path]::GetTempPath()) ('ADC30_VBA_' + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $tempDir
$databaseCopy = Join-Path $tempDir ([IO.Path]::GetFileName($sourcePath))
Copy-Item -LiteralPath $sourcePath -Destination $databaseCopy
$app = $null
$records = New-Object 'System.Collections.Generic.List[object]'
$allEligibleNames = New-Object 'System.Collections.Generic.List[string]'
$allOutputFiles = New-Object 'System.Collections.Generic.List[string]'
$mainError = $null
try {
    $app = New-Object -ComObject Access.Application
    # Automation reads only a disposable copy; Access may create lock/cache files.
    $app.OpenCurrentDatabase($databaseCopy, $false)
    try { $project = $app.VBE.ActiveVBProject }
    catch {
        throw ('Cannot inspect Access VBA. Enable "Trust access to the VBA project object model" in Access Trust Center. ' + $_.Exception.Message)
    }
    if ($null -eq $project) { throw 'No VBA project is accessible in the database copy.' }

    foreach ($component in $project.VBComponents) {
        $moduleName = [string]$component.Name
        if ($moduleName -notmatch $allowedNamePattern) { continue }
        $allEligibleNames.Add($moduleName)
    }
    if ($allEligibleNames.Count -eq 0) { throw 'No x__UTIL_ADC VBA modules found. No baseline was created.' }
    foreach ($protectedName in $protectedNames) {
        if (-not ($allEligibleNames -contains $protectedName)) {
            throw "Required protected module not found: $protectedName. No baseline was created."
        }
    }
    # Reject duplicate/case-colliding module names on case-insensitive Windows.
    $uniqueNames = @($allEligibleNames | Sort-Object -Unique)
    if ($uniqueNames.Count -ne $allEligibleNames.Count) { throw 'Duplicate ADC module names detected.' }

    foreach ($component in $project.VBComponents) {
        $moduleName = [string]$component.Name
        if ($moduleName -notmatch $allowedNamePattern) { continue }
        if ($moduleName -match '[<>:"/\\|?*]') { throw "Invalid component filename: $moduleName" }

        $componentType = [int]$component.Type
        $extension = switch ($componentType) {
            1   { '.bas'; break } # Standard module
            2   { '.cls'; break } # Class
            3   { '.frm'; break } # UserForm
            100 { '.cls'; break } # Document module
            default { throw "Unknown VBA component type $componentType ($moduleName); refusing incomplete baseline." }
        }
        $lineCount = [int]$component.CodeModule.CountOfLines
        $header = ''
        if ($lineCount -gt 0) { $header = [string]$component.CodeModule.Lines(1, [Math]::Min(160, $lineCount)) }
        $versionMatch = [regex]::Match($header, "(?im)^\s*'\s*ChatGPT Version ID:\s*(CGPT-[A-Z0-9-]+)\s*$")
        $previousMatch = [regex]::Match($header, "(?im)^\s*'\s*ChatGPT Previous Version ID:\s*(.*?)\s*$")
        $baselineMatch = [regex]::Match($header, "(?im)^\s*'\s*ChatGPT Baseline Version ID:\s*(.*?)\s*$")
        $timestampMatch = [regex]::Match($header, "(?im)^\s*'\s*ChatGPT Version Timestamp:\s*(.*?)\s*$")
        $version = if ($versionMatch.Success) { $versionMatch.Groups[1].Value } else { $null }
        # NO_EDIT_REFERENCE has historically been delivered under its bare module name.
        $basename = $moduleName
        if ($moduleName -ine 'x__UTIL_ADC_NO_EDIT_REFERENCE' -and $version) {
            $basename += '_' + $version
        }
        $filename = $basename + $extension
        $dest = Join-Path $stagePath $filename
        if (Test-Path -LiteralPath $dest) { throw "Filename collision: $filename" }

        # Native VBE export: no edits, repairs, rewrites or format conversion.
        $component.Export($dest)
        if (-not (Test-Path -LiteralPath $dest -PathType Leaf)) { throw "No file produced for $moduleName" }
        $sourceText = Get-Content -LiteralPath $dest -Raw -Encoding Default
        $nameCheck = [regex]::Match($sourceText, '(?im)^\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"')
        if (-not $nameCheck.Success) {
            throw "Exported component missing Attribute VB_Name: $moduleName"
        }
        if (-not [string]::Equals($nameCheck.Groups[1].Value, $moduleName, [StringComparison]::Ordinal)) {
            throw "Exported VB_Name mismatch: $moduleName / $($nameCheck.Groups[1].Value)"
        }
        if ($version -and -not $sourceText.Contains('ChatGPT Version ID: ' + $version)) {
            throw "Version header mismatch between VBE and export: $moduleName"
        }
        if ($moduleName -notmatch $allowedNamePattern) { throw "Unexpected module: $moduleName" }
        $fileInfo = Get-Item -LiteralPath $dest
        if ($fileInfo.Length -le 0) { throw "Zero-byte exported file: $moduleName" }
        $fileSha = (Get-FileHash -LiteralPath $dest -Algorithm SHA256).Hash.ToLowerInvariant()
        $allOutputFiles.Add($dest)
        # VBE may emit a .frx sidecar for a UserForm. Preserve it exactly if present.
        if ($extension -eq '.frm') {
            $frx = Join-Path $stagePath ($basename + '.frx')
            if (Test-Path -LiteralPath $frx -PathType Leaf) { $allOutputFiles.Add($frx) }
        }
        $records.Add([pscustomobject]@{
            module_name=$moduleName
            component_type=$componentType
            export_filename=$filename
            export_bytes=$fileInfo.Length
            export_sha256=$fileSha
            code_lines=$lineCount
            chatgpt_version_id=$version
            chatgpt_previous_version_id=$(if ($previousMatch.Success) {$previousMatch.Groups[1].Value} else {$null})
            chatgpt_baseline_version_id=$(if ($baselineMatch.Success) {$baselineMatch.Groups[1].Value} else {$null})
            chatgpt_version_timestamp=$(if ($timestampMatch.Success) {$timestampMatch.Groups[1].Value} else {$null})
            protected=($moduleName -in $protectedNames)
        })
    }
    if ($records.Count -ne $allEligibleNames.Count) {
        throw "Export count mismatch: $($records.Count) exported; $($allEligibleNames.Count) eligible."
    }
}
catch { $mainError = $_ }
finally {
    if ($null -ne $app) {
        try { $app.CloseCurrentDatabase() } catch { }
        try { $app.Quit() } catch { }
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) } catch { }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    try { Remove-Item -LiteralPath $tempDir -Recurse -Force } catch { Write-Warning "Disposable Access copy remains at $tempDir" }
}
$afterSha = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash.ToLowerInvariant()
if ($afterSha -ne $sourceSha256) { throw 'CRITICAL: original source database hash changed. Refusing baseline ZIP.' }
if ($null -ne $mainError) { throw $mainError }

# ZIP contents: only the exact exported VBA components, at archive root.
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::Open($zipPath, [System.IO.Compression.ZipArchiveMode]::Create)
try {
    foreach ($file in $allOutputFiles) {
        $entryName = [IO.Path]::GetFileName($file)
        if ($entryName -notmatch '^x__UTIL_ADC(?:_|\.|$)') { throw "Unexpected archive file name: $entryName" }
        [void][System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $file, $entryName, [IO.Compression.CompressionLevel]::Optimal)
    }
}
finally { $zip.Dispose() }

# Independently read the completed ZIP and hash each member against staged bytes.
$verifyZip = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
try {
    if ($verifyZip.Entries.Count -ne $allOutputFiles.Count) { throw 'ZIP member count does not match source files.' }
    foreach ($entry in $verifyZip.Entries) {
        if ($entry.FullName -ne $entry.Name) { throw "Non-flat ZIP path: $($entry.FullName)" }
        if ($entry.Name -notmatch '^x__UTIL_ADC(?:_|\.|$)') { throw "Non-ADC ZIP entry: $($entry.Name)" }
        $originalFile = Join-Path $stagePath $entry.Name
        if (-not (Test-Path -LiteralPath $originalFile)) { throw "Unknown ZIP entry: $($entry.Name)" }
        $stream = $entry.Open()
        try { $entrySha = Get-StreamSha256 $stream } finally { $stream.Dispose() }
        $originalSha = (Get-FileHash -LiteralPath $originalFile -Algorithm SHA256).Hash.ToLowerInvariant()
        if ($entrySha -ne $originalSha) { throw "ZIP/source mismatch: $($entry.Name)" }
    }
}
finally { $verifyZip.Dispose() }

# Sidecar manifest: not inside the ZIP (archive format matches historic flat ZIPs).
$manifest = [ordered]@{
    utility_release=$release
    build_status='complete-vba-export-local-verified'
    created_at=(Get-Date).ToString('o')
    source_database=[IO.Path]::GetFileName($sourcePath)
    source_database_sha256=$sourceSha256
    source_database_original_unchanged=$true
    source_filter='VB_Name = x__UTIL_ADC OR VB_Name starts with x__UTIL_ADC_'
    all_eligible_module_count=$allEligibleNames.Count
    exported_module_count=$records.Count
    archive_file_count=$allOutputFiles.Count
    flat_module_only_archive=$true
    zip_filename=[IO.Path]::GetFileName($zipPath)
    zip_sha256=(Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
    modules=@($records.ToArray())
}
$manifest | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
Write-Host ('ADC full VBA baseline created: ' + $zipPath)
Write-Host ("Module count: $($records.Count); ZIP member count: $($allOutputFiles.Count)")
Write-Host ('Manifest (separate from ZIP): ' + $manifestPath)
Write-Host ('Original Access database hash unchanged: ' + $afterSha)
Write-Host 'All exported source files and ZIP entries were checked. No code was modified.'