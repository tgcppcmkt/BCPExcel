param(
    [Parameter(Mandatory=$true)][string]$StateFolder,
    [Parameter(Mandatory=$true)][string]$WindowTitle
)
# Called directly from the form. No elevation or execution-policy changes.
$ErrorActionPreference = 'Stop'
try {
    $full = [IO.Path]::GetFullPath($StateFolder)
    $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if (-not $full.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -or
        [IO.Path]::GetFileName($full) -notmatch '^MentionLayout-[a-z0-9-]+$') {
        throw 'Invalid layout state folder'
    }
    $parentId = (Get-CimInstance Win32_Process -Filter "ProcessId=$PID").ParentProcessId
    Add-Type -Path (Join-Path $PSScriptRoot 'mention-layout-host.cs')
    [MentionLayoutHost]::Run($full, [int]$parentId, $WindowTitle)
} catch {
    if ($full -and [IO.Directory]::Exists($full) -and $full.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase)) {
        [IO.File]::WriteAllText((Join-Path $full 'error.txt'), $_.Exception.Message)
    }
    exit 1
}
