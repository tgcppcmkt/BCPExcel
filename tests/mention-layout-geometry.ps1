# Runs pure geometry from the shipped C# source. No Windows APIs are called.
$ErrorActionPreference='Stop'
Add-Type -Path (Join-Path $PSScriptRoot '../メンション依頼フォーム_HTA_最新版/mention-layout-host.cs')
$cases=0
function Assert-Inside($result,$area) {
    if ($result.Left -lt $area.Left+8 -or $result.Top -lt $area.Top+8 -or
        $result.Right -gt $area.Right-8 -or $result.Bottom -gt $area.Bottom-8) {
        throw "Window clipped: $($result | ConvertTo-Json -Compress)"
    }
}
# Taskbar already subtracted; test primary, right, left, above, below and offsets.
foreach ($area in @(
    [MentionLayoutHost+Rect]::new(0,0,1920,1040),
    [MentionLayoutHost+Rect]::new(1920,0,2560,1400),
    [MentionLayoutHost+Rect]::new(-1920,0,1920,1040),
    [MentionLayoutHost+Rect]::new(0,-1080,1920,1040),
    [MentionLayoutHost+Rect]::new(0,1080,1600,860),
    [MentionLayoutHost+Rect]::new(-2560,-300,2560,1400)
)) {
    foreach ($scale in @(1.0,1.25,1.5,2.0)) {
        foreach ($height in @(520,760,1150)) {
            $old=[MentionLayoutHost+Rect]::new($area.Left+50,$area.Bottom-450,[int](1560*$scale),[int](370*$scale))
            $client=[MentionLayoutHost+Rect]::new(0,0,[int](1544*$scale),[int](330*$scale))
            $request=[int[]]@(1,1548,$height,1544,330,1560,370,0,1)
            $next=[MentionLayoutHost]::Desired($old,$client,$area,$request)
            Assert-Inside $next $area
            # Integer rounding of an existing window can shift the inferred scale by ~1px.
            if ($next.Height -lt [Math]::Min([Math]::Ceiling($height*$scale)+$old.Height-$client.Height,$area.Height-16)-2) {
                throw 'Incorrect DPI height calculation'
            }
            $cases++
        }
    }
}
# Exact screenshot-like regression: 3 cards opened in a low starting window.
$area=[MentionLayoutHost+Rect]::new(1920,0,1600,800)
$old=[MentionLayoutHost+Rect]::new(1960,66,1528,360)
$next=[MentionLayoutHost]::Fit($old,$area,1528,1000,$false)
Assert-Inside $next $area
if ($next.Left -lt 1920 -or $next.Top -ge 66 -or $next.Height -ne 784) { throw 'Screenshot regression' }
$cases++
# Shrinking should preserve an already valid placement, and centering stays local.
$old=[MentionLayoutHost+Rect]::new(-1800,100,1600,800)
$area=[MentionLayoutHost+Rect]::new(-1920,0,1920,1040)
$next=[MentionLayoutHost]::Fit($old,$area,1500,370,$false)
if ($next.Left -ne $old.Left -or $next.Top -ne $old.Top) { throw 'Shrink moved a valid window' }
$cases++
$next=[MentionLayoutHost]::Fit($old,$area,1500,370,$true)
Assert-Inside $next $area
if ($next.Left -ge 0) { throw 'Initial center changed monitor' }
$cases++
# Reject partial/invalid IPC records; invoke the shipped parser, not a copy.
$parser=[MentionLayoutHost].GetMethod('ReadRequest',[Reflection.BindingFlags]'Static,NonPublic')
$file=[IO.Path]::GetTempFileName()
try {
    foreach ($record in @('1|1548|760|1500|330|1560|370|0|1','1|1548|760','1|1548|760|1500|330|1560|370|0|2','1|1548|760|0|330|1560|370|0|1')) {
        [IO.File]::WriteAllText($file,$record)
        $arguments=[object[]]@($file,$null)
        $actual=$parser.Invoke($null,$arguments)
        $expected=$record -eq '1|1548|760|1500|330|1560|370|0|1'
        if ($actual -ne $expected) { throw 'IPC record validation failed' }
        $cases++
    }
} finally { [IO.File]::Delete($file) }
# Parse the distributed PowerShell entry point without executing it.
$tokens=$null; $errors=$null
$null=[Management.Automation.Language.Parser]::ParseFile(
    (Join-Path $PSScriptRoot '../メンション依頼フォーム_HTA_最新版/mention-layout-host.ps1'),[ref]$tokens,[ref]$errors)
if ($errors.Count) { throw ($errors | Out-String) }
Write-Output "$cases geometry/protocol cases passed; C# compiled; PowerShell parsed. Native Windows calls remain untested."
