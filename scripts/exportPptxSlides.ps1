param(
  [Parameter(Mandatory=$true)][string]$pptxPath,
  [Parameter(Mandatory=$true)][string]$outputDir,
  [int]$width = 1920,
  [int]$height = 1080
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path $outputDir)) {
  New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

$resolvedPptx = (Resolve-Path $pptxPath).Path
$resolvedOut = (Resolve-Path $outputDir).Path

function Invoke-ExportSlides {
  param([string]$filePath, [string]$outDir, [int]$w, [int]$h)

  $ppt = $null
  $pres = $null
  try {
    $ppt = New-Object -ComObject PowerPoint.Application
    # Suppress all dialogs and prompts
    try { $ppt.DisplayAlerts = 1 } catch {}
    
    # Open(FileName, ReadOnly=1, Untitled=0, WithWindow=0)
    $pres = $ppt.Presentations.Open($filePath, 1, 0, 0)
    $count = $pres.Slides.Count
    
    for ($i = 1; $i -le $count; $i++) {
      $slide = $pres.Slides.Item($i)
      $outPath = Join-Path $outDir ("slide_{0}.png" -f $i)
      $slide.Export($outPath, "PNG", $w, $h)
      try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($slide) | Out-Null } catch {}
    }
    
    return $count
  } finally {
    if ($pres -ne $null) {
      try { $pres.Close() } catch {}
      try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($pres) | Out-Null } catch {}
    }
    if ($ppt -ne $null) {
      try { $ppt.Quit() } catch {}
      try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null } catch {}
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
  }
}

$maxAttempts = 3
$succeeded = $false
$lastError = $null

for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
  try {
    $slideCount = Invoke-ExportSlides -filePath $resolvedPptx -outDir $resolvedOut -w $width -h $height
    Write-Output "SUCCESS:$slideCount"
    $succeeded = $true
    break
  } catch {
    $lastError = $_
    Stop-Process -Name POWERPNT -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 800
  }
}

if (-not $succeeded) {
  Write-Error "Failed to export slides after $maxAttempts attempts: $($lastError.Exception.Message)"
  exit 1
}
