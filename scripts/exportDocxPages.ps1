param(
  [Parameter(Mandatory=$true)][string]$docxPath,
  [Parameter(Mandatory=$true)][string]$outputDir,
  [int]$width = 1600,
  [int]$height = 2260
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

if (-not (Test-Path $outputDir)) {
  New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

$resolvedDocx = (Resolve-Path $docxPath).Path
$resolvedOut = (Resolve-Path $outputDir).Path

function Invoke-ExportDocxPages {
  param([string]$filePath, [string]$outDir, [int]$w, [int]$h)

  $word = $null
  $doc = $null

  try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    try { $word.DisplayAlerts = 0 } catch {}

    # Open(FileName, ConfirmConversions=0, ReadOnly=1, AddToRecentFiles=0)
    $doc = $word.Documents.Open($filePath, $false, $true, $false)
    $panes = $doc.ActiveWindow.Panes
    if ($panes.Count -lt 1) {
      return 0
    }

    $pages = $panes.Item(1).Pages
    $count = $pages.Count
    if ($count -lt 1) { return 0 }

    for ($i = 1; $i -le $count; $i++) {
      $page = $pages.Item($i)
      $bits = $page.EnhMetaFileBits

      if ($bits -ne $null -and $bits.Length -gt 0) {
        $ms = New-Object System.IO.MemoryStream (,$bits)
        $meta = New-Object System.Drawing.Imaging.Metafile $ms
        $bmp = New-Object System.Drawing.Bitmap $w, $h
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
        $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
        $g.Clear([System.Drawing.Color]::White)
        $g.DrawImage($meta, 0, 0, $w, $h)

        $outPath = Join-Path $outDir ("slide_{0}.png" -f $i)
        $bmp.Save($outPath, [System.Drawing.Imaging.ImageFormat]::Png)

        $g.Dispose()
        $bmp.Dispose()
        $meta.Dispose()
        $ms.Dispose()
      }

      try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($page) | Out-Null } catch {}
    }

    return $count
  } finally {
    if ($doc -ne $null) {
      try { $doc.Close($false) } catch {}
      try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null } catch {}
    }
    if ($word -ne $null) {
      try { $word.Quit($false) } catch {}
      try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch {}
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
    $pageCount = Invoke-ExportDocxPages -filePath $resolvedDocx -outDir $resolvedOut -w $width -h $height
    Write-Output "SUCCESS:$pageCount"
    $succeeded = $true
    break
  } catch {
    $lastError = $_
    Stop-Process -Name WINWORD -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 800
  }
}

if (-not $succeeded) {
  Write-Error "Failed to export Word document pages after $maxAttempts attempts: $($lastError.Exception.Message)"
  exit 1
}
