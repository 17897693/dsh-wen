# Rasterize a PDF page to PNG with Windows' built-in PDF renderer (WinRT), so
# scanned PDFs can be OCR'd without screenshots. Dense pages are cut into
# horizontal bands in JS (see png.js) because vision models cap output length.
# ASCII-only (PowerShell 5.1 reads BOM-less files as ANSI).
#   pdf-render.ps1 <pdf> <outDir> <pageIndex0based> <outName>   single page (legacy)
#   pdf-render.ps1 <pdf> <outDir> <pagesSpec1based>             batch: ONE document
#     load for the whole page list ("1,3,5-8"; -1 = last page), writes page-<n>.png.
#     Loading a big scanned PDF costs seconds per call, so batch mode is what makes
#     a 20-page OCR pass take one load instead of twenty.
#
# DSH_OFFICE_RENDER_SCALE (optional, opt-in): rasterize at 96dpi * scale instead of
#   WinRT's native default (which follows the system DPI: 120dpi on many machines).
#   Unset / unparsable / outside 0.5-4 => no PdfPageRenderOptions at all, i.e. the
#   exact previous behaviour. When set, A4 gives a deterministic 794x1123 * scale,
#   so the same value means the same pixels on every machine.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Runtime.WindowsRuntime
# WinRT types must be projected explicitly in Windows PowerShell 5.1.
[Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime] | Out-Null
[Windows.Data.Pdf.PdfPageRenderOptions, Windows.Data.Pdf, ContentType = WindowsRuntime] | Out-Null
[Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime] | Out-Null
[Windows.Storage.Streams.IRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime] | Out-Null

$src = $args[0]
$outDir = $args[1]
$batch = ($args.Count -lt 4)
New-Item -ItemType Directory -Path $outDir -Force | Out-Null

$asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
  $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1'
})[0]
$asTaskAction = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
  $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction'
})[0]

function AwaitOp($op, $type) {
  $m = $asTaskGeneric.MakeGenericMethod($type)
  $t = $m.Invoke($null, @($op))
  $t.Wait(-1) | Out-Null
  $t.Result
}
function AwaitAct($act) {
  $t = $asTaskAction.Invoke($null, @($act))
  $t.Wait(-1) | Out-Null
}

# 0.0 means "no options, native WinRT path"; a value in 0.5..4.0 means 96dpi*scale.
function Get-RenderScale {
  $raw = $env:DSH_OFFICE_RENDER_SCALE
  if (-not $raw) { return 0.0 }
  $v = 0.0
  try {
    if (-not [double]::TryParse($raw, [ref]$v)) { return 0.0 }
  } catch { return 0.0 }
  if ($v -lt 0.5 -or $v -gt 4.0) { return 0.0 }
  return $v
}

# `$page.Size` is in DIPs (A4 = 793.7 x 1122.5), but DestinationWidth/Height are
# DIPs too and the raster that comes out is dest * (systemDpi / 96) physical
# pixels (measured: dest 794 -> 993 px on this 120dpi box). So to land on exactly
# 96dpi*scale pixels the destination has to be divided by that factor. Nothing in
# a non-UI PowerShell thread reports the machine DPI, so measure it once per
# invocation: render page 1 natively and compare its PNG width with Size.Width.
$script:DpiFactor = 0.0

function Get-DpiFactor($doc) {
  if ($script:DpiFactor -gt 0) { return $script:DpiFactor }
  $f = 1.0
  try {
    $p = $doc.GetPage(0)
    $w = $p.Size.Width
    $ms = New-Object System.IO.MemoryStream
    $ras = [System.IO.WindowsRuntimeStreamExtensions]::AsOutputStream($ms)
    AwaitAct ($p.RenderToStreamAsync($ras))
    $b = $ms.ToArray()
    if ($b.Length -gt 24 -and $w -gt 0) {
      $pw = [int]$b[16] * 16777216 + [int]$b[17] * 65536 + [int]$b[18] * 256 + [int]$b[19]
      if ($pw -gt 0) { $f = $pw / $w }
    }
    $p.Dispose()
    $ms.Dispose()
  } catch { $f = 1.0 }
  if ($f -le 0) { $f = 1.0 }
  $script:DpiFactor = $f
  return $f
}

function Render-Page($doc, $pageNo1Based, $outPath) {
  $page = $doc.GetPage($pageNo1Based - 1)
  $fs = [System.IO.File]::Create($outPath)
  $ras = [System.IO.WindowsRuntimeStreamExtensions]::AsOutputStream($fs)
  $scale = Get-RenderScale
  if ($scale -gt 0) {
    $factor = Get-DpiFactor $doc
    $dw = [int][math]::Round($page.Size.Width * $scale / $factor)
    $dh = [int][math]::Round($page.Size.Height * $scale / $factor)
    if ($dw -lt 1) { $dw = 1 }
    if ($dh -lt 1) { $dh = 1 }
    $opts = New-Object Windows.Data.Pdf.PdfPageRenderOptions
    $opts.DestinationWidth = $dw
    $opts.DestinationHeight = $dh
    AwaitAct ($page.RenderToStreamAsync($ras, $opts))
  } else {
    AwaitAct ($page.RenderToStreamAsync($ras))
  }
  $fs.Close()
  $page.Dispose()
}

$sf = AwaitOp ([Windows.Storage.StorageFile]::GetFileFromPathAsync($src)) ([Windows.Storage.StorageFile])
$doc = AwaitOp ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($sf)) ([Windows.Data.Pdf.PdfDocument])
if ($env:DSH_OFFICE_RENDER_QUIET -ne '1') { "PDF pages = $($doc.PageCount)" }

if (-not $batch) {
  $pageIndex = [int]$args[2]
  $outName = $args[3]
  $outPath = Join-Path $outDir $outName
  Render-Page $doc ($pageIndex + 1) $outPath
  if ($env:DSH_OFFICE_RENDER_QUIET -ne '1') { "SAVED $outPath ($((Get-Item $outPath).Length) bytes)" }
  exit 0
}

# ---- batch mode: one load, many pages ----
$pages = New-Object System.Collections.Generic.List[int]
foreach ($part in ($args[2] -split ',')) {
  $t = $part.Trim()
  if (-not $t) { continue }
  if ($t -match '^(-?\d+)-(-?\d+)$') {
    $a = [int]$Matches[1]; $b = [int]$Matches[2]
    if ($a -lt 1) { $a = 1 }; if ($b -lt 1 -or $b -gt $doc.PageCount) { $b = [int]$doc.PageCount }
    if ($a -gt $b) { continue }
    for ($n = $a; $n -le $b; $n++) { if (-not $pages.Contains($n)) { $pages.Add($n) } }
  } elseif ($t -match '^-?\d+$') {
    $n = [int]$t
    if ($n -lt 1) { $n = [int]$doc.PageCount }
    if ($n -ge 1 -and $n -le [int]$doc.PageCount -and -not $pages.Contains($n)) { $pages.Add($n) }
  }
}
foreach ($n in $pages) {
  $outPath = Join-Path $outDir ("page-{0}.png" -f $n)
  try {
    Render-Page $doc $n $outPath
    "SAVED $outPath ($((Get-Item $outPath).Length) bytes)"
  } catch {
    "FAILED $n $($_.Exception.Message)"
  }
}
"DONE $($pages.Count)"
