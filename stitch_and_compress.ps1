# SPDX-License-Identifier: Apache-2.0
#
# Tool to stitch together MOVI*.avi clips -> MP4 with audio cleanup
# Supports libx264 (CPU) and optional GPU encoders (NVENC/QSV/AMF).
# Auto-falls back to libx264 if the chosen hardware encoder isn't available.
#
# By thomas169

param(
  [int]$Start,
  [int]$End
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# --- Tools check ---
foreach ($t in 'ffmpeg','ffprobe') {
  if (-not (Get-Command $t -ErrorAction SilentlyContinue)) {
    throw "$t not found in PATH."
  }
}

# --- Helpers ---
function ReadOrDefault([string]$prompt, [string]$default) {
  $v = Read-Host "$prompt (default $default)"
  if ([string]::IsNullOrWhiteSpace($v)) { return $default } else { return $v }
}

function Escape-FilterPath {
  param([string]$Path)
  $p = $Path -replace '\\','/'   # use forward slashes
  $p = $p -replace ':','\:'      # escape drive-letter colon (C:\ -> C\:/)
  $p = $p -replace "'","\\'"     # escape single quotes if any
  return $p
}

function Test-Filter {
  param([string]$Name)
  $list = & ffmpeg -hide_banner -loglevel error -filters 2>$null
  return ($list -match "\s$([regex]::Escape($Name))\s")
}


function Build_UniquePath {
  param([string]$Path)

  Write-Host "Build_UniquePath $Path" -ForegroundColor Yellow

  $dir  = [System.IO.Path]::GetDirectoryName($Path)
  if ([string]::IsNullOrEmpty($dir)) { $dir = (Get-Location).Path }
  $base = [System.IO.Path]::GetFileNameWithoutExtension($Path)
  $ext  = [System.IO.Path]::GetExtension($Path)

  $original = [System.IO.Path]::Combine($dir, $base + $ext)
  if (-not (Test-Path -LiteralPath $original)) {
    return $original
  }

  # Always append _<n> AFTER the full base name (don’t touch numbers already in the base)
  $n = 1
  do {
    $candidate = [System.IO.Path]::Combine($dir, ("{0}_{1}{2}" -f $base, $n, $ext))
    $n++
  } while (Test-Path -LiteralPath $candidate)

  return $candidate
}




# Build the actual stabilisation filter chain for the encode pass
function Build-StabVF {
  param(
    [bool]$UseVidStab,
    [string]$TrfPath,
    [int]$Smooth,
    [string]$RSMode  # 'off','mild','med','strong' (maps to deshake params)
  )
  if ($UseVidStab) {
    $inEsc = Escape-FilterPath $TrfPath
    #return "vidstabtransform=input='$inEsc':smoothing=$Smooth:optzoom=2:interpol=bicubic:crop=black"
    return "vidstabtransform=input='$inEsc':smoothing=$($Smooth):optzoom=2:interpol=bicubic:crop=black"
  } else {
    # deshake fallback; tweak rx/ry/edge to taste
    switch ($RSMode) {
      'strong' { $rx=32; $ry=18; $edge='mirror' }
      'med'    { $rx=24; $ry=12; $edge='mirror' }
      'mild'   { $rx=16; $ry=8;  $edge='mirror' }
      default  { $rx=20; $ry=10; $edge='mirror' }
    }
    # blocksize=16 improves stability on wavy/jello footage; contrast helps lock features
    #return "deshake=rx=$rx:ry=$ry:blocksize=16:contrast=0.25:edge=$edge"
    return "deshake=rx=$($rx):ry=$($ry):blocksize=16:contrast=0.25:edge=$edge"

  }
}

function Test-Encoder {
  param([string]$Enc)

  # Probe encoder with a tiny synthetic input; suppress errors & don't throw
  $args = @(
    '-hide_banner','-loglevel','error',
    '-f','lavfi','-i','color=s=16x16:d=0.05:r=5',
    '-t','0.05','-an',
    '-c:v', $Enc,
    '-f','null','-'
  )

  $prev = $global:ErrorActionPreference
  try {
    $global:ErrorActionPreference = 'Continue'
    $null = & ffmpeg @args 2>$null
    return ($LASTEXITCODE -eq 0)
  } catch {
    return $false
  } finally {
    $global:ErrorActionPreference = $prev
  }
}

function Map-NvencPreset {
  param([string]$Preset)
  # Accept x264-style names or NVENC p1..p7; map sensibly to NVENC.
  switch -Regex ($Preset.ToLower()) {
    'p[1-7]'     { return $Preset.ToLower() }
    'ultrafast'  { return 'p1' }
    'superfast'  { return 'p2' }
    'veryfast'   { return 'p3' }
    'faster'     { return 'p4' }
    'fast'       { return 'p5' }
    'medium'     { return 'p5' }
    'slow'       { return 'p6' }
    'slower'     { return 'p7' }
    'placebo'    { return 'p7' }
    default      { return 'p6' }
  }
}

function Get-VideoArgs {
  param(
    [string]$Encoder,  # libx264 | h264_nvenc | hevc_nvenc | h264_qsv | hevc_qsv | h264_amf | hevc_amf
    [int]$CRF,
    [string]$Preset
  )
  switch ($Encoder) {
    'libx264' {
      return @('-c:v','libx264','-preset',$Preset,'-crf',"$CRF",'-pix_fmt','yuv420p')
    }
    'h264_nvenc' {
      $cq = [Math]::Min([Math]::Max($CRF, 15), 28)
      $nvPreset = Map-NvencPreset $Preset
      return @('-c:v','h264_nvenc','-preset',$nvPreset,'-rc','vbr','-cq',"$cq",'-b:v','0','-maxrate','0','-pix_fmt','yuv420p')
    }
    'hevc_nvenc' {
      $cq = [Math]::Min([Math]::Max($CRF, 15), 30)
      $nvPreset = Map-NvencPreset $Preset
      return @('-c:v','hevc_nvenc','-preset',$nvPreset,'-rc','vbr','-cq',"$cq",'-b:v','0','-maxrate','0','-pix_fmt','yuv420p')
    }
    'h264_qsv' {
      $icq = [Math]::Min([Math]::Max($CRF, 15), 28)
      return @('-c:v','h264_qsv','-global_quality',"$icq",'-look_ahead','1','-pix_fmt','nv12')
    }
    'hevc_qsv' {
      $icq = [Math]::Min([Math]::Max($CRF, 15), 30)
      return @('-c:v','hevc_qsv','-global_quality',"$icq",'-look_ahead','1','-pix_fmt','nv12')
    }
    'h264_amf' {
      $qp = [int][Math]::Round([Math]::Min([Math]::Max(($CRF - 2), 16), 30))
      return @('-c:v','h264_amf','-rc','cqp','-qp_i',"$qp",'-qp_p',"$qp",'-qp_b',"$qp",'-pix_fmt','yuv420p')
    }
    'hevc_amf' {
      $qp = [int][Math]::Round([Math]::Min([Math]::Max(($CRF - 2), 18), 32))
      return @('-c:v','hevc_amf','-rc','cqp','-qp_i',"$qp",'-qp_p',"$qp",'-qp_b',"$qp",'-pix_fmt','yuv420p')
    }
    default { throw "Unknown encoder '$Encoder'." }
  }
}

# --- Collect and select clips ---
$clips = Get-ChildItem -File -Filter "MOVI*.avi" | ForEach-Object {
  if ($_ -match '^MOVI(\d+)\.avi$') {
    [PSCustomObject]@{
      Name          = $_.Name
      File          = $_.FullName
      Number        = [int]$Matches[1]
      PadWide       = $Matches[1].Length
      LastWriteTime = $_.LastWriteTime
      SizeMB        = [Math]::Round($_.Length / 1MB, 2)
    }
  }
} | Sort-Object Number

if (-not $clips) { throw "No MOVI*.avi files found." }

$padWidth = ($clips | Select-Object -First 1).PadWide
Write-Host "`nFound $($clips.Count) clips:" -ForegroundColor Cyan
$clips | Format-Table Number, Name, LastWriteTime, SizeMB -AutoSize

$minNum = $clips[0].Number
$maxNum = $clips[-1].Number

if (-not $PSBoundParameters.ContainsKey('Start')) {
  $in = Read-Host "Enter START number (default $minNum)"
  if ($in -match '^\d+$') { $Start = [int]$in } else { $Start = $minNum }
}
if (-not $PSBoundParameters.ContainsKey('End')) {
  $in = Read-Host "Enter END number (default $maxNum)"
  if ($in -match '^\d+$') { $End = [int]$in } else { $End = $maxNum }
}
if ($Start -gt $End) { throw "Start ($Start) must be <= End ($End)." }

$sel = $clips | Where-Object { $_.Number -ge $Start -and $_.Number -le $End } | Sort-Object Number
if (-not $sel) { throw "No clips in the range [$Start..$End]." }

$startStr = $Start.ToString("D$padWidth")
$endStr   = $End.ToString("D$padWidth")
Write-Host "`nSelected range: MOVI$startStr.avi ... MOVI$endStr.avi ($($sel.Count) files)" -ForegroundColor Cyan

# Concat list
$concatPath = Join-Path $PWD "concat.txt"
$sel | ForEach-Object { "file '$($_.File)'" } | Set-Content -Encoding ASCII $concatPath
Write-Host "Concat list -> $concatPath"

# Outputs
$outMP4 = Join-Path $PWD ("stitched_{0}_{1}.mp4" -f $startStr, $endStr)
$outAVI = Join-Path $PWD ("stitched_{0}_{1}.avi" -f $startStr, $endStr)

# Ensure unique filenames (appends _1, _2, ...)
$outMP4 = Build_UniquePath $outMP4
$outAVI = Build_UniquePath $outAVI

# Detect FPS from first file (e.g., 30000/1001)
$firstFile = $sel[0].File
$fps = & ffprobe -v error -select_streams v:0 -show_entries stream=avg_frame_rate -of default=nw=1:nk=1 -- $firstFile
if (-not $fps) { $fps = "30" }

# --- User options ---
Write-Host ""
$wantAVI = ReadOrDefault "Also output a synced AVI (re-encoded, larger) [y/N].........................." "N"
$crf     = ReadOrDefault "Video quality CRF 15-28 (lower = better/larger).............................." "19"
$preset  = ReadOrDefault "x264/NVENC preset (x264: ultrafast..placebo | NVENC: p1..p7)................." "slow"
$abitrate= ReadOrDefault "Audio bitrate (e.g. 128k, 160k, 192k)........................................" "160k"
$dn      = ReadOrDefault "Denoise strength afftdn=nr 12-30 (higher = stronger)........................." "18"
$hp      = ReadOrDefault "High-pass cutoff Hz (reduce wind/rumble)....................................." "80"
$audofs  = ReadOrDefault "Constant audio offset in ms (+ve delay audio; -ve delay video; 0 = none)....." "0"
$encoder = ReadOrDefault "Encoder (libx264/h264_nvenc/hevc_nvenc/h264_qsv/hevc_qsv/h264_amf/hevc_amf).." "libx264"
$doStab  = ReadOrDefault "Enable stabilisation [y/N]..................................................." "N"
$stabStr = ReadOrDefault "Stabilisation smoothing (higher = steadier/croppier) 10-60..................." "30"
$rsMode  = ReadOrDefault "Rolling-shutter mitigation (off/mild/med/strong)............................." "med"
$doDeflk = ReadOrDefault "Deflicker to reduce shimmer [y/N]............................................" "N"


# --- place this immediately after you read $encoder ---
if (-not (Test-Encoder $encoder)) {
  Write-Warning "Encoder '$encoder' not available here (driver/build/device). Falling back to libx264."
  $encoder = 'libx264'
}
$videoArgs = Get-VideoArgs -Encoder $encoder -CRF $crf -Preset $preset

Write-Host "Using video encoder: $encoder"

# --- Filters ---
$afilters = @()
$vfilters = @()

# Constant A/V offset
if ($audofs -match '^\-?\d+$') {
  $ofsMs = [int]$audofs
  if ($ofsMs -gt 0) {
    # delay audio
    $afilters += "adelay=${ofsMs}|${ofsMs}"
  } elseif ($ofsMs -lt 0) {
    # delay video (pad leading black)
    $sec = [string]([math]::Abs($ofsMs)/1000.0)
    $vfilters += "tpad=start_duration=$sec"
  }
}

$useVidStab = $false

$trf = Join-Path $PWD "transforms.trf"
$trfEsc = Escape-FilterPath $trf


if ($doStab -match '^[Yy]') {
  if (Test-Filter 'vidstabdetect' -and Test-Filter 'vidstabtransform') {
    $useVidStab = $true
    Write-Host "Analysing shake (vidstabdetect)..." -ForegroundColor Yellow
    # Detect pass over the concatenated stream
    # --- vid.stab detect pass over the concatenated stream (force monotonic PTS) ---
    # --- vid.stab detect pass over the concatenated stream (force monotonic PTS) ---
    $null = & ffmpeg -hide_banner -nostats -y -nostdin -loglevel error `
      -fflags +genpts -avoid_negative_ts make_zero `
      -f concat -safe 0 -i $concatPath `
      -vf "fps=$fps,vidstabdetect=shakiness=5:accuracy=15:result='$trfEsc'" `
      -an -f null - 1>$null 2>$null
    if ($LASTEXITCODE -ne 0) {
      Write-Warning "vidstabdetect failed; falling back to deshake."
      $useVidStab = $false
    }
  } else {
    Write-Warning "vid.stab filters not present in this FFmpeg build; using deshake fallback."
  }
}


# --- optional stabilisation ---
if ($doStab -match '^[Yy]') {
  $vfilters += (Build-StabVF -UseVidStab:$useVidStab -TrfPath $trf -Smooth ([int]$stabStr) -RSMode $rsMode)
}


if ($doDeflk -match '^[Yy]') {
  if (Test-Filter 'deflicker') {
    $vfilters += "deflicker"
  } else {
    Write-Warning "deflicker filter not in this FFmpeg build; skipping."
  }
}

# Audio: drift fix + cleanup
$afilters += "aresample=async=1000"
$afilters += "highpass=f=$hp"
$afilters += "afftdn=nr=$dn"
$afilters += "acompressor=threshold=-20dB:ratio=3:attack=8:release=250:knee=6dB:makeup=3dB"
$afilters += "alimiter=limit=-1dB"
$af = ($afilters -join ",")

# Video: legal range + format
$vfilters += "scale=in_range=full:out_range=tv"
$vfilters += "format=yuv420p"
$vf = ($vfilters -join ",")


Write-Host "Video filter chain: $vf" -ForegroundColor DarkCyan
Write-Host "Audio filter chain: $af" -ForegroundColor DarkCyan

# --- Concat -> MP4 with progress ---
Write-Host "`nConcatenating and transcoding -> MP4 (synced) -> $outMP4" -ForegroundColor Cyan
& ffmpeg -hide_banner -y -nostdin -stats -stats_period 0.5 -v error `
  -f concat -safe 0 -i $concatPath `
  @videoArgs `
  -vf $vf `
  -fps_mode cfr -r $fps `
  -c:a aac -b:a $abitrate `
  -af $af `
  -movflags +faststart `
  -- $outMP4
if ($LASTEXITCODE -ne 0) { throw "Concat->MP4 transcode failed." }

# --- Optional synced AVI (for archival) ---
if ($wantAVI -match '^[Yy]') {
  Write-Host "Also writing synced AVI -> $outAVI" -ForegroundColor Yellow
  & ffmpeg -hide_banner -y -nostdin -stats -stats_period 0.5 -v error `
    -f concat -safe 0 -i $concatPath `
    -c:v mjpeg -qscale:v 3 `
    -fps_mode cfr -r $fps `
    -c:a pcm_s16le -af "aresample=async=1000" `
    -- $outAVI
  if ($LASTEXITCODE -ne 0) { Write-Warning "AVI encode failed." }
}

Write-Host "`nDone. Created:" -ForegroundColor Green
Write-Host "  $outMP4"
if ($wantAVI -match '^[Yy]') { Write-Host "  $outAVI" }

Write-Host @"
Tips:
- Smaller MP4: raise CRF (e.g., 21–23) or use preset "medium".
- Stronger noise cut: raise afftdn=nr (e.g., 22–26). For wind, increase highpass to 100–150 Hz.
- Auto-detected FPS: $fps   (if wrong, hard-set -r in the script).
- Constant offset? Re-run with a positive/negative ms.
- GPU encoders require proper drivers; if unavailable, this script falls back to libx264 automatically.
"@
