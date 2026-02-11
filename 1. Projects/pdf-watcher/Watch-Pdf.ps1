$DummyForBOM = ""
#Requires -Version 5.1
<#
.SYNOPSIS
    PDF -> Markdown 변환기 (MinerU)
.DESCRIPTION
    ┌─────────────────────────────────────────────────────┐
    │  input/ 폴더에 PDF를 넣으면 자동으로:               │
    │                                                     │
    │  1. MinerU  변환 (이미지 추출 + 문서 구조)          │
    │       ↓                                             │
    │  2. output/최종.md + output/최종_images/             │
    │  3. 원본 PDF → archive/ 이동                       │
    └─────────────────────────────────────────────────────┘

    출력 구조:
      output/{PDF이름}.md              ← 변환 결과
      output/{PDF이름}_images/         ← 추출 이미지

.USAGE
    powershell -ExecutionPolicy Bypass -File Watch-Pdf.ps1
#>

# ── 인코딩 ────────────────────────────────────────────────
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
$null = & chcp 65001 2>$null

# ── 경로 ──────────────────────────────────────────────────
$BaseDir    = Split-Path -Parent $MyInvocation.MyCommand.Path
$InputDir   = Join-Path $BaseDir "input"
$OutputDir  = Join-Path $BaseDir "output"
$ArchiveDir = Join-Path $BaseDir "archive"
$LogDir     = Join-Path $BaseDir "logs"
$LogFile    = Join-Path $LogDir  "watcher.log"

foreach ($dir in @($InputDir, $OutputDir, $ArchiveDir, $LogDir)) {
    if (-not (Test-Path $dir)) {
        New-Item -Path $dir -ItemType Directory -Force | Out-Null
    }
}

# ── 로깅 ──────────────────────────────────────────────────
function Write-Log {
    param([string]$Message)
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "[$ts] $Message"
    Write-Host $entry
    [System.IO.File]::AppendAllText(
        $LogFile, "$entry`r`n",
        [System.Text.UTF8Encoding]::new($false)
    )
}

# ══════════════════════════════════════════════════════════
#  파이프라인: PDF → MinerU 변환 → 저장
# ══════════════════════════════════════════════════════════
function Invoke-Convert {
    param([string]$FilePath)

    $fileName = Split-Path -Leaf $FilePath
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
    $ts = Get-Date -Format "yyyyMMdd_HHmmss"

    Start-Sleep -Seconds 3
    if (-not (Test-Path $FilePath)) {
        Write-Log "SKIP: $fileName"
        return
    }

    $retries = 0
    while ($retries -lt 5) {
        try {
            $s = [System.IO.File]::Open($FilePath, 'Open', 'Read', 'None')
            $s.Close(); break
        }
        catch { $retries++; Start-Sleep -Seconds 2 }
    }

    Write-Log "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
    Write-Log "NEW: $fileName"

    # ── 중복 검사: archive에 동일 파일 존재 여부 ──────────
    $archivePath = Join-Path $ArchiveDir $fileName
    if (Test-Path $archivePath) {
        Write-Log "  [SKIP] 이미 변환된 파일입니다: archive/$fileName"
        Remove-Item -Path $FilePath -Force -EA SilentlyContinue
        Write-Log "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
        return
    }

    $outputFile = Join-Path $OutputDir "${baseName}.md"

    $tmp = [System.IO.Path]::GetTempPath()
    $mnrDir = Join-Path $tmp "pdf_mineru_${ts}"
    New-Item -Path $mnrDir -ItemType Directory -Force | Out-Null

    # ── [1/2] MinerU 변환 ──────────────────────────────────
    Write-Log "  [1/2] MinerU 변환 시작..."
    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    $p = Start-Process -FilePath "mineru" `
        -ArgumentList "-p `"$FilePath`" -o `"$mnrDir`" -m auto" `
        -NoNewWindow -PassThru -Wait
    $sw.Stop()

    $mdFiles = @(Get-ChildItem $mnrDir -Recurse -Filter "*.md" -EA SilentlyContinue)
    $ok = $mdFiles.Count -gt 0
    Write-Log "  [1/2] MinerU 완료: $(if($ok){'OK'}else{'FAIL'}) ($("{0:N1}" -f $sw.Elapsed.TotalMinutes)분)"

    if (-not $ok) {
        Write-Log "  [FAIL] 변환 실패"
        Remove-Item $mnrDir -Recurse -Force -EA SilentlyContinue
        return
    }

    # ── [2/2] 결과 저장 ────────────────────────────────────
    $enc = [System.Text.UTF8Encoding]::new($false)

    # MD 복사
    [System.IO.File]::WriteAllText($outputFile,
        [System.IO.File]::ReadAllText($mdFiles[0].FullName, $enc), $enc)

    # 이미지 복사
    $imgFiles = @(Get-ChildItem $mnrDir -Recurse -Include "*.png","*.jpg","*.jpeg" -EA SilentlyContinue)
    if ($imgFiles.Count -gt 0) {
        $outImgDir = Join-Path $OutputDir "${baseName}_images"
        New-Item -Path $outImgDir -ItemType Directory -Force | Out-Null
        foreach ($f in $imgFiles) {
            Copy-Item -Path $f.FullName -Destination $outImgDir -Force
        }
        Write-Log "    이미지 $($imgFiles.Count)개 추출"
    }

    Write-Log "  [2/2] 저장 완료 → $outputFile"

    Remove-Item $mnrDir -Recurse -Force -EA SilentlyContinue

    # ── input 정리: archive로 이동 ─────────────────────────
    try {
        Move-Item -Path $FilePath -Destination (Join-Path $ArchiveDir $fileName) -Force
        Write-Log "  완료 (archive로 이동: $fileName)"
    }
    catch { Write-Log "  [ERROR] archive 이동 실패: $($_.Exception.Message)" }
    Write-Log "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
}

# ══════════════════════════════════════════════════════════
#  시작
# ══════════════════════════════════════════════════════════
Write-Log ""
Write-Log "================================================"
Write-Log " PDF -> Markdown 변환기 (MinerU)"
Write-Log "================================================"
Write-Log " Input:  $InputDir"
Write-Log " Output: $OutputDir"
Write-Log "================================================"
Write-Log " PDF를 input/ 폴더에 넣으세요. 종료: Ctrl+C"
Write-Log ""

$existing = Get-ChildItem $InputDir -Filter "*.pdf" -File -EA SilentlyContinue
if ($existing) {
    Write-Log "기존 PDF $($existing.Count)개 처리..."
    foreach ($pdf in $existing) { Invoke-Convert $pdf.FullName }
}

# ── 폴더 감시 ────────────────────────────────────────────
$script:Queue = [System.Collections.ArrayList]::Synchronized(
    [System.Collections.ArrayList]::new()
)

$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $InputDir
$watcher.Filter = "*.pdf"
$watcher.IncludeSubdirectories = $false
$watcher.NotifyFilter = [System.IO.NotifyFilters]::FileName -bor [System.IO.NotifyFilters]::LastWrite
$watcher.EnableRaisingEvents = $true

$action = {
    $fp = $Event.SourceEventArgs.FullPath
    if ([System.IO.Path]::GetExtension($fp).ToLower() -eq ".pdf") {
        $Event.MessageData.Q.Add($fp) | Out-Null
    }
}

Register-ObjectEvent $watcher "Created" -Action $action -MessageData @{Q=$script:Queue} | Out-Null
Register-ObjectEvent $watcher "Renamed" -Action $action -MessageData @{Q=$script:Queue} | Out-Null

Write-Log "input/ 감시 중..."

try {
    while ($true) {
        while ($script:Queue.Count -gt 0) {
            $fp = $script:Queue[0]
            $script:Queue.RemoveAt(0)
            Start-Sleep -Seconds 3
            if (Test-Path $fp) { Invoke-Convert $fp }
        }
        Start-Sleep -Seconds 2
    }
}
finally {
    $watcher.EnableRaisingEvents = $false
    $watcher.Dispose()
    Get-EventSubscriber | Unregister-Event
    Write-Log "종료."
}
