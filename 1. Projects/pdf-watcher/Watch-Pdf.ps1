$DummyForBOM = ""
#Requires -Version 5.1
<#
.SYNOPSIS
    PDF -> Markdown 통합 변환기 (Marker + MinerU 병합)
.DESCRIPTION
    ┌─────────────────────────────────────────────────────┐
    │  input/ 폴더에 PDF를 넣으면 자동으로:               │
    │                                                     │
    │  1. MinerU  변환 (이미지 추출)                      │
    │       ↓ 완료 후                                     │
    │  2. Marker  변환 (문서 구조)                        │
    │       ↓                                             │
    │  3. 병합 ─→ output/최종.md + output/최종_images/    │
    │       ↓                                             │
    │  4. 중간 결과 ─→ output/_intermediate/              │
    │  5. 원본 PDF 삭제 (input/ 비움)                     │
    └─────────────────────────────────────────────────────┘

    병합 전략:
      Marker  → 문서 구조 (계층적 제목, MD 테이블, 참조 링크)
      MinerU  → 이미지 추출 (실제 이미지 파일)
      정리    → HTML sub/sup 태그를 LaTeX로 통일

    출력 구조:
      output/{PDF이름}.md              ← 병합 결과
      output/{PDF이름}_images/         ← 추출 이미지
      output/_intermediate/{PDF이름}_marker.md  ← Marker 원본
      output/_intermediate/{PDF이름}_mineru.md  ← MinerU 원본

.USAGE
    powershell -ExecutionPolicy Bypass -File Watch-Pdf.ps1
#>

# ── 인코딩 ────────────────────────────────────────────────
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
$null = & chcp 65001 2>$null

# ── 경로 ──────────────────────────────────────────────────
$BaseDir         = Split-Path -Parent $MyInvocation.MyCommand.Path
$InputDir        = Join-Path $BaseDir "input"
$OutputDir       = Join-Path $BaseDir "output"
$IntermediateDir = Join-Path $OutputDir "_intermediate"
$LogDir          = Join-Path $BaseDir "logs"
$LogFile         = Join-Path $LogDir  "watcher.log"

foreach ($dir in @($InputDir, $OutputDir, $IntermediateDir, $LogDir)) {
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
#  병합: Marker 구조 + MinerU 이미지 + LaTeX 정리
# ══════════════════════════════════════════════════════════
function Merge-Outputs {
    param(
        [string]$MarkerMdPath,
        [string]$MineruMdPath,
        [string]$MineruTempDir,
        [string]$OutputMdPath,
        [string]$BaseName
    )

    $enc = [System.Text.UTF8Encoding]::new($false)
    $markerText = [System.IO.File]::ReadAllText($MarkerMdPath, $enc)
    $mineruText = [System.IO.File]::ReadAllText($MineruMdPath, $enc)

    $merged = $markerText

    # ── 이미지: MinerU 추출 파일로 교체 ──────────────────
    $imgRx = '!\[[^\]]*\]\(([^)]+)\)'
    $mrkImgs = [regex]::Matches($markerText, $imgRx)
    $mnrImgs = [regex]::Matches($mineruText, $imgRx)

    $imgFiles = @(
        Get-ChildItem -Path $MineruTempDir -Recurse `
            -Include "*.png","*.jpg","*.jpeg" `
            -ErrorAction SilentlyContinue
    )

    $outImgDir = Join-Path (Split-Path $OutputMdPath) "${BaseName}_images"
    if ($imgFiles.Count -gt 0) {
        New-Item -Path $outImgDir -ItemType Directory -Force | Out-Null
        foreach ($f in $imgFiles) {
            Copy-Item -Path $f.FullName -Destination $outImgDir -Force
        }
        Write-Log "    이미지 $($imgFiles.Count)개 추출"
    }

    $n = [Math]::Min($mrkImgs.Count, $mnrImgs.Count)
    for ($i = 0; $i -lt $n; $i++) {
        $old = $mrkImgs[$i].Groups[1].Value
        $new = "${BaseName}_images/$([System.IO.Path]::GetFileName($mnrImgs[$i].Groups[1].Value))"
        $merged = $merged.Replace("($old)", "($new)")
    }

    # ── HTML → LaTeX 통일 ────────────────────────────────
    $merged = $merged -replace '([A-Za-z]+)<sub>([^<]+)</sub>([A-Za-z\d]*)', '$$$1_{$2}$3$$'
    $merged = $merged -replace '([A-Za-z]+)<sup>([^<]+)</sup>([A-Za-z\d]*)', '$$$1^{$2}$3$$'
    $merged = $merged -replace '<sub>([^<]+)</sub>', '_{$1}'
    $merged = $merged -replace '<sup>([^<]+)</sup>', '^{$1}'

    [System.IO.File]::WriteAllText($OutputMdPath, $merged, $enc)
}

# ══════════════════════════════════════════════════════════
#  파이프라인: PDF → 병렬 변환 → 병합 → 저장
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

    $outputFile  = Join-Path $OutputDir "${baseName}.md"

    $tmp = [System.IO.Path]::GetTempPath()
    $mrkDir = Join-Path $tmp "pdf_marker_${ts}"
    $mnrDir = Join-Path $tmp "pdf_mineru_${ts}"
    New-Item -Path $mrkDir -ItemType Directory -Force | Out-Null
    New-Item -Path $mnrDir -ItemType Directory -Force | Out-Null

    # ── [1/3] 순차 변환: MinerU → Marker (GPU 경합 방지) ──
    Write-Log "  [1/3] MinerU → Marker 순차 변환..."
    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    $mrkLog = Join-Path $mrkDir "_marker.log"
    $mnrLog = Join-Path $mnrDir "_mineru.log"

    # ── MinerU 먼저 실행 ──
    Write-Log "    MinerU 변환 시작..."
    $p2 = Start-Process -FilePath "mineru" `
        -ArgumentList "-p `"$FilePath`" -o `"$mnrDir`" -m auto" `
        -NoNewWindow -PassThru `
        -RedirectStandardOutput $mnrLog `
        -RedirectStandardError (Join-Path $mnrDir "_mineru_err.log")
    $p2.WaitForExit()

    $md2 = @(Get-ChildItem $mnrDir -Recurse -Filter "*.md" -EA SilentlyContinue)
    $ok2 = $md2.Count -gt 0
    Write-Log "    MinerU 완료: $(if($ok2){'OK'}else{'FAIL'})"

    # ── Marker 실행 ──
    Write-Log "    Marker 변환 시작..."
    $p1 = Start-Process -FilePath "marker_single" `
        -ArgumentList "`"$FilePath`" --output_dir `"$mrkDir`"" `
        -NoNewWindow -PassThru `
        -RedirectStandardOutput $mrkLog `
        -RedirectStandardError (Join-Path $mrkDir "_marker_err.log")
    $p1.WaitForExit()
    $sw.Stop()

    $md1 = @(Get-ChildItem $mrkDir -Recurse -Filter "*.md" -EA SilentlyContinue)
    $ok1 = $md1.Count -gt 0
    Write-Log "    Marker 완료: $(if($ok1){'OK'}else{'FAIL'})"

    Write-Log "  [1/3] 완료 ($("{0:N1}" -f $sw.Elapsed.TotalMinutes)분) MinerU:$(if($ok2){'OK'}else{'FAIL'}) Marker:$(if($ok1){'OK'}else{'FAIL'})"

    # ── 중간 결과 저장 ──
    $enc = [System.Text.UTF8Encoding]::new($false)
    if ($ok2) {
        $intMineruPath = Join-Path $IntermediateDir "${baseName}_mineru.md"
        [System.IO.File]::WriteAllText($intMineruPath,
            [System.IO.File]::ReadAllText($md2[0].FullName, $enc), $enc)
        Write-Log "    중간 저장: ${baseName}_mineru.md"
    }
    if ($ok1) {
        $intMarkerPath = Join-Path $IntermediateDir "${baseName}_marker.md"
        [System.IO.File]::WriteAllText($intMarkerPath,
            [System.IO.File]::ReadAllText($md1[0].FullName, $enc), $enc)
        Write-Log "    중간 저장: ${baseName}_marker.md"
    }

    # ── [2/3] 병합 ───────────────────────────────────────
    if ($ok1 -and $ok2) {
        Write-Log "  [2/3] 병합: Marker 구조 + MinerU 이미지"
        try {
            Merge-Outputs `
                -MarkerMdPath $md1[0].FullName `
                -MineruMdPath $md2[0].FullName `
                -MineruTempDir $mnrDir `
                -OutputMdPath $outputFile `
                -BaseName $baseName
            Write-Log "  [2/3] 병합 완료"
        }
        catch {
            Write-Log "  [2/3] 병합 실패 → Marker 대체"
            [System.IO.File]::WriteAllText($outputFile,
                [System.IO.File]::ReadAllText($md1[0].FullName, [System.Text.UTF8Encoding]::new($false)),
                [System.Text.UTF8Encoding]::new($false))
        }
    }
    elseif ($ok1) {
        Write-Log "  [2/3] Marker 단독 출력"
        [System.IO.File]::WriteAllText($outputFile,
            [System.IO.File]::ReadAllText($md1[0].FullName, [System.Text.UTF8Encoding]::new($false)),
            [System.Text.UTF8Encoding]::new($false))
    }
    elseif ($ok2) {
        Write-Log "  [2/3] MinerU 단독 출력"
        [System.IO.File]::WriteAllText($outputFile,
            [System.IO.File]::ReadAllText($md2[0].FullName, [System.Text.UTF8Encoding]::new($false)),
            [System.Text.UTF8Encoding]::new($false))
        $imgs = @(Get-ChildItem $mnrDir -Recurse -Include "*.png","*.jpg","*.jpeg" -EA SilentlyContinue)
        if ($imgs.Count -gt 0) {
            $id = Join-Path $OutputDir "${baseName}_images"
            New-Item -Path $id -ItemType Directory -Force | Out-Null
            foreach ($im in $imgs) { Copy-Item $im.FullName $id -Force }
        }
    }
    else {
        Write-Log "  [FAIL] 양쪽 모두 실패"
        Remove-Item $mrkDir -Recurse -Force -EA SilentlyContinue
        Remove-Item $mnrDir -Recurse -Force -EA SilentlyContinue
        return
    }

    Remove-Item $mrkDir -Recurse -Force -EA SilentlyContinue
    Remove-Item $mnrDir -Recurse -Force -EA SilentlyContinue

    # ── [3/3] input 정리 ──────────────────────────────────
    try {
        Remove-Item -Path $FilePath -Force
        Write-Log "  [3/3] 완료 → $outputFile (원본 삭제)"
    }
    catch { Write-Log "  [ERROR] 원본 삭제 실패: $($_.Exception.Message)" }
    Write-Log "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
}

# ══════════════════════════════════════════════════════════
#  시작
# ══════════════════════════════════════════════════════════
Write-Log ""
Write-Log "================================================"
Write-Log " PDF -> Markdown 통합 변환기"
Write-Log " Marker 구조 + MinerU 이미지 + LaTeX 통일"
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
