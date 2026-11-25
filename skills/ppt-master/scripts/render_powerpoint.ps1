param(
    [Parameter(Mandatory=$true)][string]$InputPptx,
    [Parameter(Mandatory=$true)][string]$OutputDirectory,
    [ValidateRange(640,7680)][int]$Width = 2560,
    [ValidateRange(360,4320)][int]$Height = 1440,
    [switch]$Svg
)
# Original candidate enhancement. Opens a read-only, untitled copy; never saves the source.
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$inputFile = (Resolve-Path -LiteralPath $InputPptx).Path
$outputRoot = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $outputRoot) {
    throw 'OutputDirectory must not exist; choose a new directory for each render.'
}
$hashBefore = (Get-FileHash -LiteralPath $inputFile -Algorithm SHA256).Hash
$existingProcess = @(Get-Process POWERPNT -ErrorAction SilentlyContinue).Count -gt 0
$ppt = $null
$deck = $null
$ownsApplication = $false
$report = [ordered]@{
    schema='ppt-master.native-render.v1'; status='failed'; input=$inputFile;
    input_sha256=$hashBefore; output=$outputRoot; office_version=$null;
    slide_count=0; slide_size_points=$null; png_size=@($Width,$Height);
    slides=@(); pdf=$null; pdf_method=$null; warnings=@(); errors=@(); source_unchanged=$false
}
New-Item -ItemType Directory -Path $outputRoot | Out-Null
try {
    try { $ppt = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') }
    catch {
        $ppt = New-Object -ComObject PowerPoint.Application
        # New-Object can attach to a running Office server; never quit if one
        # was already present, even when the active-object lookup failed.
        $ownsApplication = -not $existingProcess
    }
    $report.office_version = [string]$ppt.Version
    $deck = $ppt.Presentations.Open($inputFile, -1, -1, 0)
    $report.slide_count = $deck.Slides.Count
    $report.slide_size_points = @($deck.PageSetup.SlideWidth, $deck.PageSetup.SlideHeight)
    $ratio = [double]$deck.PageSetup.SlideWidth / [double]$deck.PageSetup.SlideHeight
    if ([Math]::Abs($ratio - ([double]$Width/$Height)) -gt 0.01) {
        throw 'Requested pixel aspect ratio differs from slide ratio; pass matching Width and Height.'
    }
    for ($index=1; $index -le $deck.Slides.Count; $index++) {
        $slide = $null
        try {
            $slide = $deck.Slides.Item($index)
            $pngPath = Join-Path $outputRoot ('slide-{0:D2}.png' -f $index)
            $slide.Export($pngPath, 'PNG', $Width, $Height)
            if (-not (Test-Path -LiteralPath $pngPath)) { throw "Slide $index PNG was not written." }
            $record = [ordered]@{slide=$index; png=$pngPath; svg=$null}
            if ($Svg) {
                $svgPath = Join-Path $outputRoot ('slide-{0:D2}.svg' -f $index)
                try {
                    $slide.Export($svgPath, 'SVG', $Width, $Height)
                    if (-not (Test-Path -LiteralPath $svgPath)) { throw 'No SVG output.' }
                    $record.svg=$svgPath
                } catch {
                    $report.warnings += "Slide $index native SVG unavailable: $($_.Exception.Message)"
                }
            }
            $report.slides += $record
        } finally {
            if ($null -ne $slide) {
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($slide)
            }
        }
    }
    $pdfPath = Join-Path $outputRoot 'presentation.pdf'
    # PrintRange is a required COM argument even when it is Nothing. Supply the
    # full argument list to avoid PowerShell's optional-argument binder error.
    # Print intent, slides only, no hidden slides, and no bitmap substitution.
    try {
        $deck.ExportAsFixedFormat($pdfPath, 2, 2, 0, 1, 1, 0, $null, 1, '', $false, $true, $true, $false, $false, [Type]::Missing)
        $report.pdf_method='ExportAsFixedFormat'
    } catch {
        # Some installed Office COM type libraries cannot marshal PrintRange
        # from PowerShell. SaveAs PDF on this untitled copy is also a native
        # PDF exporter and never changes the original PPTX on disk.
        $report.warnings += "ExportAsFixedFormat unavailable; using native SaveAs PDF: $($_.Exception.Message)"
        $deck.SaveAs($pdfPath, 32)
        $report.pdf_method='SaveAs(ppSaveAsPDF)'
    }
    if (-not (Test-Path -LiteralPath $pdfPath)) { throw 'Native PDF export was not written.' }
    $report.pdf=$pdfPath
    $report.status='pass'
} catch {
    $report.errors += $_.Exception.Message
} finally {
    if ($null -ne $deck) {
        # Close only the untitled copy we opened; never close another deck.
        try { $deck.Saved = -1; $deck.Close() } catch { $report.warnings += $_.Exception.Message }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($deck)
    }
    if ($null -ne $ppt) {
        if ($ownsApplication) {
            try { if ($ppt.Presentations.Count -eq 0) { $ppt.Quit() } } catch {}
        }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ppt)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    $hashAfter = (Get-FileHash -LiteralPath $inputFile -Algorithm SHA256).Hash
    $report.source_unchanged = $hashBefore -eq $hashAfter
    if (-not $report.source_unchanged) {
        $report.status='failed'
        $report.errors += 'Source file hash changed during render.'
    }
    $json = $report | ConvertTo-Json -Depth 8
    $json | Set-Content -LiteralPath (Join-Path $outputRoot 'render-report.json') -Encoding UTF8
    Write-Output $json
}
if ($report.status -ne 'pass') { exit 1 }
