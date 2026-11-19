[CmdletBinding()]
param(
    [string]$DocxPath,
    [string]$PdfPath,
    [string]$MetadataPath
)

# 使用自己的隐藏 Word COM 实例更新新生成文档的字段并导出预览。
# 不查找/终止现有 WINWORD 进程，不打开 Markdown 或其他用户文档。
$ErrorActionPreference = 'Stop'
$reportRoot = Split-Path -Parent $PSScriptRoot
if (-not $DocxPath) { $DocxPath = Join-Path $reportRoot '实验报告_审阅版.docx' }
if (-not $PdfPath) { $PdfPath = Join-Path $reportRoot '实验报告_审阅版.pdf' }
if (-not $MetadataPath) { $MetadataPath = Join-Path $reportRoot '审阅版预览\Word分页.json' }
$docxAbsolute = (Resolve-Path -LiteralPath $DocxPath).Path
$pdfAbsolute = [System.IO.Path]::GetFullPath($PdfPath)
$metadataAbsolute = [System.IO.Path]::GetFullPath($MetadataPath)
$safeReportRoot = [System.IO.Path]::GetFullPath($reportRoot).TrimEnd('\') + '\'
foreach ($target in @($docxAbsolute, $pdfAbsolute, $metadataAbsolute)) {
    if (-not $target.StartsWith($safeReportRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Word review output must stay within the report-material directory: $target"
    }
}
if ([System.IO.Path]::GetFileName($docxAbsolute) -ne '实验报告_审阅版.docx') {
    throw 'This helper only edits the newly generated 实验报告_审阅版.docx.'
}
foreach ($folder in @((Split-Path -Parent $pdfAbsolute), (Split-Path -Parent $metadataAbsolute))) {
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
}

function Get-PhysicalPage([object]$Document, [int]$Position) {
    $r = $Document.Range($Position, $Position)
    try { return [int]$r.Information(3) } # wdActiveEndPageNumber: absolute physical page.
    finally { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($r) }
}

function Get-CleanWordText([string]$Text) {
    return ($Text -replace '[\r\a\f]+$', '').Trim()
}

$word = $null
$doc = $null
$result = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3 # msoAutomationSecurityForceDisable, no macro execution.
    # ConfirmConversions=false, ReadOnly=false, AddToRecentFiles=false.
    $doc = $word.Documents.Open($docxAbsolute, $false, $false, $false)
    $doc.Repaginate()
    [void]$doc.Fields.Update()
    foreach ($toc in $doc.TablesOfContents) { [void]$toc.Update() }
    foreach ($section in $doc.Sections) {
        foreach ($header in $section.Headers) {
            if ($header.Exists) { [void]$header.Range.Fields.Update() }
        }
        foreach ($footer in $section.Footers) {
            if ($footer.Exists) { [void]$footer.Range.Fields.Update() }
        }
    }
    $doc.Repaginate()
    [void]$doc.Fields.Update()
    $doc.Save()
    $doc.Repaginate()
    $pageCount = [int]$doc.ComputeStatistics(2) # wdStatisticPages.
    $headings = @()
    $captions = @()
    $paragraphCount = [int]$doc.Paragraphs.Count
    for ($i = 1; $i -le $paragraphCount; $i++) {
        $paragraph = $doc.Paragraphs.Item($i)
        $paragraphText = Get-CleanWordText ([string]$paragraph.Range.Text)
        if (-not $paragraphText) { continue }
        $styleName = [string]$paragraph.Style.NameLocal
        $start = [int]$paragraph.Range.Start
        $end = [int]$paragraph.Range.End
        if ([int]$paragraph.OutlineLevel -eq 1) {
            $headings += [PSCustomObject]@{
                text = $paragraphText
                style = $styleName
                start = $start
                physical_page = Get-PhysicalPage $doc $start
            }
        }
        if ($styleName -match '(?i)Caption|图注|题注|插图说明') {
            $captions += [PSCustomObject]@{
                text = $paragraphText
                style = $styleName
                start = $start
                end = $end
                physical_page_start = Get-PhysicalPage $doc $start
                physical_page_end = Get-PhysicalPage $doc ([Math]::Max($start, $end - 1))
            }
        }
    }
    $images = @()
    $imageCount = [int]$doc.InlineShapes.Count
    for ($i = 1; $i -le $imageCount; $i++) {
        $inline = $doc.InlineShapes.Item($i)
        $start = [int]$inline.Range.Start
        $end = [int]$inline.Range.End
        $afterCaption = $captions | Where-Object { $_.start -ge $end } | Sort-Object start | Select-Object -First 1
        $page = Get-PhysicalPage $doc $start
        $images += [PSCustomObject]@{
            index = $i
            type = [int]$inline.Type
            alt_text = [string]$inline.AlternativeText
            start = $start
            physical_page = $page
            width_pt = [double]$inline.Width
            height_pt = [double]$inline.Height
            following_caption = if ($afterCaption) { $afterCaption.text } else { $null }
            following_caption_page = if ($afterCaption) { $afterCaption.physical_page_start } else { $null }
            caption_on_same_page = if ($afterCaption) { $page -eq $afterCaption.physical_page_start } else { $false }
        }
    }
    $sections = @()
    for ($i = 1; $i -le $doc.Sections.Count; $i++) {
        $section = $doc.Sections.Item($i)
        $sectionStart = [int]$section.Range.Start
        $sectionEnd = [Math]::Max($sectionStart, [int]$section.Range.End - 1)
        $sections += [PSCustomObject]@{
            index = $i
            physical_page_start = Get-PhysicalPage $doc $sectionStart
            physical_page_end = Get-PhysicalPage $doc $sectionEnd
            orientation = [int]$section.PageSetup.Orientation
            width_pt = [double]$section.PageSetup.PageWidth
            height_pt = [double]$section.PageSetup.PageHeight
            top_margin_pt = [double]$section.PageSetup.TopMargin
            bottom_margin_pt = [double]$section.PageSetup.BottomMargin
            left_margin_pt = [double]$section.PageSetup.LeftMargin
            right_margin_pt = [double]$section.PageSetup.RightMargin
        }
    }
    $pages = @()
    for ($page = 1; $page -le $pageCount; $page++) {
        $pageStartRange = $doc.GoTo(1, 1, $page) # wdGoToPage, wdGoToAbsolute.
        $pageStart = [int]$pageStartRange.Start
        if ($page -lt $pageCount) {
            $nextPageRange = $doc.GoTo(1, 1, ($page + 1))
            $pageEnd = [int]$nextPageRange.Start
        } else {
            $pageEnd = [int]$doc.Content.End
        }
        $pageRange = $doc.Range($pageStart, $pageEnd)
        $pages += [PSCustomObject]@{
            physical_page = $page
            printed_page_at_start = [int]$pageStartRange.Information(1)
            section_at_start = [int]$pageStartRange.Information(2)
            start = $pageStart
            end = $pageEnd
            text = [string]$pageRange.Text
        }
    }
    # Export all content as print-quality PDF with heading bookmarks and tags.
    $doc.ExportAsFixedFormat($pdfAbsolute, 17, $false, 0, 0, 1, $pageCount, 0, $true, $true, 1, $true, $true, $false)
    $doc.Save()
    $result = [ordered]@{
        status = 'exported_with_native_word'
        word_version = [string]$word.Version
        docx = $docxAbsolute
        pdf = $pdfAbsolute
        physical_pages = $pageCount
        paragraphs = $paragraphCount
        inline_images = $imageCount
        editable_math_objects = [int]$doc.OMaths.Count
        tables = [int]$doc.Tables.Count
        heading1_pages = $headings
        image_pages = $images
        caption_pages = $captions
        sections = $sections
        pages = $pages
    }
} catch {
    $errorResult = [ordered]@{
        status = 'word_export_failed'
        error = $_.Exception.Message
        script_stack = $_.ScriptStackTrace
        docx = $docxAbsolute
        pdf = $pdfAbsolute
    }
    $errorPath = Join-Path (Split-Path -Parent $metadataAbsolute) 'Word导出错误.json'
    [System.IO.File]::WriteAllText($errorPath, ($errorResult | ConvertTo-Json -Depth 8), [System.Text.UTF8Encoding]::new($false))
    throw
} finally {
    if ($null -ne $doc) {
        try { $doc.Close(0) } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc) } catch { }
    }
    if ($null -ne $word) {
        try { $word.Quit() } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) } catch { }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
if ($null -ne $result) {
    $result.docx_sha256 = (Get-FileHash -LiteralPath $docxAbsolute -Algorithm SHA256).Hash.ToLowerInvariant()
    $result.pdf_sha256 = (Get-FileHash -LiteralPath $pdfAbsolute -Algorithm SHA256).Hash.ToLowerInvariant()
    [System.IO.File]::WriteAllText($metadataAbsolute, ($result | ConvertTo-Json -Depth 12), [System.Text.UTF8Encoding]::new($false))
    [PSCustomObject]@{
        status = $result.status
        physical_pages = $result.physical_pages
        inline_images = $result.inline_images
        editable_math_objects = $result.editable_math_objects
        pdf = $pdfAbsolute
        metadata = $metadataAbsolute
    } | ConvertTo-Json -Depth 4
}
