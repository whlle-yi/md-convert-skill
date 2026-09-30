<#
.SYNOPSIS
    docx2pdf.ps1 - 通过 Word COM 将 .docx 转为 .pdf（Windows 零第三方依赖）。
.DESCRIPTION
    由 md_convert.py 调用；也可独立使用：

        powershell -NoProfile -ExecutionPolicy Bypass -File docx2pdf.ps1 `
            -InPath C:\path\to\in.docx -OutPath C:\path\to\out.pdf

    转换前会更新目录与全部域，保证 TOC 页码正确。
#>
param(
    [Parameter(Mandatory = $true)][string]$InPath,
    [Parameter(Mandatory = $true)][string]$OutPath
)

$ErrorActionPreference = 'Stop'
$wdExportFormatPDF = 17
$wdDoNotSaveChanges = 0

$InPath = [System.IO.Path]::GetFullPath($InPath)
$OutPath = [System.IO.Path]::GetFullPath($OutPath)
if (-not (Test-Path -LiteralPath $InPath)) {
    throw "输入文件不存在：$InPath"
}

$word = $null
$doc = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open($InPath)  # 需更新目录域，须可写打开；关闭时不保存

    foreach ($toc in @($doc.TablesOfContents)) { $toc.Update() | Out-Null }
    $doc.Fields.Update() | Out-Null
    $doc.Repaginate()

    $doc.ExportAsFixedFormat($OutPath, $wdExportFormatPDF)
    if (-not (Test-Path -LiteralPath $OutPath)) {
        throw "Word 未生成输出文件：$OutPath"
    }
    Write-Output "OK $OutPath"
}
finally {
    if ($doc -ne $null) { $doc.Close($wdDoNotSaveChanges) | Out-Null }
    if ($word -ne $null) { $word.Quit() | Out-Null }
    if ($doc -ne $null) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) }
    if ($word -ne $null) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
}
