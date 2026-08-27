# 将 report.docx 转为 report.pdf（wdFormatPDF=17）
# 用法：powershell -NoProfile -ExecutionPolicy Bypass -File docx2pdf.ps1
# 要求：本机已安装 Microsoft Office / Word。若 Word 不可用，改用
#   soffice --headless --nologo --norestore --convert-to pdf --outdir <目录> report.docx
$ErrorActionPreference = "Stop"
$src = Join-Path $PSScriptRoot "report.docx"
$dst = Join-Path $PSScriptRoot "report.pdf"
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open($src, $false, $true)
    $doc.SaveAs2($dst, 17)   # 17 = wdFormatPDF
    $doc.Close($false)
    Write-Output "PDF_OK"
} finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}