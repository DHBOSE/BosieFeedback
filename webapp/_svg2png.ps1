param([string]$Svg, [string]$Png, [double]$PicW = 400, [double]$PicH = 200, [int]$W = 1920, [int]$H = 1080)
$ErrorActionPreference = "Stop"
$ppt = New-Object -ComObject PowerPoint.Application
try {
    $pres = $ppt.Presentations.Add()
    $slide = $pres.Slides.Add(1, 12)
    $null = $slide.Shapes.AddPicture($Svg, $false, $true, 20, 20, $PicW, $PicH)
    $slide.Export($Png, "PNG", $W, $H)
    $pres.Close()
} finally {
    $ppt.Quit()
}
Write-Output "exported"
