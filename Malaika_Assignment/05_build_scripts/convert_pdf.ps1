$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

$files = @(
    "Malaika_MGMT268_Assessment1_FINAL.docx",
    "Malaika_MGMT268_Assessment1_WITH_VIDEO.docx"
)

$base = "C:\Users\admin\Documents\SCQF-L6-Course-Materials\Malaika_Assignment\"

foreach ($f in $files) {
    $src = $base + $f
    $pdf = $src -replace "\.docx$", ".pdf"
    if (Test-Path $src) {
        Write-Host "Converting: $f"
        $doc = $word.Documents.Open($src, $false, $true)
        Start-Sleep -Seconds 1
        $doc.SaveAs([ref]$pdf, [ref]17)
        $doc.Close([ref]$false)
        $size = [math]::Round((Get-Item $pdf).Length / 1KB)
        Write-Host "  Saved: $pdf  ($size KB)"
    } else {
        Write-Host "  Not found: $src"
    }
}

$word.Quit()
Write-Host "Done."
