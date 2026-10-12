try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    Write-Output "MS Word COM is available"
    $docxPath = "C:\Users\Arindam\Desktop\blockchain-order-tracking\BCSE329P_DA9_Report3_CovantisTrust.docx"
    $pdfPath = "C:\Users\Arindam\Desktop\blockchain-order-tracking\BCSE329P_DA9_Report3_CovantisTrust.pdf"
    $doc = $word.Documents.Open($docxPath)
    $wdFormatPDF = 17
    $doc.SaveAs([ref]$pdfPath, [ref]$wdFormatPDF)
    $wdStatisticPages = 2
    $pages = $doc.ComputeStatistics($wdStatisticPages)
    $doc.Close()
    $word.Quit()
    Write-Output "Successfully exported DA-9 PDF! Page count: $pages"
} catch {
    Write-Output "Error during conversion: $_"
}
