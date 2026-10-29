try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    Write-Output "MS Word COM interface initialized"
    $docxPath = "C:\Users\Arindam\Desktop\blockchain-order-tracking\BCSE329L_Final_Course_Project_Report.docx"
    $pdfPath = "C:\Users\Arindam\Desktop\blockchain-order-tracking\BCSE329L_Final_Course_Project_Report.pdf"
    $doc = $word.Documents.Open($docxPath)
    $wdFormatPDF = 17
    $doc.SaveAs([ref]$pdfPath, [ref]$wdFormatPDF)
    $wdStatisticPages = 2
    $pages = $doc.ComputeStatistics($wdStatisticPages)
    $doc.Close()
    $word.Quit()
    Write-Output "Successfully exported Final Report PDF! Page count: $pages"
} catch {
    Write-Output "Error during conversion: $_"
}
