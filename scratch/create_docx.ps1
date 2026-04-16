try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $doc = $word.Documents.Add()
    $content = Get-Content -Path "d:\leave-request-management\documents\UserGuide.md" -Raw
    
    # Simple formatting: replace Markdown headers with bold text for visibility
    $content = $content -replace '^# (.*)$', "TITLE: `$1"
    $content = $content -replace '^## (.*)$', "SECTION: `$1"
    $content = $content -replace '^### (.*)$', "SUBSECTION: `$1"
    
    $word.Selection.TypeText($content)
    
    $outputPath = "d:\leave-request-management\documents\UserGuide.docx"
    if (Test-Path $outputPath) { Remove-Item $outputPath }
    
    $doc.SaveAs([ref]$outputPath)
    $doc.Close()
    $word.Quit()
    Write-Host "SUCCESS: UserGuide.docx created at $outputPath"
} catch {
    Write-Error "Failed to create Word document: $_"
    if ($word) { $word.Quit() }
}
