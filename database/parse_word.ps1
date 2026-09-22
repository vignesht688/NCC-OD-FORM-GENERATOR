$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $files = @(
        "c:\Users\jayav\OneDrive\Desktop\od form for ncc\7-jul-2026.docx",
        "c:\Users\jayav\OneDrive\Desktop\od form for ncc\8-jul-2026.docx",
        "c:\Users\jayav\OneDrive\Desktop\od form for ncc\9-jul-2026.docx",
        "c:\Users\jayav\OneDrive\Desktop\od form for ncc\NCC OD 1ST JUL.doc"
    )
    foreach ($file in $files) {
        Write-Output "=== FILE: $file ==="
        if (Test-Path $file) {
            $doc = $word.Documents.Open($file)
            Write-Output "Tables count: $($doc.Tables.Count)"
            if ($doc.Tables.Count -gt 0) {
                # Get the first table (which usually contains the cadets list)
                $table = $doc.Tables.Item(1)
                for ($r = 1; $r -le $table.Rows.Count; $r++) {
                    $rowCells = @()
                    for ($c = 1; $c -le $table.Columns.Count; $c++) {
                        try {
                            $cellText = $table.Cell($r, $c).Range.Text
                            # Clean word cellular characters
                            $cellText = $cellText.Replace("`r`a", "").Replace("`a", "").Replace("`r", "").Trim()
                            $rowCells += $cellText
                        } catch {
                            $rowCells += ""
                        }
                    }
                    # Print as line values separated by pipe
                    Write-Output ($rowCells -join " | ")
                }
            } else {
                Write-Output "No tables found."
            }
            $doc.Close($false)
        } else {
            Write-Output "File not found!"
        }
    }
} catch {
    Write-Error $_.Exception.Message
} finally {
    $word.Quit()
}
