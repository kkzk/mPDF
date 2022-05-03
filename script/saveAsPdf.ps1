using namespace System.Collections.Generic

$xlTypePDF = 0
$xlQualityStandard = 0
# $xlQualityMinimum = 1
$xlQuality = $xlQualityStandard

$wdExportFormatPDF = 17  # PDF

function Save-Word {
    [CmdletBinding()]
    param (
        [string]$jsonFilename=".mpdf.json",
        [string]$wordFilename        
    )
    
    begin {
        $wordPath = Resolve-Path $wordFilename        
        $pdfPath = Join-Path (Resolve-Path ".") ".mPDF" | join-path -ChildPath $wordFilename
        $pdfPath = [io.path]::ChangeExtension($pdfPath, ".pdf")
        Write-Host "save $wordPath as $pdfPath"
    }
    
    process {
        [System.__ComObject]$wordApplication = New-Object -ComObject Word.Application
        New-Item ($pdfPath | Split-path -Parent) -itemType Directory -Force > $null
        $wordDocument = $wordApplication.Documents.Open($wordPath.ToString(), 0, $true, $true)
        $wordDocument.ExportAsFixedFormat($pdfPath, $wdExportFormatPDF)
        $wordDocument.Close()
        $wordApplication.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordDocument)  | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordApplication) | Out-Null
        [GC]::collect()
    }
    
    end {
        
    }
}

function Save-Excel {
    [CmdletBinding()]
    param (
        [string]$jsonFilename=".mpdf.json",
        [string]$xlsFilename
    )
    
    begin {
        $xlsPath = Resolve-Path $xlsFilename
        $pdfPath = Join-Path (Resolve-Path ".") ".mPDF" | join-path -ChildPath $xlsFilename
        $pdfPath = [io.path]::ChangeExtension($pdfPath, ".pdf")
        if (Test-Path $jsonFilename) {
            $json = Get-Content $jsonFilename -Encoding UTF8 | ConvertFrom-Json
        } else {
            $json = $null
        }
        Write-Host "save $xlsPath as $pdfPath"
    }
    
    process {
        [System.__ComObject]$excel = New-Object -ComObject Excel.Application
        [System.__ComObject]$book = $excel.Workbooks.Open($xlsPath, 0, $true)

        $xls = $json | Where-Object { $_.name -eq $xlsFilename }
        $sheetNames = $xls.worksheets | Where-Object { $_.visible -eq "true" } | ForEach-Object { $_.name }
        Write-Host "  sheet: $sheetNames"
        $replase = $true
        foreach ($sheet in $book.sheets) {
            if ($sheetNames.Contains($sheet.Name)) {
                $sheet.Select($replase)
                $replase = $false
            }            
        }
        New-Item ($pdfPath | Split-path -Parent) -itemType Directory -Force > $null
        Write-Debug "`$book.ActiveSheet.ExportAsFixedFormat($xlTypePDF, $pdfPath, $xlQuality)"
        $book.ActiveSheet.ExportAsFixedFormat($xlTypePDF, $pdfPath, $xlQuality)
    }
    
    end {
        $book.Saved = $true
        $book.Close()
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)  | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($book) | Out-Null
        # [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null
        [GC]::collect()
    }
}

function Save-Pdf {
    [CmdletBinding()]
    param (
        [string]$jsonFilename=".mpdf.json",
        [string]$document        
    )
    
    begin {
        $ext = [System.IO.Path]::GetExtension($document)
    }
    
    process {
        switch ($ext) {
            ".xlsx" {
                Save-Excel $jsonFilename $document
              }
            ".docx" {
                Save-Word $jsonFilename $document
            }
            Default {}
        }
        
    }
    
    end {
        
    }
}
