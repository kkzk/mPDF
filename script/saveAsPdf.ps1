# Converts an Office document to PDF via COM.
# Usage: powershell.exe -NoProfile -ExecutionPolicy Bypass -File saveAsPdf.ps1 <job.json>
#
# job.json (UTF-8):
#   { "kind": "excel" | "word", "source": "<abs path>", "output": "<abs path>", "sheets": ["Sheet1", ...] }
# Paths are passed through a JSON file rather than the command line so that
# non-ASCII file names survive regardless of the console code page.
param(
    [Parameter(Mandatory = $true)]
    [string]$JobPath
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$xlTypePDF = 0
$xlQualityStandard = 0
$wdExportFormatPDF = 17

function Release-ComObject($obj) {
    if ($null -ne $obj) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj)
    }
}

function Save-Excel([string]$source, [string]$output, [string[]]$sheets) {
    $excel = $null
    $book = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        # Open(Filename, UpdateLinks:=0, ReadOnly:=True)
        $book = $excel.Workbooks.Open($source, 0, $true)

        $replace = $true
        foreach ($name in $sheets) {
            $sheet = $book.Worksheets.Item($name)
            $sheet.Select($replace)
            Release-ComObject $sheet
            $replace = $false
        }
        $book.ActiveSheet.ExportAsFixedFormat($xlTypePDF, $output, $xlQualityStandard)
    }
    finally {
        if ($null -ne $book) {
            $book.Saved = $true
            $book.Close($false)
        }
        if ($null -ne $excel) {
            $excel.Quit()
        }
        Release-ComObject $book
        Release-ComObject $excel
    }
}

function Save-Word([string]$source, [string]$output) {
    $word = $null
    $doc = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0  # wdAlertsNone
        # Open(FileName, ConfirmConversions:=False, ReadOnly:=True, AddToRecentFiles:=False)
        $doc = $word.Documents.Open($source, $false, $true, $false)
        $doc.ExportAsFixedFormat($output, $wdExportFormatPDF)
    }
    finally {
        if ($null -ne $doc) {
            $doc.Close($false)
        }
        if ($null -ne $word) {
            $word.Quit()
        }
        Release-ComObject $doc
        Release-ComObject $word
    }
}

try {
    $job = Get-Content -LiteralPath $JobPath -Raw -Encoding UTF8 | ConvertFrom-Json
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $job.output) | Out-Null

    switch ($job.kind) {
        'excel' { Save-Excel $job.source $job.output $(if ($job.sheets) { @($job.sheets) } else { @() }) }
        'word' { Save-Word $job.source $job.output }
        default { throw "Unsupported kind: $($job.kind)" }
    }
}
catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
}
finally {
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
exit 0
