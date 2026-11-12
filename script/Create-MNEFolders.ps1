<#
.SYNOPSIS
    Creates a folder (with standard subfolders) for every MNE listed in an Excel workbook.

.DESCRIPTION
    Reads the worksheet named "MNE" in the source workbook. Cell A1 must contain the
    header "MNE"; every non-empty cell below it in column A is treated as an MNE.
    For each MNE a folder is created under the target location, containing the
    subfolders: Documents, Analysis, Comms, Evidence, Misc.

    Uses the ImportExcel module if it is installed (no Excel needed); otherwise falls
    back to Microsoft Excel via COM automation.

.PARAMETER Source
    Full path to the Excel workbook (.xlsx / .xls).

.PARAMETER Target
    Folder in which the MNE folders will be created. Created if it does not exist.

.EXAMPLE
    .\Create-MNEFolders.ps1 -Source "C:\Data\MNE_List.xlsx" -Target "D:\Cases"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Source,

    [Parameter(Mandatory = $true)]
    [string]$Target
)

$ErrorActionPreference = 'Stop'

$SheetName  = 'MNE'
$HeaderName = 'MNE'
$SubFolders = @('Documents', 'Analysis', 'Comms', 'Evidence', 'Misc')

# --- Validate inputs -------------------------------------------------------
if (-not (Test-Path -LiteralPath $Source -PathType Leaf)) {
    throw "Source file not found: $Source"
}
$Source = (Resolve-Path -LiteralPath $Source).Path

if (-not (Test-Path -LiteralPath $Target)) {
    New-Item -ItemType Directory -Path $Target -Force | Out-Null
}

# --- Read the MNE list from the spreadsheet ---------------------------------
function Get-MNEList {
    param([string]$Path)

    $values = @()

    if (Get-Module -ListAvailable -Name ImportExcel) {
        Import-Module ImportExcel
        $header = (Import-Excel -Path $Path -WorksheetName $SheetName -NoHeader -StartRow 1 -EndRow 1 -StartColumn 1 -EndColumn 1).P1
        if ("$header".Trim() -ne $HeaderName) {
            throw "Cell A1 on sheet '$SheetName' must be '$HeaderName' (found '$header')."
        }
        $rows = Import-Excel -Path $Path -WorksheetName $SheetName -StartColumn 1 -EndColumn 1
        foreach ($row in $rows) { $values += "$($row.$HeaderName)" }
    }
    else {
        $excel = $null; $workbook = $null; $sheet = $null
        try {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            $workbook = $excel.Workbooks.Open($Path, 0, $true)   # read-only
            $sheet = $workbook.Worksheets.Item($SheetName)

            $header = $sheet.Cells.Item(1, 1).Text
            if ("$header".Trim() -ne $HeaderName) {
                throw "Cell A1 on sheet '$SheetName' must be '$HeaderName' (found '$header')."
            }

            # Cycle down column A until the list runs out (first empty cell).
            $row = 2
            while ($true) {
                $value = $sheet.Cells.Item($row, 1).Text
                if ([string]::IsNullOrWhiteSpace($value)) { break }
                $values += $value
                $row++
            }
        }
        finally {
            if ($workbook) { $workbook.Close($false) }
            if ($excel)    { $excel.Quit() }
            foreach ($o in @($sheet, $workbook, $excel)) {
                if ($o) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) }
            }
            [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        }
    }

    # Trim, drop blanks, remove duplicates
    $values | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -Unique
}

$mneList = @(Get-MNEList -Path $Source)

if ($mneList.Count -eq 0) {
    Write-Warning "No MNEs found under the '$HeaderName' header on sheet '$SheetName'."
}

# --- Create folders ---------------------------------------------------------
$invalidChars = [System.IO.Path]::GetInvalidFileNameChars()

foreach ($mne in $mneList) {
    # Replace any characters that are not allowed in folder names
    $folderName = -join ($mne.ToCharArray() | ForEach-Object { if ($invalidChars -contains $_) { '_' } else { $_ } })
    $mnePath = Join-Path -Path $Target -ChildPath $folderName

    New-Item -ItemType Directory -Path $mnePath -Force | Out-Null
    foreach ($sub in $SubFolders) {
        New-Item -ItemType Directory -Path (Join-Path $mnePath $sub) -Force | Out-Null
    }
    Write-Host "Created: $mnePath"
}

Write-Host "Done"
