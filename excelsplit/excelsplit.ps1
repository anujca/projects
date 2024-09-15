# Split-ExcelFile.ps1

Param(
    [string]$InputExcelFile,
    [string]$TabName,
    [string]$ColumnName,
    [string]$Password
)

Write-Host "Starting the Excel splitting process..." -ForegroundColor Green

# Set the path to the configuration file
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$configFilePath = Join-Path $scriptDir "excelsplit.conf"

# Initialize default values
$defaultInputExcelFile = ""
$defaultTabName = ""
$defaultColumnName = ""
$defaultPassword = ""
$logFileName = "output.log"
$logLevel = "INFO"  # Default log level
$filePrefix = ""    # Custom file prefix (optional)

# Read defaults from the configuration file if it exists
if (Test-Path $configFilePath) {
    Write-Host "Reading default values from configuration file 'excelsplit.conf'."
    $configContent = Get-Content $configFilePath | Where-Object { $_ -match '=' }
    foreach ($line in $configContent) {
        $key, $value = $line -split '=', 2
        $key = $key.Trim()
        $value = $value.Trim()
        switch ($key) {
            "InputExcelFile" { $defaultInputExcelFile = $value }
            "TabName" { $defaultTabName = $value }
            "ColumnName" { $defaultColumnName = $value }
            "Password" { $defaultPassword = $value }
            "LogFileName" { $logFileName = $value }
            "LogLevel" { $logLevel = $value.ToUpper() }
            "FilePrefix" { $filePrefix = $value }
            default {
                Write-Warning "Unknown configuration key '$key' in 'excelsplit.conf'."
            }
        }
    }
} else {
    Write-Host "Configuration file 'excelsplit.conf' not found. Default settings will be used."
}

# Set base output folder to the script's directory
$baseOutputFolder = $scriptDir

# Clean FilePrefix for use in folder and filenames
$cleanPrefix = $filePrefix
if (-not $cleanPrefix) {
    $cleanPrefix = $ColumnName.Trim()
}
$invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
foreach ($char in $invalidChars) {
    $cleanPrefix = $cleanPrefix -replace [RegEx]::Escape($char), '_'
}
if (-not $cleanPrefix) {
    $cleanPrefix = "Output"
}

# Create output folder with format <FilePrefix><timestamp>
$timestamp = Get-Date -Format 'yyyyMMddHHmmss'
$outputFolderName = "${cleanPrefix}_${timestamp}"
$outputFolder = Join-Path $baseOutputFolder $outputFolderName
Write-Host "Creating output folder for this run at '$outputFolder'..."
if (-not (Test-Path $outputFolder)) {
    New-Item -ItemType Directory -Path $outputFolder -Force | Out-Null
}

# Set log file path
$logFilePath = Join-Path $outputFolder $logFileName

# Function to check if the log level permits logging a message
function Should-LogMessage {
    param (
        [string]$messageLevel
    )
    $levels = @("ERROR", "INFO", "DEBUG")
    return ($levels.IndexOf($messageLevel) -le $levels.IndexOf($logLevel))
}

# Function to write messages to both console and log file
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO",
        [ConsoleColor]$Color = "White"
    )
    if (Should-LogMessage $Level) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logMessage = "[$timestamp] [$Level] $Message"
        Write-Host $Message -ForegroundColor $Color
        Add-Content -Path $logFilePath -Value $logMessage
    }
}

# Function to write error messages to both console and log file
function Write-ErrorLog {
    param (
        [string]$Message
    )
    Write-Log $Message "ERROR" "Red"
}

# Log the log level (for debugging purposes)
Write-Log "Log level is set to $logLevel" "DEBUG"

Write-Log "Starting the Excel splitting process..."

# Prompt for InputExcelFile if not provided
if (-not $InputExcelFile) {
    $promptMessage = "Enter the path to the input Excel file"
    if ($defaultInputExcelFile) {
        $promptMessage += " [default: $defaultInputExcelFile]"
    }
    $InputExcelFile = Read-Host $promptMessage
    if (-not $InputExcelFile -and $defaultInputExcelFile) {
        $InputExcelFile = $defaultInputExcelFile
    }
}

# Resolve the input file path
if (-not [System.IO.Path]::IsPathRooted($InputExcelFile)) {
    # The path is relative; combine it with the script's directory
    $InputExcelFile = Join-Path $scriptDir $InputExcelFile
}

# Check if the input Excel file exists
if (-not (Test-Path $InputExcelFile)) {
    Write-ErrorLog "Input Excel file '$InputExcelFile' not found."
    Read-Host "Press Enter to exit..."
    exit
}

# Prompt for TabName if not provided
if (-not $TabName) {
    $promptMessage = "Enter the worksheet (tab) name"
    if ($defaultTabName) {
        $promptMessage += " [default: $defaultTabName]"
    }
    $TabName = Read-Host $promptMessage
    if (-not $TabName -and $defaultTabName) {
        $TabName = $defaultTabName
    }
}

# Prompt for ColumnName if not provided
if (-not $ColumnName) {
    $promptMessage = "Enter the column name to split the data on"
    if ($defaultColumnName) {
        $promptMessage += " [default: $defaultColumnName]"
    }
    $ColumnName = Read-Host $promptMessage
    if (-not $ColumnName -and $defaultColumnName) {
        $ColumnName = $defaultColumnName
    }
}

# Function to prompt for a password if it fails
function Prompt-ForPassword {
    return Read-Host "Enter the password for the Excel file (leave blank if not password-protected)" -AsSecureString
}

# Try opening the Excel file with the supplied or default password
function Try-OpenExcelFile {
    param (
        [string]$excelFilePath,
        [string]$password
    )
    try {
        if ($password) {
            $workbook = $excel.Workbooks.Open($excelFilePath, 0, $false, 5, $password)
        } else {
            $workbook = $excel.Workbooks.Open($excelFilePath)
        }
        return $workbook
    } catch {
        return $null
    }
}

Write-Log "Opening Excel application..."

# Check if Excel is installed
try {
    $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
} catch {
    Write-ErrorLog "Microsoft Excel is not installed on this system."
    Read-Host "Press Enter to exit..."
    exit
}

$excel.Visible = $false
$excel.DisplayAlerts = $false

# Attempt to open the Excel file with the password from config or input
$workbook = $null
$maxAttempts = 3
$attempts = 0
$Password = $defaultPassword

while (-not $workbook -and $attempts -lt $maxAttempts) {
    if (-not $Password) {
        $Password = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR((Prompt-ForPassword)))
    }

    $workbook = Try-OpenExcelFile -excelFilePath $InputExcelFile -password $Password

    if (-not $workbook) {
        Write-ErrorLog "Failed to open Excel file. Incorrect password. Attempt $($attempts + 1) of $maxAttempts."
        if ($attempts -lt ($maxAttempts - 1)) {
            Write-Host "Please try again."
            $Password = $null  # Reset password to prompt again
        }
    }

    $attempts++
}

if (-not $workbook) {
    Write-ErrorLog "Failed to open the Excel file after $maxAttempts attempts."
    Read-Host "Press Enter to exit..."
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
    exit
}

Write-Log "Excel file opened successfully."

# Find the specified worksheet (case-insensitive, trimmed)
$worksheet = $null
foreach ($sheet in $workbook.Worksheets) {
    if ($sheet.Name.Trim().ToLower() -eq $TabName.Trim().ToLower()) {
        $worksheet = $sheet
        break
    }
}

if (-not $worksheet) {
    Write-ErrorLog "Worksheet '$TabName' not found."
    Read-Host "Press Enter to exit..."
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
    exit
}

Write-Log "Worksheet '$($worksheet.Name)' found. Proceeding with data splitting..."

# Get header row and column index (case-insensitive, trimmed)
$usedRange = $worksheet.UsedRange
$headerRow = $usedRange.Rows.Item(1)
$headers = @()
for ($col = 1; $col -le $headerRow.Columns.Count; $col++) {
    $headers += $headerRow.Columns.Item($col).Text.Trim()
}

# Initialize columnIndex to -1
$columnIndex = -1

# Loop through headers to find the first match of ColumnName
for ($i = 0; $i -lt $headers.Count; $i++) {
    if ($headers[$i].ToLower() -eq $ColumnName.Trim().ToLower()) {
        $columnIndex = $i + 1  # Excel columns are 1-based
        break  # Use the first matching column
    }
}

if ($columnIndex -eq -1) {
    Write-ErrorLog "Column '$ColumnName' not found in worksheet '$($worksheet.Name)'."
    Read-Host "Press Enter to exit..."
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
    exit
}

Write-Log "Column '$($headers[$columnIndex - 1])' found at index $columnIndex. Splitting data..." "DEBUG"

# Create a hashtable to hold data groups
$data = @{}
$rowCount = $usedRange.Rows.Count

# Loop through each row starting from row 2 (assuming row 1 is header)
for ($row = 2; $row -le $rowCount; $row++) {
    $cell = $worksheet.Cells.Item($row, $columnIndex)
    $value = $cell.Text.Trim()

    if (-not $value) {
        $value = "unknown"
    }

    # Clean value to replace invalid filename characters with an underscore
    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
    $cleanValue = $value
    foreach ($char in $invalidChars) {
        $cleanValue = $cleanValue -replace [RegEx]::Escape($char), '_'
    }
    if (-not $cleanValue) {
        $cleanValue = "unknown"
    }

    if (-not $data.ContainsKey($cleanValue)) {
        $data[$cleanValue] = New-Object System.Collections.ArrayList
    }
    $data[$cleanValue].Add($row)
}

$totalGroups = $data.Keys.Count
Write-Log "Data will be split into $totalGroups files based on the distinct values in '$ColumnName'."

# Hashtable to keep track of filename counts for duplicates
$filenameCounts = @{}
$groupIndex = 1

foreach ($key in $data.Keys) {
    $rows = $data[$key]

    if (-not $key) {
        $key = "unknown"
    }

    # Prepare filename components
    $cleanKey = $key
    if (-not $cleanKey) {
        $cleanKey = "unknown"
    }

    $filenameBase = "${cleanPrefix}_${cleanKey}"

    if ($filenameBase.StartsWith("_")) {
        $filenameBase = $filenameBase.TrimStart('_')
    }

    # Handle duplicate filenames by adding a numerical sequence
    if ($filenameCounts.ContainsKey($filenameBase)) {
        $filenameCounts[$filenameBase] += 1
        $sequence = $filenameCounts[$filenameBase]
        $filename = "${filenameBase}_${sequence}.xlsx"
    } else {
        $filenameCounts[$filenameBase] = 1
        $filename = "${filenameBase}.xlsx"
    }
    $filepath = Join-Path $outputFolder $filename

    Write-Log "[$groupIndex/$totalGroups] Saving group '$key' to '$filename'..."

    # Create new workbook and copy data
    $newWorkbook = $excel.Workbooks.Add()
    $newSheet = $newWorkbook.Sheets.Item(1)

    # Copy header row with formatting
    $headerRange = $worksheet.Range($worksheet.Cells.Item(1,1), $worksheet.Cells.Item(1,$usedRange.Columns.Count))
    [void]$headerRange.Copy($newSheet.Cells.Item(1,1))  # Suppressing output

    # Find the S.No. column index (case-insensitive), using the first match
    $sNoColumnIndex = -1
    for ($i = 0; $i -lt $headers.Count; $i++) {
        if ($headers[$i].Trim().ToLower() -eq "s.no.") {
            $sNoColumnIndex = $i + 1  # Excel columns are 1-based
            break  # Use the first matching 'S.No.' column
        }
    }

    $destRow = 2
    $seqNum = 1
    foreach ($sourceRow in $rows) {
        $sourceRange = $worksheet.Range($worksheet.Cells.Item($sourceRow,1), $worksheet.Cells.Item($sourceRow,$usedRange.Columns.Count))
        [void]$sourceRange.Copy($newSheet.Cells.Item($destRow,1))  # Suppressing output

        if ($sNoColumnIndex -ge 1) {
            $newSheet.Cells.Item($destRow,$sNoColumnIndex).Value = $seqNum
        }
        $seqNum++
        $destRow++
    }

    # Adjust column widths to fit content
    [void]$newSheet.Columns.AutoFit()  # Suppressing output

    try {
        $newWorkbook.SaveAs($filepath)
        $newWorkbook.Close($false)
        Write-Log "Successfully saved '$filename'."
    } catch {
        Write-ErrorLog "Failed to save workbook '$filename': $($_.Exception.Message)"
        $newWorkbook.Close($false)
    }

    $groupIndex++
}

# Close original workbook and quit Excel
$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)

Write-Log "Data has been successfully split and saved in '$outputFolder'." -Color Green

Write-Host ""
Write-Host "Process completed. You may now close this window." -ForegroundColor Cyan
Read-Host "Press Enter to exit..."
