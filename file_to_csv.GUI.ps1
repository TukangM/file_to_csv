Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Install-Module -Name ImportExcel
Import-Module -Name ImportExcel

# Function to execute the file-to-CSV conversion and convert to XLSX
function ConvertToCSV {
    $inputdir = $textboxInputDir.Text
    $outputfile = $textboxOutputFile.Text
    $filetype = $textboxFileType.Text

    if (-not (Test-Path -Path $inputdir -PathType Container)) {
        Write-Host "Input directory not found!" -ForegroundColor Red
        return
    }

    $consoleLog.AppendText("Please wait..." + [Environment]::NewLine)

    Get-ChildItem -Path $inputdir $filetype -Recurse | Export-Csv -Path $outputfile -Encoding ASCII -NoTypeInformation

    $consoleLog.AppendText("Done! Making .xlsx version" + [Environment]::NewLine)

    $outputfileXLSX = $outputfile -replace '.csv$', '.xlsx'
    $data = Import-Csv -Path $outputfile
    $data | Export-Excel -Path $outputfileXLSX 
    $consoleLog.AppendText("Done make .xlsx!" + [Environment]::NewLine)
}

# Function to open the folder browse dialog
function BrowseFolder {
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select Input Directory"

    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textboxInputDir.Text = $folderBrowser.SelectedPath
    }
}

# Function to open the save file dialog for output CSV file
function BrowseSaveFile {
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv"
    $saveFileDialog.FileName = "output.csv"

    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textboxOutputFile.Text = $saveFileDialog.FileName
    }
}

$form = New-Object Windows.Forms.Form
$form.Text = "File to CSV and XLSX Converter"
$form.Size = New-Object Drawing.Size(500, 300)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.Topmost = $true

$labelInputDir = New-Object Windows.Forms.Label
$labelInputDir.Text = "Input Directory:"
$labelInputDir.Location = New-Object Drawing.Point(15, 15)
$form.Controls.Add($labelInputDir)

$textboxInputDir = New-Object Windows.Forms.TextBox
$textboxInputDir.Location = New-Object Drawing.Point(120, 15)
$textboxInputDir.Size = New-Object Drawing.Size(250, 20)
$form.Controls.Add($textboxInputDir)

$buttonBrowseInput = New-Object Windows.Forms.Button
$buttonBrowseInput.Text = "Browse"
$buttonBrowseInput.Location = New-Object Drawing.Point(375, 15)
$buttonBrowseInput.Add_Click({ BrowseFolder })
$form.Controls.Add($buttonBrowseInput)

$labelOutputFile = New-Object Windows.Forms.Label
$labelOutputFile.Text = "Output CSV File:"
$labelOutputFile.Location = New-Object Drawing.Point(15, 45)
$form.Controls.Add($labelOutputFile)

$textboxOutputFile = New-Object Windows.Forms.TextBox
$textboxOutputFile.Location = New-Object Drawing.Point(120, 45)
$textboxOutputFile.Size = New-Object Drawing.Size(250, 20)
$form.Controls.Add($textboxOutputFile)

$buttonBrowseOutput = New-Object Windows.Forms.Button
$buttonBrowseOutput.Text = "Browse"
$buttonBrowseOutput.Location = New-Object Drawing.Point(375, 45)
$buttonBrowseOutput.Add_Click({ BrowseSaveFile })
$form.Controls.Add($buttonBrowseOutput)

$labelFileType = New-Object Windows.Forms.Label
$labelFileType.Text = "File Type:"
$labelFileType.Location = New-Object Drawing.Point(15, 75)
$form.Controls.Add($labelFileType)

$textboxFileType = New-Object Windows.Forms.TextBox
$textboxFileType.Text = "*"
$textboxFileType.Location = New-Object Drawing.Point(120, 75)
$textboxFileType.Size = New-Object Drawing.Size(100, 20)
$form.Controls.Add($textboxFileType)

$buttonExecute = New-Object Windows.Forms.Button
$buttonExecute.Text = "Execute!"
$buttonExecute.Location = New-Object Drawing.Point(15, 105)
$buttonExecute.Add_Click({ ConvertToCSV })
$form.Controls.Add($buttonExecute)

$consoleLog = New-Object Windows.Forms.TextBox
$consoleLog.Multiline = $true
$consoleLog.ScrollBars = "Vertical"
$consoleLog.Location = New-Object Drawing.Point(15, 140)
$consoleLog.Size = New-Object Drawing.Size(460, 120)
$form.Controls.Add($consoleLog)

$form.ShowDialog()
