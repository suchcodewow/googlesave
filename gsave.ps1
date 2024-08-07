$driveFolder = "G:\My Drive\ImportedDownloads"
$downloadsFolder = "$env:USERPROFILE\downloads"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Prep
if (-Not (test-path $driveFolder)) { new-item -path $driveFolder -ItemType "directory" }

# Functions
Function Add-toGoogle() {
    $fileName = $listBox.SelectedItems[0].SubItems[2].Text
    $fileExt = [System.IO.Path]::GetExtension($fileName)
    $fileBase = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
    write-host $fileBase
    if ([System.IO.Path]::GetExtension($fileName) -eq ".zip") {
        # write-host "Zip file"
        Expand-Archive -LiteralPath $fileName -Destination "$driveFolder\$fileBase"
    }
    else {
        # write-host "normal file"
        Copy-Item -Path $fileName -Destination $driveFolder
    }

    Copy-Item -Path $fileName -Destination $driveFolder
    $listBox.SelectedItems[0].SubItems[3].Text = ("✅")

}

$allDownloads = get-ChildItem -Path $downloadsFolder | Sort-Object LastWriteTime -Descending | Select-Object -first 20
$GDownloads = Get-ChildItem -Path $driveFolder
# $allDownloads
$listBox = New-Object System.Windows.Forms.ListView
$ListBox.Location = '5,5'
# $ListBox.Size = '780,380'
$ListBox.Width = '1000'
$ListBox.Height = '700'
$listBox.View = 'Details'
$listBox.FullRowSelect = $True
$listBox_Column1 = New-Object System.Windows.Forms.ColumnHeader
$listBox_Column1.Text = "Name"
$listBox_Column1.Width = -1
$listBox_Column2 = New-Object System.Windows.Forms.ColumnHeader
$listBox_Column2.Text = "Date"
$listBox_Column2.Width = -1
$listBox_Column3 = New-Object System.Windows.Forms.ColumnHeader
$listBox_Column3.Text = "FileName"
$listBox_Column3.Width = 0
$listBox_Column4 = New-Object System.Windows.Forms.ColumnHeader
$listBox_Column4.Text = "Sent"
$listBox_Column4.Width = 50
$listBox.Columns.AddRange(@(
        $listBox_Column1,
        $listBox_Column2,
        $listBox_Column3,
        $listBox_Column4)

)
foreach ($download in $allDownloads) {
    $entry = New-Object System.Windows.Forms.ListViewItem($download.Name)
    [void]$entry.SubItems.Add($download.LastWriteTime.toString())
    [void]$entry.SubItems.Add($download.FullName)
    $destinationName = "$driveFolder\$($download.Name)"
    write-host "destination file is" $destinationName
    if (test-Path $destinationName) {
        # write-host "File Exists"
        [void]$entry.SubItems.Add("✅")
    }
    else {
        # write-host "new file"
        [void]$entry.SubItems.Add("")
    }
    [void]$entry.SubItems.Add("")
    [void]$listBox.Items.Add($entry)

}
$Form = New-Object Windows.Forms.Form
$Form.Text = "Send to Google"
$Form.Width = 1020
$Form.Height = 800
$Form.BackColor = "gray"
$Form.StartPosition = "CenterScreen"
$Form.Controls.Add($listBox)
$listBox.Add_ItemActivate({ Add-toGoogle })
$Form.Add_Shown({ $Form.Activate() })
$Form.ShowDialog()