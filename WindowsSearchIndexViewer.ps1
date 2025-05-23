# Windows Search Index Viewer - PowerShell WinForms GUI
# Author: [Your Name]
# Date: [Today's Date]
# Description: Interrogates the Windows Search Index and displays a collapsible tree view of indexed file paths.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Query the Windows Search Index for file/folder paths
$sql = "SELECT System.ItemName, System.ItemPathDisplay FROM SYSTEMINDEX"
$provider = "Provider=Search.CollatorDSO;Extended Properties='Application=Windows';"
$connector = New-Object System.Data.OleDb.OleDbDataAdapter -ArgumentList $sql, $provider
$dataset = New-Object System.Data.DataSet

if ($connector.Fill($dataset)) {
    $index = $dataset.Tables[0]
    $outputFile = "IndexedItems.txt"
    # Extract the path column (second column)
    $colPath = $index.Columns[1].ColumnName
    $lines = $index | ForEach-Object {
        $path = $_[$colPath]
        if ($path) { "$path" }
    }
    $lines | Set-Content -Encoding UTF8 -Path $outputFile
    Write-Output "Windows Search Index interrogated. Output written to: $outputFile"
} else {
    Write-Output "No results returned from Windows Search Index."
}

Write-Output "Building tree structure from: $outputFile"
$txtFile = "IndexedItems.txt"

# Check for output file
if (-not (Test-Path $txtFile)) {
    [System.Windows.Forms.MessageBox]::Show("File not found: $txtFile")
    return ''
}

# Read all indexed paths
$paths = Get-Content $txtFile | Where-Object { $_ -and $_.Trim() -ne '' }

# Build a nested hashtable tree from the paths
$tree = @{}
foreach ($path in $paths) {
    $parts = $path -split '[\\/]' | Where-Object { $_ -ne '' }
    $current = $tree
    foreach ($part in $parts) {
        if (-not $current.ContainsKey($part)) {
            $current[$part] = @{ children = @{}; count = 0 }
        }
        $current[$part].count++
        $current = $current[$part].children
    }
}

# --- WinForms GUI Setup ---

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Windows Search Index"
$form.Width = 900
$form.Height = 700

# Create the TreeView for displaying the folder/file tree
$treeView = New-Object System.Windows.Forms.TreeView
$treeView.Dock = 'Fill'
$treeView.HideSelection = $false

# Create a ComboBox to select sorting mode
$comboBox = New-Object System.Windows.Forms.ComboBox
$comboBox.Items.AddRange(@('Alphabetical', 'By Count'))
$comboBox.SelectedIndex = 0
$comboBox.Dock = [System.Windows.Forms.DockStyle]::Top

# Add controls to the form
$form.Controls.Add($treeView)
$form.Controls.Add($comboBox)
$form.Controls.SetChildIndex($comboBox, 0)

# Recursively add nodes to the TreeView, supporting two sort modes
function Add-Nodes {
    param($parentNode, $tree, $sortMode)
    if ($sortMode -eq 'By Count') {
        foreach ($key in ($tree.Keys | Sort-Object -Property @{ Expression = { $tree[$_].count }; Descending = $true }, @{ Expression = { $_ }; Descending = $false })) {
            $childNode = New-Object System.Windows.Forms.TreeNode $key
            Add-Nodes $childNode $tree[$key].children $sortMode
            $parentNode.Nodes.Add($childNode) | Out-Null
            if ($childNode.Nodes.Count -gt 0) {
                $childNode.Text += " (" + $childNode.GetNodeCount($true) + ")"
            }
        }
    } else {
        foreach ($key in ($tree.Keys | Sort-Object)) {
            $childNode = New-Object System.Windows.Forms.TreeNode $key
            Add-Nodes $childNode $tree[$key].children $sortMode
            $parentNode.Nodes.Add($childNode) | Out-Null
            if ($childNode.Nodes.Count -gt 0) {
                $childNode.Text += " (" + $childNode.GetNodeCount($true) + ")"
            }
        }
    }
}

# Refresh the tree view based on the selected sort mode
function Update-Tree {
    $treeView.Nodes.Clear()
    $rootNode = New-Object System.Windows.Forms.TreeNode "Root"
    Add-Nodes $rootNode $tree $comboBox.SelectedItem
    [void]$treeView.Nodes.Add($rootNode)
    $rootNode.Expand()
}

# Initial population
Update-Tree

# Update tree when sort mode changes
$comboBox.add_SelectedIndexChanged({ Update-Tree })

# Show the form
[void]$form.ShowDialog()
