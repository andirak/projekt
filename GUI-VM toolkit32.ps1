Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function CreateGreetingLabel {
    $label = New-Object Windows.Forms.Label
    $label.Text = "Hyper-V VM creation kit"
    $label.Font = New-Object Drawing.Font("Arial", 24, [Drawing.FontStyle]::Bold)
    $label.AutoSize = $true
    $label.Location = New-Object Drawing.Point(10, 10)
    $label.ForeColor = [System.Drawing.Color]::Black
    return $label
}

function CreateTextBox($text, $size, $location) {
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Text = $text
    $textBox.Multiline = $False
    $textBox.Size = New-Object System.Drawing.Size($size.Width, $size.Height)
    $textBox.Location = New-Object System.Drawing.Point($location.X, $location.Y)
    return $textBox
}

function CreateLabel($text, $location) {
    $label = New-Object Windows.Forms.Label
    $label.Text = $text
    $label.AutoSize = $true
    $label.Location = New-Object Drawing.Point($location.X, $location.Y)
    $label.ForeColor = [System.Drawing.Color]::Black
    return $label
}

function CreateComboBox($location) {
    $comboBox = New-Object System.Windows.Forms.ComboBox
    $comboBox.Text = ""
    $comboBox.Width = 100
    $comboBox.AutoSize = $true
    $comboBox.Location = New-Object System.Drawing.Point($location.X, $location.Y)
    @(2, 4, 6, 8, 10, 12, 14, 16) | ForEach-Object { [void] $comboBox.Items.Add($_) }
    $comboBox.SelectedIndex = 0
    return $comboBox
}

function ValidateInputs($vmName, $memory, $vhdSize) {
    if ([string]::IsNullOrWhiteSpace($vmName)) {
        [System.Windows.Forms.MessageBox]::Show("VM Name cannot be empty", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }
    if (-not [int]::TryParse($memory, [ref]$null)) {
        [System.Windows.Forms.MessageBox]::Show("Invalid memory size", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }
    if (-not [int]::TryParse($vhdSize, [ref]$null)) {
        [System.Windows.Forms.MessageBox]::Show("Invalid VHD size", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return $false
    }
    return $true
}

$GreetingLabel = CreateGreetingLabel

$VMNameBox = CreateTextBox "NewVM" (New-Object System.Drawing.Size 100, 20) (New-Object System.Drawing.Point 10, 150)
$VMNameLabel = CreateLabel "VM Name" (New-Object System.Drawing.Point 10, 180)

$MemoryComboBox = CreateComboBox (New-Object System.Drawing.Point 200, 170)
$MemoryLabel = CreateLabel "Memory (GB)" (New-Object System.Drawing.Point 200, 150)

$VHDSizeBox = CreateTextBox "40" (New-Object System.Drawing.Size 100, 20) (New-Object System.Drawing.Point 10, 220)
$VHDSizeLabel = CreateLabel "Virtual disk size (GB)" (New-Object System.Drawing.Point 10, 250)

$CreateButton = New-Object System.Windows.Forms.Button
$CreateButton.Location = New-Object System.Drawing.Point (200, 220)
$CreateButton.Size = New-Object System.Drawing.Size(160, 30)
$CreateButton.Font = New-Object System.Drawing.Font("Lucida Console", 18, [System.Drawing.FontStyle]::Regular)
$CreateButton.BackColor = "LightGray"
$CreateButton.Text = "Submit"
$CreateButton.Add_Click({
    $VMName = $VMNameBox.Text
    $Index = $MemoryComboBox.SelectedIndex
    $VMMem = [string]$MemoryComboBox.Items[$Index]
    $VHDX = $VHDSizeBox.Text

    if (ValidateInputs $VMName $VMMem $VHDX) {
        $VHDPath = "C:\temp\" + $VMName + ".VHDX"
        $NewVMCommand = "New-VM -Name $VMName -MemoryStartupBytes ${VMMem}GB -NewVHDPath $VHDPath -NewVHDSizeBytes ${VHDX}GB"
        Invoke-Expression $NewVMCommand
    }
})

$Form = New-Object Windows.Forms.Form
$Form.Text = "VM Creation Kit"
$Form.Width = 550
$Form.Height = 350
$Form.BackColor = "Red"

$Form.Controls.AddRange(@(
    $GreetingLabel,
    $VMNameBox,
    $VMNameLabel,
    $VHDSizeBox,
    $VHDSizeLabel,
    $MemoryComboBox,
    $MemoryLabel,
    $CreateButton
))

$Form.Add_Shown({ $Form.Activate() })
$Form.ShowDialog()
