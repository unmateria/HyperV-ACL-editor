#Requires -RunAsAdministrator

# Load Windows Forms assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Ensure Hyper-V module is available
try {
    Import-Module Hyper-V -ErrorAction Stop
} catch {
    [System.Windows.Forms.MessageBox]::Show(
        "Hyper-V module not found. Please install the Hyper-V role.",
        "Error", [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    exit
}

# ----------------------------------------------------------------------
# Helper function: Show ACL Editor form (Add / Edit)
# ----------------------------------------------------------------------
function Show-ACLEditor {
    param(
        $EditRule = $null   # Existing rule object (for editing) or $null for new rule
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = if ($EditRule) { "Edit ACL Rule" } else { "Add ACL Rule" }
    $form.Size = New-Object System.Drawing.Size(450, 460)
    $form.StartPosition = "CenterParent"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    # Action
    $lblAction = New-Object System.Windows.Forms.Label
    $lblAction.Location = New-Object System.Drawing.Point(20, 20)
    $lblAction.Size = New-Object System.Drawing.Size(100, 20)
    $lblAction.Text = "Action:"
    $form.Controls.Add($lblAction)

    $cmbAction = New-Object System.Windows.Forms.ComboBox
    $cmbAction.Location = New-Object System.Drawing.Point(150, 20)
    $cmbAction.Size = New-Object System.Drawing.Size(200, 20)
    $cmbAction.DropDownStyle = "DropDownList"
    $cmbAction.Items.AddRange(@("Allow", "Deny"))
    $form.Controls.Add($cmbAction)

    # Direction
    $lblDirection = New-Object System.Windows.Forms.Label
    $lblDirection.Location = New-Object System.Drawing.Point(20, 50)
    $lblDirection.Size = New-Object System.Drawing.Size(100, 20)
    $lblDirection.Text = "Direction:"
    $form.Controls.Add($lblDirection)

    $cmbDirection = New-Object System.Windows.Forms.ComboBox
    $cmbDirection.Location = New-Object System.Drawing.Point(150, 50)
    $cmbDirection.Size = New-Object System.Drawing.Size(200, 20)
    $cmbDirection.DropDownStyle = "DropDownList"
    $cmbDirection.Items.AddRange(@("Inbound", "Outbound"))
    $form.Controls.Add($cmbDirection)

    # Local IP Address
    $lblLocalIP = New-Object System.Windows.Forms.Label
    $lblLocalIP.Location = New-Object System.Drawing.Point(20, 80)
    $lblLocalIP.Size = New-Object System.Drawing.Size(100, 20)
    $lblLocalIP.Text = "Local IP:"
    $form.Controls.Add($lblLocalIP)

    $txtLocalIP = New-Object System.Windows.Forms.TextBox
    $txtLocalIP.Location = New-Object System.Drawing.Point(150, 80)
    $txtLocalIP.Size = New-Object System.Drawing.Size(200, 20)
    $form.Controls.Add($txtLocalIP)

    # Remote IP Address
    $lblRemoteIP = New-Object System.Windows.Forms.Label
    $lblRemoteIP.Location = New-Object System.Drawing.Point(20, 110)
    $lblRemoteIP.Size = New-Object System.Drawing.Size(100, 20)
    $lblRemoteIP.Text = "Remote IP:"
    $form.Controls.Add($lblRemoteIP)

    $txtRemoteIP = New-Object System.Windows.Forms.TextBox
    $txtRemoteIP.Location = New-Object System.Drawing.Point(150, 110)
    $txtRemoteIP.Size = New-Object System.Drawing.Size(200, 20)
    $form.Controls.Add($txtRemoteIP)

    # Local Port
    $lblLocalPort = New-Object System.Windows.Forms.Label
    $lblLocalPort.Location = New-Object System.Drawing.Point(20, 140)
    $lblLocalPort.Size = New-Object System.Drawing.Size(100, 20)
    $lblLocalPort.Text = "Local Port:"
    $form.Controls.Add($lblLocalPort)

    $txtLocalPort = New-Object System.Windows.Forms.TextBox
    $txtLocalPort.Location = New-Object System.Drawing.Point(150, 140)
    $txtLocalPort.Size = New-Object System.Drawing.Size(200, 20)
    $form.Controls.Add($txtLocalPort)

    # Remote Port
    $lblRemotePort = New-Object System.Windows.Forms.Label
    $lblRemotePort.Location = New-Object System.Drawing.Point(20, 170)
    $lblRemotePort.Size = New-Object System.Drawing.Size(100, 20)
    $lblRemotePort.Text = "Remote Port:"
    $form.Controls.Add($lblRemotePort)

    $txtRemotePort = New-Object System.Windows.Forms.TextBox
    $txtRemotePort.Location = New-Object System.Drawing.Point(150, 170)
    $txtRemotePort.Size = New-Object System.Drawing.Size(200, 20)
    $form.Controls.Add($txtRemotePort)

    # Protocol
    $lblProtocol = New-Object System.Windows.Forms.Label
    $lblProtocol.Location = New-Object System.Drawing.Point(20, 200)
    $lblProtocol.Size = New-Object System.Drawing.Size(100, 20)
    $lblProtocol.Text = "Protocol:"
    $form.Controls.Add($lblProtocol)

    $cmbProtocol = New-Object System.Windows.Forms.ComboBox
    $cmbProtocol.Location = New-Object System.Drawing.Point(150, 200)
    $cmbProtocol.Size = New-Object System.Drawing.Size(200, 20)
    $cmbProtocol.DropDownStyle = "DropDown"   # Editable
    $cmbProtocol.Items.AddRange(@("", "TCP", "UDP", "ICMP (1)"))
    $form.Controls.Add($cmbProtocol)

    # Weight
    $lblWeight = New-Object System.Windows.Forms.Label
    $lblWeight.Location = New-Object System.Drawing.Point(20, 230)
    $lblWeight.Size = New-Object System.Drawing.Size(100, 20)
    $lblWeight.Text = "Weight:"
    $form.Controls.Add($lblWeight)

    $numWeight = New-Object System.Windows.Forms.NumericUpDown
    $numWeight.Location = New-Object System.Drawing.Point(150, 230)
    $numWeight.Size = New-Object System.Drawing.Size(80, 20)
    $numWeight.Minimum = -2147483648
    $numWeight.Maximum = 2147483647
    $numWeight.Value = 1
    $form.Controls.Add($numWeight)

    # Stateful
    $chkStateful = New-Object System.Windows.Forms.CheckBox
    $chkStateful.Location = New-Object System.Drawing.Point(20, 260)
    $chkStateful.Size = New-Object System.Drawing.Size(150, 20)
    $chkStateful.Text = "Stateful"
    $form.Controls.Add($chkStateful)

    # Idle Timeout (seconds)
    $lblIdleTimeout = New-Object System.Windows.Forms.Label
    $lblIdleTimeout.Location = New-Object System.Drawing.Point(20, 290)
    $lblIdleTimeout.Size = New-Object System.Drawing.Size(120, 20)
    $lblIdleTimeout.Text = "Idle Timeout:"
    $form.Controls.Add($lblIdleTimeout)

    $numIdleTimeout = New-Object System.Windows.Forms.NumericUpDown
    $numIdleTimeout.Location = New-Object System.Drawing.Point(150, 290)
    $numIdleTimeout.Size = New-Object System.Drawing.Size(80, 20)
    $numIdleTimeout.Minimum = 0
    $numIdleTimeout.Maximum = 2147483647
    $numIdleTimeout.Value = 0
    $form.Controls.Add($numIdleTimeout)

    # Isolation ID
    $lblIsolationID = New-Object System.Windows.Forms.Label
    $lblIsolationID.Location = New-Object System.Drawing.Point(20, 320)
    $lblIsolationID.Size = New-Object System.Drawing.Size(120, 20)
    $lblIsolationID.Text = "Isolation ID:"
    $form.Controls.Add($lblIsolationID)

    $numIsolationID = New-Object System.Windows.Forms.NumericUpDown
    $numIsolationID.Location = New-Object System.Drawing.Point(150, 320)
    $numIsolationID.Size = New-Object System.Drawing.Size(80, 20)
    $numIsolationID.Minimum = 0
    $numIsolationID.Maximum = 2147483647
    $numIsolationID.Value = 0
    $form.Controls.Add($numIsolationID)

    # OK / Cancel buttons
    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Location = New-Object System.Drawing.Point(120, 370)
    $btnOK.Size = New-Object System.Drawing.Size(75, 23)
    $btnOK.Text = "OK"
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $btnOK
    $form.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(220, 370)
    $btnCancel.Size = New-Object System.Drawing.Size(75, 23)
    $btnCancel.Text = "Cancel"
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $btnCancel
    $form.Controls.Add($btnCancel)

    # Pre‑fill fields if editing
    if ($EditRule -ne $null) {
        $cmbAction.SelectedItem = $EditRule.Action.ToString()
        $cmbDirection.SelectedItem = $EditRule.Direction.ToString()
        $txtLocalIP.Text = if ($EditRule.LocalIPAddress) { $EditRule.LocalIPAddress } else { "" }
        $txtRemoteIP.Text = if ($EditRule.RemoteIPAddress) { $EditRule.RemoteIPAddress } else { "" }
        $txtLocalPort.Text = if ($EditRule.LocalPort) { $EditRule.LocalPort } else { "" }
        $txtRemotePort.Text = if ($EditRule.RemotePort) { $EditRule.RemotePort } else { "" }

        $proto = $EditRule.Protocol
        if ($proto -eq "1") { $proto = "ICMP (1)" }
        $cmbProtocol.Text = $proto

        $numWeight.Value = $EditRule.Weight
        $chkStateful.Checked = if ($EditRule.Stateful) { $true } else { $false }
        $numIdleTimeout.Value = if ($EditRule.IdleSessionTimeout) { $EditRule.IdleSessionTimeout } else { 0 }
        $numIsolationID.Value = if ($EditRule.IsolationID) { $EditRule.IsolationID } else { 0 }
    } else {
        $cmbAction.SelectedIndex = 0   # Allow
        $cmbDirection.SelectedIndex = 0 # Inbound
        $txtLocalIP.Text = ""
        $txtRemoteIP.Text = ""
        $txtLocalPort.Text = ""
        $txtRemotePort.Text = ""
        $cmbProtocol.Text = ""
        $numWeight.Value = 1
        $chkStateful.Checked = $false
        $numIdleTimeout.Value = 0
        $numIsolationID.Value = 0
    }

    $result = $form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        # Prepare output object
        $out = [PSCustomObject]@{
            Action            = $cmbAction.SelectedItem.ToString()
            Direction         = $cmbDirection.SelectedItem.ToString()
            LocalIPAddress    = $txtLocalIP.Text.Trim()
            RemoteIPAddress   = $txtRemoteIP.Text.Trim()
            LocalPort         = $txtLocalPort.Text.Trim()
            RemotePort        = $txtRemotePort.Text.Trim()
            Protocol          = $cmbProtocol.Text.Trim()
            Weight            = $numWeight.Value
            Stateful          = $chkStateful.Checked
            IdleSessionTimeout= if ($numIdleTimeout.Value -gt 0) { $numIdleTimeout.Value } else { $null }
            IsolationID       = if ($numIsolationID.Value -gt 0) { $numIsolationID.Value } else { $null }
        }

        # Convert "ICMP (1)" to "1"
        if ($out.Protocol -eq "ICMP (1)") {
            $out.Protocol = "1"
        }

        # Remove empty optional parameters (they will be omitted from the splat)
        foreach ($key in @("LocalIPAddress","RemoteIPAddress","LocalPort","RemotePort","Protocol")) {
            if ([string]::IsNullOrEmpty($out.$key)) {
                $out.PSObject.Properties.Remove($key)
            }
        }
        if ($out.IdleSessionTimeout -eq $null) { $out.PSObject.Properties.Remove("IdleSessionTimeout") }
        if ($out.IsolationID -eq $null)        { $out.PSObject.Properties.Remove("IsolationID") }

        return $out
    } else {
        return $null
    }
}

# ----------------------------------------------------------------------
# Main Form
# ----------------------------------------------------------------------
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = "Hyper-V Extended ACL Editor"
$mainForm.Size = New-Object System.Drawing.Size(1000, 600)
$mainForm.StartPosition = "CenterScreen"

# VM selection
$lblVM = New-Object System.Windows.Forms.Label
$lblVM.Location = New-Object System.Drawing.Point(10, 20)
$lblVM.Size = New-Object System.Drawing.Size(100, 20)
$lblVM.Text = "VM or Host:"
$mainForm.Controls.Add($lblVM)

$cmbVM = New-Object System.Windows.Forms.ComboBox
$cmbVM.Location = New-Object System.Drawing.Point(120, 20)
$cmbVM.Size = New-Object System.Drawing.Size(200, 20)
$cmbVM.DropDownStyle = "DropDownList"
$cmbVM.DisplayMember = "Name"
$mainForm.Controls.Add($cmbVM)

# Network adapter selection
$lblAdapter = New-Object System.Windows.Forms.Label
$lblAdapter.Location = New-Object System.Drawing.Point(10, 50)
$lblAdapter.Size = New-Object System.Drawing.Size(100, 20)
$lblAdapter.Text = "Network Adapter:"
$mainForm.Controls.Add($lblAdapter)

$cmbAdapter = New-Object System.Windows.Forms.ComboBox
$cmbAdapter.Location = New-Object System.Drawing.Point(120, 50)
$cmbAdapter.Size = New-Object System.Drawing.Size(300, 20)
$cmbAdapter.DropDownStyle = "DropDownList"
$cmbAdapter.DisplayMember = "Name"
$mainForm.Controls.Add($cmbAdapter)

# Refresh button
$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Location = New-Object System.Drawing.Point(450, 20)
$btnRefresh.Size = New-Object System.Drawing.Size(80, 23)
$btnRefresh.Text = "Refresh"
$mainForm.Controls.Add($btnRefresh)

# DataGridView for rules
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(10, 90)
$grid.Size = New-Object System.Drawing.Size(960, 400)
$grid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$grid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$grid.MultiSelect = $false
$grid.ReadOnly = $true
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$mainForm.Controls.Add($grid)

# Buttons
$btnAdd = New-Object System.Windows.Forms.Button
$btnAdd.Location = New-Object System.Drawing.Point(10, 500)
$btnAdd.Size = New-Object System.Drawing.Size(75, 23)
$btnAdd.Text = "Add"
$mainForm.Controls.Add($btnAdd)

$btnEdit = New-Object System.Windows.Forms.Button
$btnEdit.Location = New-Object System.Drawing.Point(100, 500)
$btnEdit.Size = New-Object System.Drawing.Size(75, 23)
$btnEdit.Text = "Edit"
$mainForm.Controls.Add($btnEdit)

$btnDelete = New-Object System.Windows.Forms.Button
$btnDelete.Location = New-Object System.Drawing.Point(190, 500)
$btnDelete.Size = New-Object System.Drawing.Size(75, 23)
$btnDelete.Text = "Delete"
$mainForm.Controls.Add($btnDelete)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Location = New-Object System.Drawing.Point(280, 500)
$btnClose.Size = New-Object System.Drawing.Size(75, 23)
$btnClose.Text = "Close"
$mainForm.Controls.Add($btnClose)

# Global variable to hold the currently selected adapter
$script:selectedAdapter = $null

# ----------------------------------------------------------------------
# Functions to populate the UI
# ----------------------------------------------------------------------
function Load-VMList {
    $cmbVM.Items.Clear()
    # Management OS entry
    $mgmt = [PSCustomObject]@{ Name = "Management OS"; IsManagementOS = $true }
    $cmbVM.Items.Add($mgmt) | Out-Null
    # All VMs
    $vms = Get-VM | Sort-Object Name
    foreach ($vm in $vms) {
        $cmbVM.Items.Add($vm) | Out-Null
    }
    if ($cmbVM.Items.Count -gt 0) {
        $cmbVM.SelectedIndex = 0
    }
}

function Load-AdapterList {
    $cmbAdapter.Items.Clear()
    $selected = $cmbVM.SelectedItem
    if ($selected.IsManagementOS) {
        $adapters = Get-VMNetworkAdapter -ManagementOS | Sort-Object Name
    } else {
        $adapters = Get-VMNetworkAdapter -VM $selected | Sort-Object Name
    }
    foreach ($adapter in $adapters) {
        $cmbAdapter.Items.Add($adapter) | Out-Null
    }
    if ($cmbAdapter.Items.Count -gt 0) {
        $cmbAdapter.SelectedIndex = 0
    } else {
        # Placeholder when no adapters exist
        $placeholder = [PSCustomObject]@{ Name = "No adapters"; Adapter = $null }
        $cmbAdapter.Items.Add($placeholder) | Out-Null
        $cmbAdapter.SelectedIndex = 0
    }
    $script:selectedAdapter = $null
}

function Load-ACLRules {
    $adapter = $cmbAdapter.SelectedItem
    if ($adapter -eq $null -or -not ($adapter -is [Microsoft.HyperV.PowerShell.VMNetworkAdapterBase])) {
        $grid.DataSource = $null
        $script:selectedAdapter = $null
        return
    }
    $script:selectedAdapter = $adapter

    try {
        $rules = Get-VMNetworkAdapterExtendedAcl -VMNetworkAdapter $adapter -ErrorAction Stop | Sort-Object Weight
    } catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to load ACL rules: $($_.Exception.Message)",
            "Error", [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        $rules = @()
    }

    $table = New-Object System.Data.DataTable
    $table.Columns.Add("Action", [string]) | Out-Null
    $table.Columns.Add("Direction", [string]) | Out-Null
    $table.Columns.Add("LocalIP", [string]) | Out-Null
    $table.Columns.Add("RemoteIP", [string]) | Out-Null
    $table.Columns.Add("LocalPort", [string]) | Out-Null
    $table.Columns.Add("RemotePort", [string]) | Out-Null
    $table.Columns.Add("Protocol", [string]) | Out-Null
    $table.Columns.Add("Weight", [string]) | Out-Null
    $table.Columns.Add("Stateful", [string]) | Out-Null
    $table.Columns.Add("IdleTimeout", [string]) | Out-Null
    $table.Columns.Add("IsolationID", [string]) | Out-Null
    $table.Columns.Add("Rule", [object]) | Out-Null   # hidden column

    foreach ($rule in $rules) {
        $row = $table.NewRow()
        $row["Action"] = $rule.Action
        $row["Direction"] = $rule.Direction
        $row["LocalIP"] = if ($rule.LocalIPAddress) { $rule.LocalIPAddress } else { "" }
        $row["RemoteIP"] = if ($rule.RemoteIPAddress) { $rule.RemoteIPAddress } else { "" }
        $row["LocalPort"] = if ($rule.LocalPort) { $rule.LocalPort } else { "" }
        $row["RemotePort"] = if ($rule.RemotePort) { $rule.RemotePort } else { "" }
        $row["Protocol"] = if ($rule.Protocol) { $rule.Protocol } else { "" }
        $row["Weight"] = $rule.Weight
        $row["Stateful"] = if ($rule.Stateful) { "Yes" } else { "No" }
        $row["IdleTimeout"] = if ($rule.IdleSessionTimeout) { $rule.IdleSessionTimeout } else { "" }
        $row["IsolationID"] = if ($rule.IsolationID) { $rule.IsolationID } else { "" }
        $row["Rule"] = $rule
        $table.Rows.Add($row) | Out-Null
    }

    $grid.DataSource = $table
    # Hide the "Rule" column
    if ($grid.Columns.Contains("Rule")) {
        $grid.Columns["Rule"].Visible = $false
    }
}

# ----------------------------------------------------------------------
# Event handlers
# ----------------------------------------------------------------------
$cmbVM.add_SelectedIndexChanged({
    Load-AdapterList
    Load-ACLRules
})

$cmbAdapter.add_SelectedIndexChanged({
    Load-ACLRules
})

$btnRefresh.add_Click({
    Load-VMList
    Load-AdapterList
    Load-ACLRules
})

$btnAdd.add_Click({
    if ($script:selectedAdapter -eq $null) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please select a network adapter first.",
            "Error", [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }
    $newRule = Show-ACLEditor
    if ($newRule -ne $null) {
        try {
            $params = @{
                VMNetworkAdapter = $script:selectedAdapter
                Action           = $newRule.Action
                Direction        = $newRule.Direction
                Weight           = $newRule.Weight
                Stateful         = $newRule.Stateful
            }
            if ($newRule.PSObject.Properties["LocalIPAddress"])    { $params.LocalIPAddress    = $newRule.LocalIPAddress }
            if ($newRule.PSObject.Properties["RemoteIPAddress"])   { $params.RemoteIPAddress   = $newRule.RemoteIPAddress }
            if ($newRule.PSObject.Properties["LocalPort"])         { $params.LocalPort         = $newRule.LocalPort }
            if ($newRule.PSObject.Properties["RemotePort"])        { $params.RemotePort        = $newRule.RemotePort }
            if ($newRule.PSObject.Properties["Protocol"])          { $params.Protocol          = $newRule.Protocol }
            if ($newRule.PSObject.Properties["IdleSessionTimeout"]){ $params.IdleSessionTimeout= $newRule.IdleSessionTimeout }
            if ($newRule.PSObject.Properties["IsolationID"])       { $params.IsolationID       = $newRule.IsolationID }

            Add-VMNetworkAdapterExtendedAcl @params -ErrorAction Stop
            Load-ACLRules
            [System.Windows.Forms.MessageBox]::Show(
                "Rule added successfully.",
                "Success", [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to add rule: $($_.Exception.Message)",
                "Error", [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    }
})

$btnEdit.add_Click({
    if ($grid.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please select a rule to edit.",
            "Error", [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }
    $selectedRow = $grid.SelectedRows[0]
    $ruleObj = $selectedRow.DataBoundItem["Rule"]
    if ($ruleObj -eq $null) {
        [System.Windows.Forms.MessageBox]::Show(
            "No rule selected.",
            "Error", [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }
    $editedRule = Show-ACLEditor -EditRule $ruleObj
    if ($editedRule -ne $null) {
        try {
            # Remove old rule
            $ruleObj | Remove-VMNetworkAdapterExtendedAcl -ErrorAction Stop
            # Add new rule
            $params = @{
                VMNetworkAdapter = $script:selectedAdapter
                Action           = $editedRule.Action
                Direction        = $editedRule.Direction
                Weight           = $editedRule.Weight
                Stateful         = $editedRule.Stateful
            }
            if ($editedRule.PSObject.Properties["LocalIPAddress"])    { $params.LocalIPAddress    = $editedRule.LocalIPAddress }
            if ($editedRule.PSObject.Properties["RemoteIPAddress"])   { $params.RemoteIPAddress   = $editedRule.RemoteIPAddress }
            if ($editedRule.PSObject.Properties["LocalPort"])         { $params.LocalPort         = $editedRule.LocalPort }
            if ($editedRule.PSObject.Properties["RemotePort"])        { $params.RemotePort        = $editedRule.RemotePort }
            if ($editedRule.PSObject.Properties["Protocol"])          { $params.Protocol          = $editedRule.Protocol }
            if ($editedRule.PSObject.Properties["IdleSessionTimeout"]){ $params.IdleSessionTimeout= $editedRule.IdleSessionTimeout }
            if ($editedRule.PSObject.Properties["IsolationID"])       { $params.IsolationID       = $editedRule.IsolationID }

            Add-VMNetworkAdapterExtendedAcl @params -ErrorAction Stop
            Load-ACLRules
            [System.Windows.Forms.MessageBox]::Show(
                "Rule updated successfully.",
                "Success", [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to update rule: $($_.Exception.Message)",
                "Error", [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            Load-ACLRules   # refresh to reflect any partial change
        }
    }
})

$btnDelete.add_Click({
    if ($grid.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please select a rule to delete.",
            "Error", [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }
    $selectedRow = $grid.SelectedRows[0]
    $ruleObj = $selectedRow.DataBoundItem["Rule"]
    if ($ruleObj -eq $null) {
        [System.Windows.Forms.MessageBox]::Show(
            "No rule selected.",
            "Error", [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }
    $answer = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to delete the selected rule?",
        "Confirm Delete", [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($answer -eq "Yes") {
        try {
            $ruleObj | Remove-VMNetworkAdapterExtendedAcl -ErrorAction Stop
            Load-ACLRules
            [System.Windows.Forms.MessageBox]::Show(
                "Rule deleted successfully.",
                "Success", [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to delete rule: $($_.Exception.Message)",
                "Error", [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    }
})

$btnClose.add_Click({
    $mainForm.Close()
})

# ----------------------------------------------------------------------
# Initialize and show form
# ----------------------------------------------------------------------
Load-VMList
Load-AdapterList
Load-ACLRules

# Display the form
$mainForm.ShowDialog() | Out-Null
