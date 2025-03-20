# AD Management Tool - Tabbed GUI Interface
# This script combines User Creation, Distribution Group Management, and Contact Creation
# Author: Claude
# GitHub: https://github.com/MoshikoKar/AD-Management-Tool
# Version: 1.0.0

# Load Required Assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web

# Set Error Action Preference
$ErrorActionPreference = "Stop"

# Config and Log Paths
$ConfigFolder = "$env:APPDATA\AD-Management-Tool"
$ConfigFile = "$ConfigFolder\config.xml"
$LogsFolder = "$env:USERPROFILE\Documents\AD_Management_Logs"
$UserLogFile = "$LogsFolder\AD_User_Creation.log"
$GroupLogFile = "$LogsFolder\AD_Group_Management.log"
$ContactLogFile = "$LogsFolder\AD_Contact_Creation.log"

# Ensure Directories Exist
if (-not (Test-Path $LogsFolder)) { New-Item -ItemType Directory -Path $LogsFolder -Force | Out-Null }
if (-not (Test-Path $ConfigFolder)) { New-Item -ItemType Directory -Path $ConfigFolder -Force | Out-Null }

# Default Configuration
$Config = @{
    CompanyName = "YourCompany"
    DomainNetBIOS = $(if ($env:USERDOMAIN) { $env:USERDOMAIN } else { "DOMAIN" })
    DomainFQDN = $(if ($env:USERDNSDOMAIN) { $env:USERDNSDOMAIN } else { "domain.com" })
    UserOU = "OU=Users,DC=$(($env:USERDNSDOMAIN -split '\.')[0]),DC=$(($env:USERDNSDOMAIN -split '\.')[1])"
    GroupOU = "OU=Groups,DC=$(($env:USERDNSDOMAIN -split '\.')[0]),DC=$(($env:USERDNSDOMAIN -split '\.')[1])"
    ContactOU = "OU=Contacts,DC=$(($env:USERDNSDOMAIN -split '\.')[0]),DC=$(($env:USERDNSDOMAIN -split '\.')[1])"
    ExchangeServer = "exchange.$env:USERDNSDOMAIN"
    ExchangeVersion = "Exchange 2019"
    UserNamingPattern = "{0}.{1}" # FirstName.LastName
    ForceDefaultOU = $false
    AddressBookPathUser = "CN=All Users,CN=All Address Lists,CN=Address Lists Container,CN={0},CN=Microsoft Exchange,CN=Services,CN=Configuration,DC={1},DC={2}"
    AddressBookPathGroup = "CN=All Groups,CN=All Address Lists,CN=Address Lists Container,CN={0},CN=Microsoft Exchange,CN=Services,CN=Configuration,DC={1},DC={2}"
    EnableEmailFeatures = $true
    DefaultPrimaryDomain = $(if ($env:USERDNSDOMAIN) { $env:USERDNSDOMAIN } else { "domain.com" })
}

# Utility Functions
function Load-Configuration {
    if (Test-Path $ConfigFile) {
        try {
            $script:Config = Import-Clixml -Path $ConfigFile
            Write-Host "Configuration loaded from $ConfigFile"
            return $true
        } catch {
            Write-Warning "Error loading configuration: $_"
            return $false
        }
    }
    return $false
}

function Save-Configuration {
    try {
        $Config | Export-Clixml -Path $ConfigFile
        Write-Host "Configuration saved to $ConfigFile"
        return $true
    } catch {
        Write-Warning "Error saving configuration: $_"
        return $false
    }
}

if (-not (Load-Configuration)) {
    try {
        $domainInfo = Get-ADDomain
        $Config.DomainNetBIOS = $domainInfo.NetBIOSName
        $Config.DomainFQDN = $domainInfo.DNSRoot
        $Config.UserOU = "OU=Users," + $domainInfo.DistinguishedName
        $Config.GroupOU = "OU=Groups," + $domainInfo.DistinguishedName
        $Config.ContactOU = "OU=Contacts," + $domainInfo.DistinguishedName
        $Config.DefaultPrimaryDomain = $domainInfo.DNSRoot
    } catch {
        Write-Warning "Could not auto-detect domain information. Using defaults."
    }
    Save-Configuration
}

function Write-Log {
    param ([string]$Message, [string]$LogFile)
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$Timestamp - $Message" | Out-File -Append -FilePath $LogFile
    Write-Host "$Timestamp - $Message"
}

function Generate-RandomPassword {
    param ([int]$Length = 14, [int]$NonAlphaNumeric = 5)
    try {
        return [System.Web.Security.Membership]::GeneratePassword($Length, $NonAlphaNumeric)
    } catch {
        $CharSet = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()-_=+[]{}|;:,.<>?"
        $Password = ""
        $Random = New-Object System.Random
        $Password += "ABCDEFGHIJKLMNOPQRSTUVWXYZ"[$Random.Next(0, 26)]
        $Password += "abcdefghijklmnopqrstuvwxyz"[$Random.Next(0, 26)]
        $Password += "0123456789"[$Random.Next(0, 10)]
        $Password += "!@#$%^&*()-_=+[]{}|;:,.<>?"[$Random.Next(0, 24)]
        for ($i = 0; $i -lt ($Length - 4); $i++) { $Password += $CharSet[$Random.Next(0, $CharSet.Length)] }
        $PasswordArray = $Password.ToCharArray()
        $n = $PasswordArray.Length
        while ($n -gt 1) {
            $n--
            $k = $Random.Next(0, $n + 1)
            $temp = $PasswordArray[$k]
            $PasswordArray[$k] = $PasswordArray[$n]
            $PasswordArray[$n] = $temp
        }
        return -join $PasswordArray
    }
}

function Update-Status {
    param ([string]$Message)
    $statusLabel.Text = $Message
    $mainForm.Refresh()
}

function Get-UsernamePreview {
    param ([string]$FirstName, [string]$LastName)
    if (-not $FirstName -or -not $LastName) { return "" }
    $FirstNameParam = if ($Config.UserNamingPattern.Contains("{0}")) { $FirstName } else { $FirstName.Substring(0, 1) }
    $LastNameParam = if ($Config.UserNamingPattern.Contains("{1}")) { $LastName } else { $LastName.Substring(0, 1) }
    return ($Config.UserNamingPattern -f $FirstNameParam, $LastNameParam).ToLower()
}

# Main Form Setup
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = "AD Management Tool"
$mainForm.Size = New-Object System.Drawing.Size(600, 600)
$mainForm.StartPosition = "CenterScreen"
$mainForm.Icon = [System.Drawing.SystemIcons]::Application
$mainForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$statusBar = New-Object System.Windows.Forms.StatusStrip
$statusBar.SizingGrip = $false
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Ready"
$statusBar.Items.Add($statusLabel)
$mainForm.Controls.Add($statusBar)

$menuStrip = New-Object System.Windows.Forms.MenuStrip
$mainForm.MainMenuStrip = $menuStrip
$mainForm.Controls.Add($menuStrip)

$fileMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$fileMenu.Text = "File"
$menuStrip.Items.Add($fileMenu)

$exportMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$exportMenuItem.Text = "Export Settings..."
$exportMenuItem.Add_Click({ Export-Settings })
$fileMenu.DropDownItems.Add($exportMenuItem)

$importMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$importMenuItem.Text = "Import Settings..."
$importMenuItem.Add_Click({ Import-Settings })
$fileMenu.DropDownItems.Add($importMenuItem)

$fileMenu.DropDownItems.Add($(New-Object System.Windows.Forms.ToolStripSeparator))

$configMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$configMenuItem.Text = "Configuration..."
$configMenuItem.Add_Click({ Show-ConfigurationDialog })
$fileMenu.DropDownItems.Add($configMenuItem)

$fileMenu.DropDownItems.Add($(New-Object System.Windows.Forms.ToolStripSeparator))

$exitMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$exitMenuItem.Text = "Exit"
$exitMenuItem.Add_Click({ $mainForm.Close() })
$fileMenu.DropDownItems.Add($exitMenuItem)

$toolsMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$toolsMenu.Text = "Tools"
$menuStrip.Items.Add($toolsMenu)

$aducMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$aducMenuItem.Text = "Open Active Directory Users and Computers"
$aducMenuItem.Add_Click({ try { Start-Process "dsa.msc" } catch { [System.Windows.Forms.MessageBox]::Show("Failed to open ADUC: $_", "Error", "OK", "Error") } })
$toolsMenu.DropDownItems.Add($aducMenuItem)

$exchangeMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$exchangeMenuItem.Text = "Open Exchange Admin Center"
$exchangeMenuItem.Add_Click({ try { Start-Process "https://outlook.office365.com/ecp/" } catch { [System.Windows.Forms.MessageBox]::Show("Failed to open Exchange Admin Center: $_", "Error", "OK", "Error") } })
$toolsMenu.DropDownItems.Add($exchangeMenuItem)

$toolsMenu.DropDownItems.Add($(New-Object System.Windows.Forms.ToolStripSeparator))

$csvTemplateMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$csvTemplateMenuItem.Text = "Generate CSV Import Template"
$csvTemplateMenuItem.Add_Click({
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "CSV Files (*.csv)|*.csv"
    $saveDialog.Title = "Save CSV Template"
    $saveDialog.FileName = "User_Import_Template.csv"
    if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $templateContent = @"
FirstName,LastName,PrimaryDomain,AdditionalDomains,Description,Office,Department,JobTitle
John,Doe,$($Config.DefaultPrimaryDomain),otherdomain.com;thirddomain.com,IT Department,Main Office,IT,System Administrator
Jane,Smith,$($Config.DefaultPrimaryDomain),,Human Resources,Branch Office,HR,HR Manager
"@
        try {
            $templateContent | Out-File -FilePath $saveDialog.FileName -Encoding utf8
            [System.Windows.Forms.MessageBox]::Show("CSV template created successfully at:`n$($saveDialog.FileName)", "Template Created", "OK", "Information")
            $openResult = [System.Windows.Forms.MessageBox]::Show("Would you like to open the template file?", "Open Template", "YesNo", "Question")
            if ($openResult -eq [System.Windows.Forms.DialogResult]::Yes) { Start-Process $saveDialog.FileName }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error creating CSV template: $_", "Error", "OK", "Error")
        }
    }
})
$toolsMenu.DropDownItems.Add($csvTemplateMenuItem)

$helpMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$helpMenu.Text = "Help"
$menuStrip.Items.Add($helpMenu)

$aboutMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$aboutMenuItem.Text = "About"
$aboutMenuItem.Add_Click({
    $aboutText = @"
AD Management Tool v1.0.0

A comprehensive tool for managing Active Directory users, groups, and contacts.

Features:
- Create and manage AD users with email configuration
- Create and update distribution groups
- Create external contacts
- Export/Import settings for reuse
- Customizable for any organization

GitHub: https://github.com/MoshikoKar/AD-Management-Tool
"@
    [System.Windows.Forms.MessageBox]::Show($aboutText, "About AD Management Tool", "OK", "Information")
})
$helpMenu.DropDownItems.Add($aboutMenuItem)

$docsMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$docsMenuItem.Text = "Documentation"
$docsMenuItem.Add_Click({ Start-Process "https://github.com/MoshikoKar/AD-Management-Tool/blob/main/README.md" })
$helpMenu.DropDownItems.Add($docsMenuItem)

$viewMenu = New-Object System.Windows.Forms.ToolStripMenuItem
$viewMenu.Text = "View"
$menuStrip.Items.Add($viewMenu)

$darkModeMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem
$darkModeMenuItem.Text = "Dark Mode"
$darkModeMenuItem.CheckOnClick = $true
$darkModeMenuItem.Add_CheckedChanged({
    if ($darkModeMenuItem.Checked) {
        $mainForm.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
        $tabControl.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 35)
        $userTab.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
        $groupTab.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
        $contactTab.BackColor = [System.Drawing.Color]::FromArgb(45, 45, 48)
        foreach ($control in $mainForm.Controls) {
            if ($control -is [System.Windows.Forms.Label]) { $control.ForeColor = [System.Drawing.Color]::White }
        }
        foreach ($tab in @($userTab, $groupTab, $contactTab)) {
            foreach ($control in $tab.Controls) {
                if ($control -is [System.Windows.Forms.Label] -and -not $control.Text.Contains("*")) { $control.ForeColor = [System.Drawing.Color]::White }
            }
        }
        $statusBar.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 35)
        $statusLabel.ForeColor = [System.Drawing.Color]::White
        Update-Status "Dark mode enabled"
    } else {
        $mainForm.BackColor = [System.Drawing.SystemColors]::Control
        $tabControl.BackColor = [System.Drawing.SystemColors]::Control
        $userTab.BackColor = [System.Drawing.SystemColors]::Control
        $groupTab.BackColor = [System.Drawing.SystemColors]::Control
        $contactTab.BackColor = [System.Drawing.SystemColors]::Control
        foreach ($control in $mainForm.Controls) {
            if ($control -is [System.Windows.Forms.Label]) { $control.ForeColor = [System.Drawing.SystemColors]::ControlText }
        }
        foreach ($tab in @($userTab, $groupTab, $contactTab)) {
            foreach ($control in $tab.Controls) {
                if ($control -is [System.Windows.Forms.Label] -and -not $control.Text.Contains("*")) { $control.ForeColor = [System.Drawing.SystemColors]::ControlText }
            }
        }
        $statusBar.BackColor = [System.Drawing.SystemColors]::Control
        $statusLabel.ForeColor = [System.Drawing.SystemColors]::ControlText
        Update-Status "Light mode enabled"
    }
})
$viewMenu.DropDownItems.Add($darkModeMenuItem)

$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(10, 27)
$tabControl.Size = New-Object System.Drawing.Size(565, 480)
$tabControl.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$mainForm.Controls.Add($tabControl)

$userTab = New-Object System.Windows.Forms.TabPage
$userTab.Text = "Create User"
$groupTab = New-Object System.Windows.Forms.TabPage
$groupTab.Text = "Manage Distribution Group"
$contactTab = New-Object System.Windows.Forms.TabPage
$contactTab.Text = "Create Contact"
$tabControl.TabPages.Add($userTab)
$tabControl.TabPages.Add($groupTab)
$tabControl.TabPages.Add($contactTab)

$buttonPanel = New-Object System.Windows.Forms.Panel
$buttonPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
$buttonPanel.Height = 40
$buttonPanel.Location = New-Object System.Drawing.Point(0, 510)

$helpButton = New-Object System.Windows.Forms.Button
$helpButton.Text = "Help"
$helpButton.Location = New-Object System.Drawing.Point(20, 5)
$helpButton.Size = New-Object System.Drawing.Size(80, 30)
$helpButton.Add_Click({
    $helpText = @"
AD Management Tool

This application provides a unified interface for common Active Directory tasks:

1. User Creation Tab:
   - Creates new AD users with email addresses
   - Configures proxy addresses for multiple domains
   - Generates secure random passwords

2. Distribution Group Tab:
   - Creates and updates universal distribution groups
   - Configures email attributes

3. Contact Creation Tab:
   - Creates external contacts in AD
   - Configures display name and email address

Logs are stored in: $LogsFolder
"@
    [System.Windows.Forms.MessageBox]::Show($helpText, "AD Management Tool Help", "OK", "Information")
})
$buttonPanel.Controls.Add($helpButton)

$allLogsButton = New-Object System.Windows.Forms.Button
$allLogsButton.Text = "Open Logs Folder"
$allLogsButton.Location = New-Object System.Drawing.Point(440, 5)
$allLogsButton.Size = New-Object System.Drawing.Size(120, 30)
$allLogsButton.Add_Click({
    if (Test-Path $LogsFolder) { Start-Process "explorer.exe" -ArgumentList $LogsFolder }
    else { [System.Windows.Forms.MessageBox]::Show("Logs folder does not exist yet.", "Information", "OK", "Information") }
})
$buttonPanel.Controls.Add($allLogsButton)
$mainForm.Controls.Add($buttonPanel)

# User Creation Tab
$userLabels = @("First Name*:", "Last Name*:", "Primary Domain*:", "Additional Domains (comma separated):", "Description:", "Office:", "Department:", "Job Title:")
$userTextBoxes = @()
$yPos = 20
foreach ($label in $userLabels) {
    $userLabel = New-Object System.Windows.Forms.Label
    $userLabel.Text = $label
    $userLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $userLabel.Size = New-Object System.Drawing.Size(200, 20)
    $userLabel.Font = if ($label.Contains("*")) { New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold) } else { New-Object System.Drawing.Font("Segoe UI", 9) }
    $userTab.Controls.Add($userLabel)
    $userTextBox = New-Object System.Windows.Forms.TextBox
    $userTextBox.Location = New-Object System.Drawing.Point(230, $yPos)
    $userTextBox.Size = New-Object System.Drawing.Size(300, 20)
    $userTab.Controls.Add($userTextBox)
    $userTextBoxes += $userTextBox
    $yPos += 40
}

$previewLabel = New-Object System.Windows.Forms.Label
$previewLabel.Text = "Username Preview:"
$previewLabel.Location = New-Object System.Drawing.Point(20, $yPos)
$previewLabel.Size = New-Object System.Drawing.Size(150, 20)
$userTab.Controls.Add($previewLabel)

$previewUsername = New-Object System.Windows.Forms.Label
$previewUsername.Text = ""
$previewUsername.Location = New-Object System.Drawing.Point(230, $yPos)
$previewUsername.Size = New-Object System.Drawing.Size(300, 20)
$previewUsername.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$userTab.Controls.Add($previewUsername)

$userTextBoxes[0].Add_TextChanged({ if ($userTextBoxes[0].Text -and $userTextBoxes[1].Text) { $previewUsername.Text = Get-UsernamePreview -FirstName $userTextBoxes[0].Text -LastName $userTextBoxes[1].Text } })
$userTextBoxes[1].Add_TextChanged({ if ($userTextBoxes[0].Text -and $userTextBoxes[1].Text) { $previewUsername.Text = Get-UsernamePreview -FirstName $userTextBoxes[0].Text -LastName $userTextBoxes[1].Text } })

$bulkImportButton = New-Object System.Windows.Forms.Button
$bulkImportButton.Text = "Bulk Import"
$bulkImportButton.Location = New-Object System.Drawing.Point(100, 360)
$bulkImportButton.Size = New-Object System.Drawing.Size(120, 30)
$bulkImportButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $openFileDialog.Title = "Select CSV File for Bulk User Import"
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $csvPath = $openFileDialog.FileName
        try {
            $csvData = Import-Csv -Path $csvPath
            $requiredHeaders = @("FirstName", "LastName", "PrimaryDomain")
            $csvHeaders = $csvData | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
            $missingHeaders = $requiredHeaders | Where-Object { $_ -notin $csvHeaders }
            if ($missingHeaders.Count -gt 0) {
                [System.Windows.Forms.MessageBox]::Show("The CSV file is missing the following required headers: $($missingHeaders -join ', ')`n`nRequired headers: $($requiredHeaders -join ', ')", "Invalid CSV Format", "OK", "Error")
                return
            }
            $confirmResult = [System.Windows.Forms.MessageBox]::Show("Ready to import $($csvData.Count) users. Would you like to continue?", "Confirm Bulk Import", "YesNo", "Question")
            if ($confirmResult -eq [System.Windows.Forms.DialogResult]::Yes) {
                $progressForm = New-Object System.Windows.Forms.Form
                $progressForm.Text = "Importing Users"
                $progressForm.Size = New-Object System.Drawing.Size(400, 150)
                $progressForm.StartPosition = "CenterScreen"
                $progressForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
                $progressForm.MaximizeBox = $false
                $progressForm.MinimizeBox = $false
                $progressBar = New-Object System.Windows.Forms.ProgressBar
                $progressBar.Location = New-Object System.Drawing.Point(10, 50)
                $progressBar.Size = New-Object System.Drawing.Size(360, 30)
                $progressBar.Minimum = 0
                $progressBar.Maximum = $csvData.Count
                $progressBar.Step = 1
                $progressForm.Controls.Add($progressBar)
                $statusLabel = New-Object System.Windows.Forms.Label
                $statusLabel.Location = New-Object System.Drawing.Point(10, 20)
                $statusLabel.Size = New-Object System.Drawing.Size(360, 20)
                $statusLabel.Text = "Preparing to import users..."
                $progressForm.Controls.Add($statusLabel)
                $progressForm.Show()
                $progressForm.Refresh()
                $successCount = 0
                $failCount = 0
                $results = @()
                foreach ($user in $csvData) {
                    $progressBar.PerformStep()
                    $statusLabel.Text = "Processing $($user.FirstName) $($user.LastName)..."
                    $progressForm.Refresh()
                    try {
                        $FirstName = $user.FirstName.Trim()
                        $LastName = $user.LastName.Trim()
                        $PrimaryDomain = $user.PrimaryDomain.Trim()
                        $AdditionalDomains = if ($user.AdditionalDomains) { $user.AdditionalDomains -split ";" | ForEach-Object { $_.Trim() } | Where-Object { $_ } } else { @() }
                        $Description = if ($user.Description) { $user.Description.Trim() } else { "" }
                        $Office = if ($user.Office) { $user.Office.Trim() } else { "" }
                        $Department = if ($user.Department) { $user.Department.Trim() } else { "" }
                        $JobTitle = if ($user.JobTitle) { $user.JobTitle.Trim() } else { "" }
                        $FirstNameParam = if ($Config.UserNamingPattern.Contains("{0}")) { $FirstName } else { $FirstName.Substring(0, 1) }
                        $LastNameParam = if ($Config.UserNamingPattern.Contains("{1}")) { $LastName } else { $LastName.Substring(0, 1) }
                        $Username = ($Config.UserNamingPattern -f $FirstNameParam, $LastNameParam).ToLower()
                        $PrimaryEmail = "$Username@$PrimaryDomain"
                        $ProxyAddresses = @("SMTP:$PrimaryEmail")
                        foreach ($domain in $AdditionalDomains) { if ($domain) { $ProxyAddresses += "smtp:$Username@$domain" } }
                        $Password = Generate-RandomPassword | ConvertTo-SecureString -AsPlainText -Force
                        $PlainPassword = Generate-RandomPassword
                        $existingUser = Get-ADUser -Filter {SamAccountName -eq $Username} -ErrorAction SilentlyContinue
                        if ($existingUser) { throw "User with username '$Username' already exists!" }
                        New-ADUser -Name "$FirstName $LastName" -GivenName $FirstName -Surname $LastName -SamAccountName $Username -UserPrincipalName $PrimaryEmail -EmailAddress $PrimaryEmail -Company $Config.CompanyName -Path $(if ($Config.ForceDefaultOU) { $Config.UserOU } else { $null }) -AccountPassword $Password -Enabled $true
                        Set-ADUser -Identity $Username -Add @{ proxyAddresses = $ProxyAddresses; targetAddress = "SMTP:$PrimaryEmail" }
                        if ($Config.EnableEmailFeatures) {
                            $domainParts = $Config.DomainFQDN.Split('.')
                            $addressBookPath = [string]::Format($Config.AddressBookPathUser, $domainParts[0].ToUpper(), $domainParts[0], $domainParts[1])
                            Set-ADUser -Identity $Username -Add @{ showInAddressBook = $addressBookPath; msExchHideFromAddressLists = $false }
                        }
                        $optionalAttribs = @{}
                        if ($Description) { $optionalAttribs["description"] = $Description }
                        if ($Office) { $optionalAttribs["physicalDeliveryOfficeName"] = $Office }
                        if ($Department) { $optionalAttribs["department"] = $Department }
                        if ($JobTitle) { $optionalAttribs["title"] = $JobTitle }
                        if ($optionalAttribs.Count -gt 0) { Set-ADUser -Identity $Username -Replace $optionalAttribs }
                        Write-Log -Message "User '$Username' created successfully via bulk import!" -LogFile $UserLogFile
                        $successCount++
                        $results += [PSCustomObject]@{ FirstName = $FirstName; LastName = $LastName; Username = $Username; Email = $PrimaryEmail; Password = $PlainPassword; Status = "Success" }
                    } catch {
                        $errorMsg = "Failed to create user '$Username': $_"
                        Write-Log -Message $errorMsg -LogFile $UserLogFile
                        $failCount++
                        $results += [PSCustomObject]@{ FirstName = $FirstName; LastName = $LastName; Username = $Username; Email = $PrimaryEmail; Password = ""; Status = "Failed: $_" }
                    }
                }
                $progressForm.Close()
                $resultPath = "$env:USERPROFILE\Documents\BulkImport_Results_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
                $results | Export-Csv -Path $resultPath -NoTypeInformation
                [System.Windows.Forms.MessageBox]::Show("Import completed with $successCount successes and $failCount failures.`n`nResults have been saved to:`n$resultPath", "Import Complete", "OK", "Information")
                $openResults = [System.Windows.Forms.MessageBox]::Show("Would you like to open the results file?", "Open Results", "YesNo", "Question")
                if ($openResults -eq [System.Windows.Forms.DialogResult]::Yes) { Start-Process $resultPath }
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error processing CSV file: $_", "Error", "OK", "Error")
        }
    }
})
$userTab.Controls.Add($bulkImportButton)

$createUserButton = New-Object System.Windows.Forms.Button
$createUserButton.Text = "Create User"
$createUserButton.Location = New-Object System.Drawing.Point(230, 360)
$createUserButton.Size = New-Object System.Drawing.Size(120, 30)
$createUserButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$createUserButton.ForeColor = [System.Drawing.Color]::White
$createUserButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$createUserButton.Add_Click({
    $FirstName = $userTextBoxes[0].Text.Trim()
    $LastName = $userTextBoxes[1].Text.Trim()
    $PrimaryDomain = $userTextBoxes[2].Text.Trim()
    $AdditionalDomains = $userTextBoxes[3].Text.Trim() -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    $Description = $userTextBoxes[4].Text.Trim()
    $Office = $userTextBoxes[5].Text.Trim()
    $Department = $userTextBoxes[6].Text.Trim()
    $JobTitle = $userTextBoxes[7].Text.Trim()
    if (-not $FirstName -or -not $LastName -or -not $PrimaryDomain) {
        $msg = "Please fill in all required fields (First Name, Last Name, Primary Domain)."
        Write-Log -Message $msg -LogFile $UserLogFile
        [System.Windows.Forms.MessageBox]::Show($msg, "Error", "OK", "Error")
        return
    }
    $FirstNameParam = if ($Config.UserNamingPattern.Contains("{0}")) { $FirstName } else { $FirstName.Substring(0, 1) }
    $LastNameParam = if ($Config.UserNamingPattern.Contains("{1}")) { $LastName } else { $LastName.Substring(0, 1) }
    $Username = ($Config.UserNamingPattern -f $FirstNameParam, $LastNameParam).ToLower()
    $PrimaryEmail = "$Username@$PrimaryDomain"
    $ProxyAddresses = @("SMTP:$PrimaryEmail")
    foreach ($domain in $AdditionalDomains) { if ($domain) { $ProxyAddresses += "smtp:$Username@$domain" } }
    try {
        Write-Log -Message "Attempting to create user: $Username ($PrimaryEmail)" -LogFile $UserLogFile
        $Password = Generate-RandomPassword | ConvertTo-SecureString -AsPlainText -Force
        $existingUser = Get-ADUser -Filter {SamAccountName -eq $Username} -ErrorAction SilentlyContinue
        if ($existingUser) {
            $errorMsg = "User with username '$Username' already exists!"
            Write-Log -Message $errorMsg -LogFile $UserLogFile
            [System.Windows.Forms.MessageBox]::Show($errorMsg, "Error", "OK", "Error")
            return
        }
        New-ADUser -Name "$FirstName $LastName" -GivenName $FirstName -Surname $LastName -SamAccountName $Username -UserPrincipalName $PrimaryEmail -EmailAddress $PrimaryEmail -Company $Config.CompanyName -Path $(if ($Config.ForceDefaultOU) { $Config.UserOU } else { $null }) -AccountPassword $Password -Enabled $true
        Set-ADUser -Identity $Username -Add @{ proxyAddresses = $ProxyAddresses; targetAddress = "SMTP:$PrimaryEmail" }
        if ($Config.EnableEmailFeatures) {
            $domainParts = $Config.DomainFQDN.Split('.')
            $addressBookPath = [string]::Format($Config.AddressBookPathUser, $domainParts[0].ToUpper(), $domainParts[0], $domainParts[1])
            Set-ADUser -Identity $Username -Add @{ showInAddressBook = $addressBookPath; msExchHideFromAddressLists = $false }
        }
        $optionalAttribs = @{}
        if ($Description) { $optionalAttribs["description"] = $Description }
        if ($Office) { $optionalAttribs["physicalDeliveryOfficeName"] = $Office }
        if ($Department) { $optionalAttribs["department"] = $Department }
        if ($JobTitle) { $optionalAttribs["title"] = $JobTitle }
        if ($optionalAttribs.Count -gt 0) { Set-ADUser -Identity $Username -Replace $optionalAttribs }
        $plainPassword = Generate-RandomPassword
        $successMsg = "User '$Username' created successfully with password: $plainPassword"
        Write-Log -Message "User '$Username' created successfully!" -LogFile $UserLogFile
        [System.Windows.Forms.MessageBox]::Show($successMsg, "Success", "OK", "Information")
        foreach ($textBox in $userTextBoxes) { $textBox.Text = "" }
        $previewUsername.Text = ""
    } catch {
        $errorMsg = "Failed to create user '$Username': $_"
        Write-Log -Message $errorMsg -LogFile $UserLogFile
        [System.Windows.Forms.MessageBox]::Show($errorMsg, "Error", "OK", "Error")
    }
})
$userTab.Controls.Add($createUserButton)

$viewUserLogButton = New-Object System.Windows.Forms.Button
$viewUserLogButton.Text = "View Log"
$viewUserLogButton.Location = New-Object System.Drawing.Point(370, 360)
$viewUserLogButton.Size = New-Object System.Drawing.Size(120, 30)
$viewUserLogButton.Add_Click({
    if (Test-Path $UserLogFile) { Start-Process "notepad.exe" -ArgumentList $UserLogFile }
    else { [System.Windows.Forms.MessageBox]::Show("Log file does not exist yet.", "Information", "OK", "Information") }
})
$userTab.Controls.Add($viewUserLogButton)

$requiredNote = New-Object System.Windows.Forms.Label
$requiredNote.Text = "* Required fields"
$requiredNote.Location = New-Object System.Drawing.Point(20, 400)
$requiredNote.Size = New-Object System.Drawing.Size(200, 20)
$requiredNote.ForeColor = [System.Drawing.Color]::Red
$userTab.Controls.Add($requiredNote)

# Distribution Group Tab
$groupLabels = @("Group Name*:", "Display Name*:", "Domain*:", "Description:", "OU Path*:")
$groupTextBoxes = @()
$yPos = 20
foreach ($label in $groupLabels) {
    $groupLabel = New-Object System.Windows.Forms.Label
    $groupLabel.Text = $label
    $groupLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $groupLabel.Size = New-Object System.Drawing.Size(200, 20)
    $groupLabel.Font = if ($label.Contains("*")) { New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold) } else { New-Object System.Drawing.Font("Segoe UI", 9) }
    $groupTab.Controls.Add($groupLabel)
    $groupTextBox = New-Object System.Windows.Forms.TextBox
    $groupTextBox.Location = New-Object System.Drawing.Point(230, $yPos)
    $groupTextBox.Size = New-Object System.Drawing.Size(300, 20)
    $groupTab.Controls.Add($groupTextBox)
    $groupTextBoxes += $groupTextBox
    $yPos += 40
}

$groupTextBoxes[4].Text = $Config.GroupOU
$groupTextBoxes[4].ReadOnly = $Config.ForceDefaultOU

$previewEmailLabel = New-Object System.Windows.Forms.Label
$previewEmailLabel.Text = "Email Preview:"
$previewEmailLabel.Location = New-Object System.Drawing.Point(20, $yPos)
$previewEmailLabel.Size = New-Object System.Drawing.Size(150, 20)
$groupTab.Controls.Add($previewEmailLabel)

$previewEmail = New-Object System.Windows.Forms.Label
$previewEmail.Text = ""
$previewEmail.Location = New-Object System.Drawing.Point(230, $yPos)
$previewEmail.Size = New-Object System.Drawing.Size(300, 20)
$previewEmail.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$groupTab.Controls.Add($previewEmail)

$groupTextBoxes[0].Add_TextChanged({ if ($groupTextBoxes[0].Text -and $groupTextBoxes[2].Text) { $previewEmail.Text = "$($groupTextBoxes[0].Text)@$($groupTextBoxes[2].Text)" } })
$groupTextBoxes[2].Add_TextChanged({ if ($groupTextBoxes[0].Text -and $groupTextBoxes[2].Text) { $previewEmail.Text = "$($groupTextBoxes[0].Text)@$($groupTextBoxes[2].Text)" } })

$createGroupButton = New-Object System.Windows.Forms.Button
$createGroupButton.Text = if ($Config.EnableEmailFeatures) { "Create Mail Group" } else { "Create Group" }
$createGroupButton.Location = New-Object System.Drawing.Point(230, 260)
$createGroupButton.Size = New-Object System.Drawing.Size(150, 30)
$createGroupButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$createGroupButton.ForeColor = [System.Drawing.Color]::White
$createGroupButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$createGroupButton.Add_Click({
    $GroupName = $groupTextBoxes[0].Text.Trim()
    $DisplayName = $groupTextBoxes[1].Text.Trim()
    $Domain = $groupTextBoxes[2].Text.Trim()
    $Description = $groupTextBoxes[3].Text.Trim()
    $OUPath = $groupTextBoxes[4].Text.Trim()
    if (-not $GroupName -or -not $DisplayName -or -not $Domain -or -not $OUPath) {
        $msg = "Please fill in all required fields."
        Write-Log -Message $msg -LogFile $GroupLogFile
        [System.Windows.Forms.MessageBox]::Show($msg, "Error", "OK", "Error")
        return
    }
    $Email = "$GroupName@$Domain"
    $ProxyAddresses = @("SMTP:$Email")
    $TargetAddress = "SMTP:$Email"
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
    } catch {
        $msg = "Failed to import ActiveDirectory module: $_"
        Write-Log -Message $msg -LogFile $GroupLogFile
        [System.Windows.Forms.MessageBox]::Show($msg, "Error", "OK", "Error")
        return
    }
    $existingGroup = Get-ADGroup -Filter {Name -eq $GroupName} -ErrorAction SilentlyContinue
    if ($existingGroup) {
        $msg = "Group '$GroupName' already exists. Updating attributes..."
        Write-Log -Message $msg -LogFile $GroupLogFile
        try {
            Set-ADGroup -Identity $GroupName -DisplayName $DisplayName -Description "Universal Distribution Group for $GroupName"
            Set-ADGroup -Identity $GroupName -Replace @{ mail = $Email; proxyAddresses = $ProxyAddresses; targetAddress = $TargetAddress; mailNickname = $GroupName }
            if ($Config.EnableEmailFeatures) {
                $domainParts = $Config.DomainFQDN.Split('.')
                $addressBookPath = [string]::Format($Config.AddressBookPathGroup, $domainParts[0].ToUpper(), $domainParts[0], $domainParts[1])
                Set-ADGroup -Identity $GroupName -Add @{ showInAddressBook = $addressBookPath; msExchHideFromAddressLists = $false }
            }
            $successMsg = "Group '$GroupName' updated successfully!"
            Write-Log -Message $successMsg -LogFile $GroupLogFile
            [System.Windows.Forms.MessageBox]::Show($successMsg, "Success", "OK", "Information")
        } catch {
            $errorMsg = "Error updating group: $_"
            Write-Log -Message $errorMsg -LogFile $GroupLogFile
            [System.Windows.Forms.MessageBox]::Show($errorMsg, "Error", "OK", "Error")
        }
    } else {
        $msg = "Creating new group '$GroupName' in '$OUPath'..."
        Write-Log -Message $msg -LogFile $GroupLogFile
        try {
            New-ADGroup -Name $GroupName -SamAccountName $GroupName -DisplayName $DisplayName -GroupCategory Distribution -GroupScope Universal -Path $OUPath -Description "Universal Distribution Group for $GroupName"
            Set-ADGroup -Identity $GroupName -Replace @{ mail = $Email; proxyAddresses = $ProxyAddresses; targetAddress = $TargetAddress; mailNickname = $GroupName }
            if ($Config.EnableEmailFeatures) {
                $domainParts = $Config.DomainFQDN.Split('.')
                $addressBookPath = [string]::Format($Config.AddressBookPathGroup, $domainParts[0].ToUpper(), $domainParts[0], $domainParts[1])
                Set-ADGroup -Identity $GroupName -Add @{ showInAddressBook = $addressBookPath; msExchHideFromAddressLists = $false }
            }
            $successMsg = "Distribution group '$GroupName' created successfully in '$OUPath'."
            Write-Log -Message $successMsg -LogFile $GroupLogFile
            [System.Windows.Forms.MessageBox]::Show($successMsg, "Success", "OK", "Information")
            foreach ($textBox in $groupTextBoxes) { if ($textBox -ne $groupTextBoxes[4]) { $textBox.Text = "" } }
            $previewEmail.Text = ""
        } catch {
            $errorMsg = "Error creating group: $_"
            Write-Log -Message $errorMsg -LogFile $GroupLogFile
            [System.Windows.Forms.MessageBox]::Show($errorMsg, "Error", "OK", "Error")
        }
    }
})
$groupTab.Controls.Add($createGroupButton)

$viewGroupLogButton = New-Object System.Windows.Forms.Button
$viewGroupLogButton.Text = "View Log"
$viewGroupLogButton.Location = New-Object System.Drawing.Point(400, 260)
$viewGroupLogButton.Size = New-Object System.Drawing.Size(120, 30)
$viewGroupLogButton.Add_Click({
    if (Test-Path $GroupLogFile) { Start-Process "notepad.exe" -ArgumentList $GroupLogFile }
    else { [System.Windows.Forms.MessageBox]::Show("Log file does not exist yet.", "Information", "OK", "Information") }
})
$groupTab.Controls.Add($viewGroupLogButton)

$groupRequiredNote = New-Object System.Windows.Forms.Label
$groupRequiredNote.Text = "* Required fields"
$groupRequiredNote.Location = New-Object System.Drawing.Point(20, 300)
$groupRequiredNote.Size = New-Object System.Drawing.Size(200, 20)
$groupRequiredNote.ForeColor = [System.Drawing.Color]::Red
$groupTab.Controls.Add($groupRequiredNote)

# Contact Creation Tab
$contactLabels = @("First Name*:", "Last Name*:", "Display Name*:", "Email Address*:", "Description:", "Company:")
$contactTextBoxes = @()
$yPos = 20
foreach ($label in $contactLabels) {
    $contactLabel = New-Object System.Windows.Forms.Label
    $contactLabel.Text = $label
    $contactLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $contactLabel.Size = New-Object System.Drawing.Size(200, 20)
    $contactLabel.Font = if ($label.Contains("*")) { New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold) } else { New-Object System.Drawing.Font("Segoe UI", 9) }
    $contactTab.Controls.Add($contactLabel)
    $contactTextBox = New-Object System.Windows.Forms.TextBox
    $contactTextBox.Location = New-Object System.Drawing.Point(230, $yPos)
    $contactTextBox.Size = New-Object System.Drawing.Size(300, 20)
    $contactTab.Controls.Add($contactTextBox)
    $contactTextBoxes += $contactTextBox
    $yPos += 40
}

$contactOULabel = New-Object System.Windows.Forms.Label
$contactOULabel.Text = "OU Path*:"
$contactOULabel.Location = New-Object System.Drawing.Point(20, $yPos)
$contactOULabel.Size = New-Object System.Drawing.Size(200, 20)
$contactOULabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$contactTab.Controls.Add($contactOULabel)

$contactOU = New-Object System.Windows.Forms.TextBox
$contactOU.Text = $Config.ContactOU
$contactOU.Location = New-Object System.Drawing.Point(230, $yPos)
$contactOU.Size = New-Object System.Drawing.Size(300, 20)
$contactOU.ReadOnly = $Config.ForceDefaultOU
$contactTab.Controls.Add($contactOU)

$contactTextBoxes[0].Add_TextChanged({ if ($contactTextBoxes[0].Text -and $contactTextBoxes[1].Text) { $contactTextBoxes[2].Text = "$($contactTextBoxes[0].Text) $($contactTextBoxes[1].Text)" } })
$contactTextBoxes[1].Add_TextChanged({ if ($contactTextBoxes[0].Text -and $contactTextBoxes[1].Text) { $contactTextBoxes[2].Text = "$($contactTextBoxes[0].Text) $($contactTextBoxes[1].Text)" } })

$createContactButton = New-Object System.Windows.Forms.Button
$createContactButton.Text = "Create Contact"
$createContactButton.Location = New-Object System.Drawing.Point(230, 300)
$createContactButton.Size = New-Object System.Drawing.Size(120, 30)
$createContactButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$createContactButton.ForeColor = [System.Drawing.Color]::White
$createContactButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$createContactButton.Add_Click({
    $FirstName = $contactTextBoxes[0].Text.Trim()
    $LastName = $contactTextBoxes[1].Text.Trim()
    $DisplayName = $contactTextBoxes[2].Text.Trim()
    $Email = $contactTextBoxes[3].Text.Trim()
    $Description = $contactTextBoxes[4].Text.Trim()
    $Company = $contactTextBoxes[5].Text.Trim()
    $OUPath = $contactOU.Text.Trim()
    $FullName = "$FirstName $LastName"
    if (-not $FirstName -or -not $LastName -or -not $DisplayName -or -not $Email -or -not $OUPath) {
        $msg = "Please fill in all required fields."
        Write-Log -Message $msg -LogFile $ContactLogFile
        [System.Windows.Forms.MessageBox]::Show($msg, "Error", "OK", "Error")
        return
    }
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        Write-Log -Message "Attempting to create contact '$FullName' in '$OUPath' with email '$Email'" -LogFile $ContactLogFile
        $filter = "mail -eq '$Email'"
        $existingContact = Get-ADObject -Filter $filter -SearchBase $OUPath -ErrorAction SilentlyContinue
        if ($existingContact) {
            $msg = "A contact with email '$Email' already exists!"
            Write-Log -Message $msg -LogFile $ContactLogFile
            [System.Windows.Forms.MessageBox]::Show($msg, "Error", "OK", "Error")
            return
        }
        $contactAttributes = @{ displayName = $DisplayName; givenName = $FirstName; sn = $LastName; mail = $Email }
        if ($Description.ConcurrentModificationException) { $contactAttributes["description"] = $Description }
        if ($Company) { $contactAttributes["company"] = $Company }
        New-ADObject -Type Contact -Name $FullName -Path $OUPath -OtherAttributes $contactAttributes
        $msg = "Contact '$FullName' created successfully!"
        Write-Log -Message $msg -LogFile $ContactLogFile
        [System.Windows.Forms.MessageBox]::Show($msg, "Success", "OK", "Information")
        foreach ($textBox in $contactTextBoxes) { $textBox.Text = "" }
    } catch {
        $errorMsg = "Failed to create contact: $_"
        Write-Log -Message $errorMsg -LogFile $ContactLogFile
        [System.Windows.Forms.MessageBox]::Show($errorMsg, "Error", "OK", "Error")
    }
})
$contactTab.Controls.Add($createContactButton)

$viewContactLogButton = New-Object System.Windows.Forms.Button
$viewContactLogButton.Text = "View Log"
$viewContactLogButton.Location = New-Object System.Drawing.Point(370, 300)
$viewContactLogButton.Size = New-Object System.Drawing.Size(120, 30)
$viewContactLogButton.Add_Click({
    if (Test-Path $ContactLogFile) { Start-Process "notepad.exe" -ArgumentList $ContactLogFile }
    else { [System.Windows.Forms.MessageBox]::Show("Log file does not exist yet.", "Information", "OK", "Information") }
})
$contactTab.Controls.Add($viewContactLogButton)

$contactRequiredNote = New-Object System.Windows.Forms.Label
$contactRequiredNote.Text = "* Required fields"
$contactRequiredNote.Location = New-Object System.Drawing.Point(20, 340)
$contactRequiredNote.Size = New-Object System.Drawing.Size(200, 20)
$contactRequiredNote.ForeColor = [System.Drawing.Color]::Red
$contactTab.Controls.Add($contactRequiredNote)

# Configuration Dialog
function Show-ConfigurationDialog {
    $configForm = New-Object System.Windows.Forms.Form
    $configForm.Text = "AD Management Tool Configuration"
    $configForm.Size = New-Object System.Drawing.Size(600, 500)
    $configForm.StartPosition = "CenterScreen"
    $configForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $configForm.MaximizeBox = $false
    $configForm.MinimizeBox = $false
    $configForm.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $configTabControl = New-Object System.Windows.Forms.TabControl
    $configTabControl.Location = New-Object System.Drawing.Point(10, 10)
    $configTabControl.Size = New-Object System.Drawing.Size(565, 400)
    $configForm.Controls.Add($configTabControl)

    $orgTab = New-Object System.Windows.Forms.TabPage
    $orgTab.Text = "Organization"
    $configTabControl.TabPages.Add($orgTab)

    $yPos = 20
    $configFields = @(@{ Label = "Company Name:"; ConfigKey = "CompanyName" }, @{ Label = "Domain NetBIOS Name:"; ConfigKey = "DomainNetBIOS" }, @{ Label = "Domain FQDN:"; ConfigKey = "DomainFQDN" }, @{ Label = "Default Primary Email Domain:"; ConfigKey = "DefaultPrimaryDomain" })
    $configTextBoxes = @{}
    foreach ($field in $configFields) {
        $label = New-Object System.Windows.Forms.Label
        $label.Text = $field.Label
        $label.Location = New-Object System.Drawing.Point(20, $yPos)
        $label.Size = New-Object System.Drawing.Size(200, 20)
        $orgTab.Controls.Add($label)
        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Location = New-Object System.Drawing.Point(230, $yPos)
        $textBox.Size = New-Object System.Drawing.Size(300, 20)
        $textBox.Text = $Config[$field.ConfigKey]
        $orgTab.Controls.Add($textBox)
        $configTextBoxes[$field.ConfigKey] = $textBox
        $yPos += 40
    }

    $adTab = New-Object System.Windows.Forms.TabPage
    $adTab.Text = "Active Directory"
    $configTabControl.TabPages.Add($adTab)

    $yPos = 20
    $adFields = @(@{ Label = "Users OU:"; ConfigKey = "UserOU" }, @{ Label = "Groups OU:"; ConfigKey = "GroupOU" }, @{ Label = "Contacts OU:"; ConfigKey = "ContactOU" })
    foreach ($field in $adFields) {
        $label = New-Object System.Windows.Forms.Label
        $label.Text = $field.Label
        $label.Location = New-Object System.Drawing.Point(20, $yPos)
        $label.Size = New-Object System.Drawing.Size(200, 20)
        $adTab.Controls.Add($label)
        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Location = New-Object System.Drawing.Point(230, $yPos)
        $textBox.Size = New-Object System.Drawing.Size(300, 20)
        $textBox.Text = $Config[$field.ConfigKey]
        $adTab.Controls.Add($textBox)
        $configTextBoxes[$field.ConfigKey] = $textBox
        $yPos += 40
    }

    $forceOUCheckbox = New-Object System.Windows.Forms.CheckBox
    $forceOUCheckbox.Text = "Force default OU paths (don't allow changes in UI)"
    $forceOUCheckbox.Location = New-Object System.Drawing.Point(20, $yPos)
    $forceOUCheckbox.Size = New-Object System.Drawing.Size(350, 20)
    $forceOUCheckbox.Checked = $Config.ForceDefaultOU
    $adTab.Controls.Add($forceOUCheckbox)

    $emailTab = New-Object System.Windows.Forms.TabPage
    $emailTab.Text = "Email Configuration"
    $configTabControl.TabPages.Add($emailTab)

    $yPos = 20
    $enableEmailCheckbox = New-Object System.Windows.Forms.CheckBox
    $enableEmailCheckbox.Text = "Enable Email Features (Configure Exchange Attributes)"
    $enableEmailCheckbox.Location = New-Object System.Drawing.Point(20, $yPos)
    $enableEmailCheckbox.Size = New-Object System.Drawing.Size(350, 20)
    $enableEmailCheckbox.Checked = $Config.EnableEmailFeatures
    $emailTab.Controls.Add($enableEmailCheckbox)
    $yPos += 40

    $exchangeLabel = New-Object System.Windows.Forms.Label
    $exchangeLabel.Text = "Exchange Server:"
    $exchangeLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $exchangeLabel.Size = New-Object System.Drawing.Size(200, 20)
    $emailTab.Controls.Add($exchangeLabel)
    $exchangeTextBox = New-Object System.Windows.Forms.TextBox
    $exchangeTextBox.Location = New-Object System.Drawing.Point(230, $yPos)
    $exchangeTextBox.Size = New-Object System.Drawing.Size(300, 20)
    $exchangeTextBox.Text = $Config.ExchangeServer
    $emailTab.Controls.Add($exchangeTextBox)
    $configTextBoxes["ExchangeServer"] = $exchangeTextBox
    $yPos += 40

    $versionLabel = New-Object System.Windows.Forms.Label
    $versionLabel.Text = "Exchange Version:"
    $versionLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $versionLabel.Size = New-Object System.Drawing.Size(200, 20)
    $emailTab.Controls.Add($versionLabel)
    $versionDropDown = New-Object System.Windows.Forms.ComboBox
    $versionDropDown.Location = New-Object System.Drawing.Point(230, $yPos)
    $versionDropDown.Size = New-Object System.Drawing.Size(300, 20)
    $versionDropDown.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $versionOptions = @("Exchange 2013", "Exchange 2016", "Exchange 2019", "Exchange Online")
    foreach ($option in $versionOptions) { $versionDropDown.Items.Add($option) | Out-Null }
    $versionDropDown.SelectedItem = $Config.ExchangeVersion
    $emailTab.Controls.Add($versionDropDown)
    $yPos += 40

    $patternLabel = New-Object System.Windows.Forms.Label
    $patternLabel.Text = "Username Pattern:"
    $patternLabel.Location = New-Object System.Drawing.Point(20, $yPos)
    $patternLabel.Size = New-Object System.Drawing.Size(200, 20)
    $emailTab.Controls.Add($patternLabel)
    $patternDropDown = New-Object System.Windows.Forms.ComboBox
    $patternDropDown.Location = New-Object System.Drawing.Point(230, $yPos)
    $patternDropDown.Size = New-Object System.Drawing.Size(300, 20)
    $patternDropDown.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $patternOptions = @("FirstName.LastName", "First.Last", "FLast", "FirstL", "LastF", "LastName.FirstName")
    foreach ($option in $patternOptions) { $patternDropDown.Items.Add($option) | Out-Null }
    $currentPattern = switch ($Config.UserNamingPattern) {
        "{0}.{1}" { "FirstName.LastName" }
        "{0}{1}" { "FLast" }
        "{1}{0}" { "LastF" }
        "{1}.{0}" { "LastName.FirstName" }
        default { "FirstName.LastName" }
    }
    $patternDropDown.SelectedItem = $currentPattern
    $emailTab.Controls.Add($patternDropDown)

    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Text = "Save Configuration"
    $saveButton.Location = New-Object System.Drawing.Point(230, 420)
    $saveButton.Size = New-Object System.Drawing.Size(150, 30)
    $saveButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    $saveButton.ForeColor = [System.Drawing.Color]::White
    $saveButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $saveButton.Add_Click({
        foreach ($key in $configTextBoxes.Keys) { $Config[$key] = $configTextBoxes[$key].Text }
        $Config.ForceDefaultOU = $forceOUCheckbox.Checked
        $Config.EnableEmailFeatures = $enableEmailCheckbox.Checked
        $Config.ExchangeVersion = $versionDropDown.SelectedItem
        $Config.UserNamingPattern = switch ($patternDropDown.SelectedItem) {
            "FirstName.LastName" { "{0}.{1}" }
            "First.Last" { "{0}.{1}" }
            "FLast" { "{0}{1}" }
            "FirstL" { "{0}{1}" }
            "LastF" { "{1}{0}" }
            "LastName.FirstName" { "{1}.{0}" }
            default { "{0}.{1}" }
        }
        if (Save-Configuration) {
            [System.Windows.Forms.MessageBox]::Show("Configuration saved successfully!", "Configuration Saved", "OK", "Information")
            $configForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $configForm.Close()
        }
    })
    $configForm.Controls.Add($saveButton)

    $resetButton = New-Object System.Windows.Forms.Button
    $resetButton.Text = "Reset to Defaults"
    $resetButton.Location = New-Object System.Drawing.Point(50, 420)
    $resetButton.Size = New-Object System.Drawing.Size(150, 30)
    $resetButton.Add_Click({
        $confirmResult = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to reset all settings to default values?", "Confirm Reset", "YesNo", "Warning")
        if ($confirmResult -eq [System.Windows.Forms.DialogResult]::Yes) {
            $script:Config = @{
                CompanyName = "YourCompany"
                DomainNetBIOS = $(if ($env:USERDOMAIN) { $env:USERDOMAIN } else { "DOMAIN" })
                DomainFQDN = $(if ($env:USERDNSDOMAIN) { $env:USERDNSDOMAIN } else { "domain.com" })
                UserOU = "OU=Users,DC=$(($env:USERDNSDOMAIN -split '\.')[0]),DC=$(($env:USERDNSDOMAIN -split '\.')[1])"
                GroupOU = "OU=Groups,DC=$(($env:USERDNSDOMAIN -split '\.')[0]),DC=$(($env:USERDNSDOMAIN -split '\.')[1])"
                ContactOU = "OU=Contacts,DC=$(($env:USERDNSDOMAIN -split '\.')[0]),DC=$(($env:USERDNSDOMAIN -split '\.')[1])"
                ExchangeServer = "exchange.$env:USERDNSDOMAIN"
                ExchangeVersion = "Exchange 2019"
                UserNamingPattern = "{0}.{1}"
                ForceDefaultOU = $false
                AddressBookPathUser = "CN=All Users,CN=All Address Lists,CN=Address Lists Container,CN={0},CN=Microsoft Exchange,CN=Services,CN=Configuration,DC={1},DC={2}"
                AddressBookPathGroup = "CN=All Groups,CN=All Address Lists,CN=Address Lists Container,CN={0},CN=Microsoft Exchange,CN=Services,CN=Configuration,DC={1},DC={2}"
                EnableEmailFeatures = $true
                DefaultPrimaryDomain = $(if ($env:USERDNSDOMAIN) { $env:USERDNSDOMAIN } else { "domain.com" })
            }
            foreach ($key in $configTextBoxes.Keys) { $configTextBoxes[$key].Text = $Config[$key] }
            $forceOUCheckbox.Checked = $Config.ForceDefaultOU
            $enableEmailCheckbox.Checked = $Config.EnableEmailFeatures
            $versionDropDown.SelectedItem = $Config.ExchangeVersion
            $patternDropDown.SelectedItem = "FirstName.LastName"
            [System.Windows.Forms.MessageBox]::Show("Settings have been reset to defaults. Click Save to apply these changes.", "Reset Complete", "OK", "Information")
        }
    })
    $configForm.Controls.Add($resetButton)

    $configForm.ShowDialog()
}

# Startup Checks
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    $warningMessage = "WARNING: This script is not running with administrative privileges. Some Active Directory operations may fail."
    Write-Warning $warningMessage
    $result = [System.Windows.Forms.MessageBox]::Show("$warningMessage`n`nDo you want to restart the script with administrator privileges?", "Administrator Rights Required", "YesNo", "Warning")
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Path)`"" -Verb RunAs
        exit
    }
}

$requiredModules = @("ActiveDirectory")
$missingModules = @()
foreach ($module in $requiredModules) { if (-not (Get-Module -ListAvailable -Name $module)) { $missingModules += $module } }
if ($missingModules.Count -gt 0) {
    $moduleWarning = "The following required modules are missing: $($missingModules -join ', ')"
    Write-Warning $moduleWarning
    [System.Windows.Forms.MessageBox]::Show("$moduleWarning`n`nPlease install the missing modules before using this tool.", "Missing Required Modules", "OK", "Warning")
}

# Final Setup
$mainForm.Add_Shown({
    $userTextBoxes[2].Text = $Config.DefaultPrimaryDomain
    $groupTextBoxes[2].Text = $Config.DefaultPrimaryDomain
    if ($Config.ForceDefaultOU) { $groupTextBoxes[4].ReadOnly = $true; $contactOU.ReadOnly = $true }
    $createGroupButton.Text = if ($Config.EnableEmailFeatures) { "Create Mail Group" } else { "Create Group" }
    Update-Status "Configuration loaded from $ConfigFile"
})

foreach ($textBox in $userTextBoxes + $groupTextBoxes + $contactTextBoxes) { $textBox.AllowDrop = $true }
$contactOU.AllowDrop = $true

$createUserButton.Add_Click({ Update-Status "Creating new user..." })
$createGroupButton.Add_Click({ Update-Status "Creating distribution group..." })
$createContactButton.Add_Click({ Update-Status "Creating new contact..." })

$mainForm.ShowDialog()