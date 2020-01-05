### New User Setup test build
### 9/20/19
### Ian Eden
### Version 1.0

#Launch Script As Adminsitrator
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) { Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs; exit }

#Function to connect to O365
function Connect-O365
{
    $ExecutionPolicy = Get-ExecutionPolicy

    if ($ExecutionPolicy -eq "RemoteSigned" -or "Unrestricted") {
        Write-Host ""
        } Else {
        Set-ExecutionPolicy RemoteSigned
        }
    Write-Host "Please sign in with a Global Administrator account" -ForegroundColor Yellow
    $UserCredential = Get-Credential
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
    try {
        Import-PSSession $Session -DisableNameChecking -ErrorAction Stop
    }
    catch {
        Write-Host "
        
        Error connecting to O365...Please check username and password..."
        Start-Sleep -Seconds 1
        Clear-Host
        Return Connect-O365
    } 
    
    #Install AzureAD Module
    if (Get-Module -ListAvailable -Name "AzureAD") { 
        Write-Host "AzureAD Module Already installed"
    } 
    Else {
        Write-Host "Installing Module 'AzureAD'..."
        Install-Module AzureAD
    }
    
        if (Get-Module -ListAvailable -Name "msonline") { 
        Write-Host "MSonline Module Already installed"
    } 
    Else {
        Write-Host "Installing Module 'MSonline'..."
        Install-Module msonline
    }
            #Connect to Azure-AD and MSOL
    Connect-AzureAD -Credential $UserCredential
    Connect-MsolService -Credential $UserCredential

}


#Function for the warning
function Prompt-warning 
{
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        $warning = New-Object System.Windows.Forms.Form
        $warning.Text = "Advisory"
        $warning.Size = New-Object System.Drawing.Size(400,200)
        $warning.StartPosition = 'CenterScreen'

        $OKButton = New-Object System.Windows.Forms.Button
        $OKButton.Location = New-Object System.Drawing.Point(150,100)
        $OKButton.Size = New-Object System.Drawing.Size(75,23)
        $OKButton.Text = 'Close'
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $warning.AcceptButton = $OKButton
        $warning.Controls.Add($OKButton)
        
        $label = New-Object System.Windows.Forms.Label
        $label.Location = New-Object System.Drawing.Point(60,20)
        $label.Size = New-Object System.Drawing.Size(280,70)

        $label.BorderStyle = "Fixed3D"
        $label.Text = "Please make sure the user's computer is: `n     Online `n     Domain Joined `n     Proper naming convention `n     Required Office Licenses are available "
        $label.ForeColor = "DarkRed"
        $warning.Controls.Add($label)

        $warning.Topmost = $true

        $result = $warning.ShowDialog()

        if ($result -eq [System.Windows.Forms.DialogResult]::OK)
            {   
            Write-Host ""
            }
}


#Function to prompt the userform
function Prompt-userform
{
         
    $BusinessPremium = "Pyroforces:ENTERPRISEPACK"
    #$MSProjectPro = "resller-account:PROJECT"

    #Get Business Premium available Licenses
    $BPTotalLicense = Get-MsolAccountSku | Where-Object AccountSkuID -eq $BusinessPremium | Select-Object ActiveUnits | Out-String
    $BPConsumedLicense = Get-MsolAccountSku | Where-Object AccountSkuID -eq "$BusinessPremium" | Select-Object ConsumedUnits | Out-String
    [int]$BPTrimTotalLicense = $BPTotalLicense.Trim("ActiveUnits
    ----------- ")
    [int]$BPTrimConsumedLicense = $BPConsumedLicense.Trim("ConsumedUnits
    -------------
            ")
    #Subtracting Total Licenses from Consumed
    $script:BPAvailableLicense = $BPTrimTotalLicense - $BPTrimConsumedLicense

    <#
    #Get available MS Project licenses.
    $ProjTotalLicense = Get-MsolAccountSku | Where-Object AccountSkuID -eq "$BusinessPremium" | Select-Object ActiveUnits | Out-String
    $BPConsumedLicense = Get-MsolAccountSku | Where-Object AccountSkuID -eq "$BusinessPremium" | Select-Object ConsumedUnits | Out-String
    [int]$ProjTrimTotalLicense = $ProjTotalLicense.Trim("ActiveUnits
    ----------- ")
    [int]$ProjTrimConsumedLicense = $ProjConsumedLicense.Trim("ConsumedUnits
    -------------
            ")
    #Subtracting Total Licenses from Consumed
    $script:ProjAvailableLicense = $ProjTrimTotalLicense - $ProjTrimConsumedLicense
    #>


    #Start of GUI Box

    Add-Type -AssemblyName System.Drawing
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $Form                                   = New-Object system.Windows.Forms.Form
    $Form.ClientSize                  = '891,712'
    $Form.text                            = "Form"
    $Form.TopMost                    = $true


    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(720,630)
    $OKButton.Size = New-Object System.Drawing.Size(75,23)
    $OKButton.Text = 'Submit'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $OKButton
    $form.Controls.Add($OKButton)

    $Cancel = New-Object System.Windows.Forms.Button
    $Cancel.Location = New-Object System.Drawing.Point(616,630)
    $Cancel.Size = New-Object System.Drawing.Size(75,23)
    $Cancel.Text = 'Cancel'
    $Cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)

    $FirstNameLabel                  = New-Object system.Windows.Forms.Label
    $FirstNameLabel.text           = "First Name"
    $FirstNameLabel.AutoSize   = $true
    $FirstNameLabel.width        = 25
    $FirstNameLabel.height       = 10
    $FirstNameLabel.location    = New-Object System.Drawing.Point(50,40)
    $FirstNameLabel.Font          = 'Microsoft Sans Serif,10'

    $LastNameLabel                   = New-Object system.Windows.Forms.Label
    $LastNameLabel.text            = "Last Name"
    $LastNameLabel.AutoSize    = $true
    $LastNameLabel.width         = 25
    $LastNameLabel.height        = 10
    $LastNameLabel.location     = New-Object System.Drawing.Point(200,40)
    $LastNameLabel.Font           = 'Microsoft Sans Serif,10'

    $EmailLabel                          = New-Object system.Windows.Forms.Label
    $EmailLabel.text                   = "Email"
    $EmailLabel.AutoSize           = $true
    $EmailLabel.width                = 25
    $EmailLabel.height               = 10
    $EmailLabel.location            = New-Object System.Drawing.Point(50,90)
    $EmailLabel.Font                  = 'Microsoft Sans Serif,10'

    $PhoneNumberLabel                   = New-Object system.Windows.Forms.Label
    $PhoneNumberLabel.text            = "Phone Number"
    $PhoneNumberLabel.AutoSize    = $true
    $PhoneNumberLabel.width         = 25
    $PhoneNumberLabel.height        = 10
    $PhoneNumberLabel.location     = New-Object System.Drawing.Point(200,90)
    $PhoneNumberLabel.Font           = 'Microsoft Sans Serif,10'

    $PasswordLabel                   = New-Object system.Windows.Forms.Label
    $PasswordLabel.text            = "Password"
    $PasswordLabel.AutoSize    = $true
    $PasswordLabel.width         = 25
    $PasswordLabel.height        = 10
    $PasswordLabel.location     = New-Object System.Drawing.Point(325,90)
    $PasswordLabel.Font           = 'Microsoft Sans Serif,10'


    $JobTitleLabel                        = New-Object system.Windows.Forms.Label
    $JobTitleLabel.text                 = "Job Title"
    $JobTitleLabel.AutoSize         = $true
    $JobTitleLabel.width              = 25
    $JobTitleLabel.height             = 10
    $JobTitleLabel.location          = New-Object System.Drawing.Point(50,140)
    $JobTitleLabel.Font                = 'Microsoft Sans Serif,10'

    $ComputerNameLabel                   = New-Object system.Windows.Forms.Label
    $ComputerNameLabel.text            = "Computer Name"
    $ComputerNameLabel.AutoSize    = $true
    $ComputerNameLabel.width         = 25
    $ComputerNameLabel.height        = 10
    $ComputerNameLabel.location     = New-Object System.Drawing.Point(200,140)
    $ComputerNameLabel.Font           = 'Microsoft Sans Serif,10'

    $LocationLabel                     = New-Object system.Windows.Forms.Label
    $LocationLabel.text              = "Location"
    $LocationLabel.AutoSize      = $true
    $LocationLabel.width           = 25
    $LocationLabel.height          = 10
    $LocationLabel.location       = New-Object System.Drawing.Point(320,140)
    $LocationLabel.Font             = 'Microsoft Sans Serif,10'

    $ShareDrivesLabel                   = New-Object system.Windows.Forms.Label
    $ShareDrivesLabel.text            = "Share Drives"
    $ShareDrivesLabel.AutoSize    = $true
    $ShareDrivesLabel.width         = 25
    $ShareDrivesLabel.height        = 10
    $ShareDrivesLabel.location     = New-Object System.Drawing.Point(100,250)
    $ShareDrivesLabel.Font           = 'Microsoft Sans Serif,10'

    $DistroListLabel                   = New-Object system.Windows.Forms.Label
    $DistroListLabel.text            = "Distrobution Lists"
    $DistroListLabel.AutoSize    = $true
    $DistroListLabel.width         = 25
    $DistroListLabel.height        = 10
    $DistroListLabel.location     = New-Object System.Drawing.Point(600,250)
    $DistroListLabel.Font            = 'Microsoft Sans Serif,10'

    $ApplicationsInstall                   = New-Object system.Windows.Forms.Label
    $ApplicationsInstall.text            = "Applications"
    $ApplicationsInstall.AutoSize    = $true
    $ApplicationsInstall.width         = 25
    $ApplicationsInstall.height        = 10
    $ApplicationsInstall.location     = New-Object System.Drawing.Point(100,450)
    $ApplicationsInstall.Font           = 'Microsoft Sans Serif,10'

    $Calendars                        = New-Object system.Windows.Forms.Label
    $Calendars.text                 = "Calendars"
    $Calendars.AutoSize         = $true
    $Calendars.width              = 25
    $Calendars.height             = 10
    $Calendars.location          = New-Object System.Drawing.Point(600,450)
    $Calendars.Font                = 'Microsoft Sans Serif,10'

    $LocationComboBox                    = New-Object system.Windows.Forms.ComboBox
    $LocationComboBox.text             = "Austin"
    $LocationComboBox.width          = 100
    $LocationComboBox.height         = 10
    $LocationComboBox.location      = New-Object System.Drawing.Point(325,160)
    $LocationComboBox.Font            = 'Microsoft Sans Serif,10'
        
            [void] $LocationComboBox.Items.Add('Austin')
            [void] $LocationComboBox.Items.Add('Houston')
            [void] $LocationComboBox.Items.Add('San Antonio')
    

    $Austin                            = New-Object system.Windows.Forms.CheckBox
    $Austin.text                     = "O: Austin"
    $Austin.AutoSize             = $true
    $Austin.width                  = 95
    $Austin.height                 = 20
    $Austin.location              = New-Object System.Drawing.Point(50,300)
    $Austin.Font                    = 'Microsoft Sans Serif,10'

    $Houston                         = New-Object system.Windows.Forms.CheckBox
    $Houston.text                  = "H: Houston"
    $Houston.AutoSize          = $true
    $Houston.width               = 95
    $Houston.height              = 20
    $Houston.location           = New-Object System.Drawing.Point(50,325)
    $Houston.Font                 = 'Microsoft Sans Serif,10'

    $User                            = New-Object system.Windows.Forms.CheckBox
    $User.text                     = "U: User"
    $User.AutoSize             = $true
    $User.width                  = 95
    $User.height                 = 20
    $User.location              = New-Object System.Drawing.Point(50,350)
    $User.Font                    = 'Microsoft Sans Serif,10'

    $Scanner                       = New-Object system.Windows.Forms.CheckBox
    $Scanner.text                = "S: Scanner"
    $Scanner.AutoSize        = $true
    $Scanner.width             = 95
    $Scanner.height            = 20
    $Scanner.location         = New-Object System.Drawing.Point(200,300)
    $Scanner.Font               = 'Microsoft Sans Serif,10'

    $Marketing                         = New-Object system.Windows.Forms.CheckBox
    $Marketing.text                  = "M: Marketing"
    $Marketing.AutoSize          = $true
    $Marketing.width               = 95
    $Marketing.height              = 20
    $Marketing.location           = New-Object System.Drawing.Point(200,325)
    $Marketing.Font                 = 'Microsoft Sans Serif,10'

    $TimberlineEst                   = New-Object system.Windows.Forms.CheckBox
    $TimberlineEst.text            = "I: Timberline Est"
    $TimberlineEst.AutoSize    = $true
    $TimberlineEst.width         = 95
    $TimberlineEst.height        = 20
    $TimberlineEst.location     = New-Object System.Drawing.Point(200,350)
    $TimberlineEst.Font           = 'Microsoft Sans Serif,10'

    $HRFolder                          = New-Object system.Windows.Forms.CheckBox
    $HRFolder.text                   = "HR Folder"
    $HRFolder.AutoSize           = $true
    $HRFolder.width                = 95
    $HRFolder.height               = 20
    $HRFolder.location            = New-Object System.Drawing.Point(350,300)
    $HRFolder.Font                  = 'Microsoft Sans Serif,10'

    $DocumentFlow                      = New-Object system.Windows.Forms.CheckBox
    $DocumentFlow.text               = "F: Document Flow"
    $DocumentFlow.AutoSize       = $true
    $DocumentFlow.width            = 95
    $DocumentFlow.height           = 20
    $DocumentFlow.location        = New-Object System.Drawing.Point(350,325)
    $DocumentFlow.Font              = 'Microsoft Sans Serif,10'

    $AllDL                             = New-Object system.Windows.Forms.CheckBox
    $AllDL.text                      = "All"
    $AllDL.AutoSize              = $true
    $AllDL.width                   = 95
    $AllDL.height                  = 20
    $AllDL.location               = New-Object System.Drawing.Point(550,300)
    $AllDL.Font                     = 'Microsoft Sans Serif,10'

    $AccountingDL                      = New-Object system.Windows.Forms.CheckBox
    $AccountingDL.text               = "Accounting"
    $AccountingDL.AutoSize       = $true
    $AccountingDL.width            = 95
    $AccountingDL.height           = 20
    $AccountingDL.location        = New-Object System.Drawing.Point(550,325)
    $AccountingDL.Font              = 'Microsoft Sans Serif,10'

    $HoustonDL                          = New-Object system.Windows.Forms.CheckBox
    $HoustonDL.text                   = "Houston"
    $HoustonDL.AutoSize           = $true
    $HoustonDL.width                = 95
    $HoustonDL.height               = 20
    $HoustonDL.location            = New-Object System.Drawing.Point(550,350)
    $HoustonDL.Font                  = 'Microsoft Sans Serif,10'

    $AustinDL                          = New-Object system.Windows.Forms.CheckBox
    $AustinDL.text                   = "Austin"
    $AustinDL.AutoSize           = $true
    $AustinDL.width                = 95
    $AustinDL.height               = 20
    $AustinDL.location            = New-Object System.Drawing.Point(700,300)
    $AustinDL.Font                  = 'Microsoft Sans Serif,10'

    $EstimatingDL                      = New-Object system.Windows.Forms.CheckBox
    $EstimatingDL.text               = "Estimating"
    $EstimatingDL.AutoSize       = $true
    $EstimatingDL.width            = 95
    $EstimatingDL.height           = 20
    $EstimatingDL.location        = New-Object System.Drawing.Point(700,325)
    $EstimatingDL.Font              = 'Microsoft Sans Serif,10'

    $MarketingDL                       = New-Object system.Windows.Forms.CheckBox
    $MarketingDL.text                = "Marketing"
    $MarketingDL.AutoSize        = $true
    $MarketingDL.width             = 95
    $MarketingDL.height            = 20
    $MarketingDL.location         = New-Object System.Drawing.Point(700,350)
    $MarketingDL.Font               = 'Microsoft Sans Serif,10'

    $MicrosoftOffice                      = New-Object system.Windows.Forms.CheckBox
    $MicrosoftOffice.text               = "Microsoft Office*"
    $MicrosoftOffice.AutoSize       = $true
    $MicrosoftOffice.width            = 95
    $MicrosoftOffice.height           = 20
    $MicrosoftOffice.location        = New-Object System.Drawing.Point(50,500)
    $MicrosoftOffice.Font              = 'Microsoft Sans Serif,10'

    $License = New-Object System.Windows.Forms.Label
    $License.Location = New-Object System.Drawing.Point(50,680)
    $License.AutoSize = $true
    $License.ForeColor = "DarkRed"
    $License.Text =  "* Number of Available Busines Premium Licenses: "
    $form.Controls.Add($License)

    $LicenseCount = New-Object System.Windows.Forms.Label
    $LicenseCount.Location = New-Object System.Drawing.Point(310,680)
    $LicenseCount.AutoSize = $true
    $LicenseCount.ForeColor = "Green"
    $LicenseCount.Text =  $BPAvailableLicense
    $form.Controls.Add($LicenseCount)

    $MicrosoftProjectViewer                    = New-Object system.Windows.Forms.CheckBox
    $MicrosoftProjectViewer.text             = "Microsoft Project Viewer"
    $MicrosoftProjectViewer.AutoSize     = $true
    $MicrosoftProjectViewer.width          = 95
    $MicrosoftProjectViewer.height         = 20
    $MicrosoftProjectViewer.location      = New-Object System.Drawing.Point(50,525)
    $MicrosoftProjectViewer.Font            = 'Microsoft Sans Serif,10'

    $AdobeDC                              = New-Object system.Windows.Forms.CheckBox
    $AdobeDC.text                       = "Adobe DC"
    $AdobeDC.AutoSize               = $true
    $AdobeDC.width                    = 95
    $AdobeDC.height                   = 20
    $AdobeDC.location                = New-Object System.Drawing.Point(50,550)
    $AdobeDC.Font                      = 'Microsoft Sans Serif,10'

    $ProcoreSync                       = New-Object system.Windows.Forms.CheckBox
    $ProcoreSync.text                = "Procore Sync"
    $ProcoreSync.AutoSize        = $true
    $ProcoreSync.width             = 95
    $ProcoreSync.height            = 20
    $ProcoreSync.location          = New-Object System.Drawing.Point(50,575)
    $ProcoreSync.Font                = 'Microsoft Sans Serif,10'

    $OpenVPN                          = New-Object system.Windows.Forms.CheckBox
    $OpenVPN.text                   = "OpenVPN"
    $OpenVPN.AutoSize           = $true
    $OpenVPN.width                = 95
    $OpenVPN.height               = 20
    $OpenVPN.location            = New-Object System.Drawing.Point(50,600)
    $OpenVPN.Font                  = 'Microsoft Sans Serif,10'

    $MicrosoftProject                   = New-Object system.Windows.Forms.CheckBox
    $MicrosoftProject.text            = "Microsoft Project**"
    $MicrosoftProject.AutoSize    = $true
    $MicrosoftProject.width         = 95
    $MicrosoftProject.height        = 20
    $MicrosoftProject.location     = New-Object System.Drawing.Point(250,500)
    $MicrosoftProject.Font           = 'Microsoft Sans Serif,10'

    $ProjLicense = New-Object System.Windows.Forms.Label
    $ProjLicense.Location = New-Object System.Drawing.Point(50,660)
    $ProjLicense.AutoSize = $true
    $ProjLicense.ForeColor = "DarkRed"
    $ProjLicense.Text =  "** Number of Available Microsoft Project Licenses: "
    $form.Controls.Add($License)

    $ProjLicenseCount = New-Object System.Windows.Forms.Label
    $ProjLicenseCount.Location = New-Object System.Drawing.Point(310,660)
    $ProjLicenseCount.AutoSize = $true
    $ProjLicenseCount.ForeColor = "Green"
    $ProjLicenseCount.Text =  $ProjAvailableLicense
    $form.Controls.Add($LicenseCount)

    $MicrosoftVisio                    = New-Object system.Windows.Forms.CheckBox
    $MicrosoftVisio.text             = "Microsoft Visio"
    $MicrosoftVisio.AutoSize     = $true
    $MicrosoftVisio.width          = 95
    $MicrosoftVisio.height         = 20
    $MicrosoftVisio.location      = New-Object System.Drawing.Point(250,525)
    $MicrosoftVisio.Font            = 'Microsoft Sans Serif,10'

    $Procore                          = New-Object system.Windows.Forms.CheckBox
    $Procore.text                   = "Procore"
    $Procore.AutoSize           = $true
    $Procore.width                = 95
    $Procore.height               = 20
    $Procore.location            = New-Object System.Drawing.Point(250,550)
    $Procore.Font                  = 'Microsoft Sans Serif,10'

    $HH2                            = New-Object system.Windows.Forms.CheckBox
    $HH2.text                     = "HH2"
    $HH2.AutoSize             = $true
    $HH2.width                  = 95
    $HH2.height                 = 20
    $HH2.location              = New-Object System.Drawing.Point(250,575)
    $HH2.Font                    = 'Microsoft Sans Serif,10'

    $TimberlineAccess                    = New-Object system.Windows.Forms.CheckBox
    $TimberlineAccess.text             = "Timberline Access"
    $TimberlineAccess.AutoSize     = $true
    $TimberlineAccess.width          = 95
    $TimberlineAccess.height         = 20
    $TimberlineAccess.location      = New-Object System.Drawing.Point(250,600)
    $TimberlineAccess.Font            = 'Microsoft Sans Serif,10'


    $FistNameTextBox                             = New-Object system.Windows.Forms.TextBox
    $FistNameTextBox.multiline             = $true
    $FistNameTextBox.width                  = 100
    $FistNameTextBox.height                 = 20
    $FistNameTextBox.location              = New-Object System.Drawing.Point(50,60)
    $FistNameTextBox.Font                    = 'Microsoft Sans Serif,10'
 
    $LastNameTextBox                            = New-Object system.Windows.Forms.TextBox
    $LastNameTextBox.multiline             = $true
    $LastNameTextBox.width                  = 100
    $LastNameTextBox.height                 = 20
    $LastNameTextBox.location              = New-Object System.Drawing.Point(200,60)
    $LastNameTextBox.Font                    = 'Microsoft Sans Serif,10'

    $EmailTextbox                           = New-Object system.Windows.Forms.TextBox
    $EmailTextbox.multiline            = $true
    $EmailTextbox.width                 = 100
    $EmailTextbox.height                = 20
    $EmailTextbox.location             = New-Object System.Drawing.Point(50,110)
    $EmailTextbox.Font                   = 'Microsoft Sans Serif,10'
    

    $PhoneNumberTextBox                           = New-Object system.Windows.Forms.TextBox
    $PhoneNumberTextBox.multiline           = $true
    $PhoneNumberTextBox.width                = 100
    $PhoneNumberTextBox.height               = 20
    $PhoneNumberTextBox.location            = New-Object System.Drawing.Point(200,110)
    $PhoneNumberTextBox.Font                  = 'Microsoft Sans Serif,10'

    $JobTitleTextBox                          = New-Object system.Windows.Forms.TextBox
    $JobTitleTextBox.multiline          = $true
    $JobTitleTextBox.width               = 100
    $JobTitleTextBox.height              = 20
    $JobTitleTextBox.location           = New-Object System.Drawing.Point(50,160)
    $JobTitleTextBox.Font                 = 'Microsoft Sans Serif,10'

    $ComputerNameTextbox                              = New-Object system.Windows.Forms.TextBox
    $ComputerNameTextbox.multiline              = $true
    $ComputerNameTextbox.width                   = 100
    $ComputerNameTextbox.height                  = 20
    $ComputerNameTextbox.location               = New-Object System.Drawing.Point(200,160)
    $ComputerNameTextbox.Font                     = 'Microsoft Sans Serif,10'

    $PasswordTextbox                              = New-Object system.Windows.Forms.TextBox
    $PasswordTextbox.multiline              = $true
    $PasswordTextbox.width                   = 100
    $PasswordTextbox.height                  = 20
    $PasswordTextbox.location               = New-Object System.Drawing.Point(325,110)
    $PasswordTextbox.Font                     = 'Microsoft Sans Serif,10'
    $PasswordTextbox.PasswordChar ='*'

    $Form.controls.AddRange(@($OkButton,$Cancel,$Passwordlabel,$Passwordtextbox,$FirstNameLabel,$LastNameLabel,$EmailLabel,$PhoneNumberLabel,$JobTitleLabel,$ComputerNameLabel,$ShareDrivesLabel,$DistroListLabel,$ApplicationsInstall,$LocationComboBox,$Austin,$Houston,$User,$Scanner,$Marketing,$TimberlineEst,$HRFolder,$DocumentFlow,$AllDL,$AccountingDL,$HoustonDL,$AustinDL,$EstimatingDL,$MarketingDL,$MicrosoftOffice,$License,$LicenseCount,$MicrosoftProjectViewer,$AdobeDC,$ProcoreSync,$OpenVPN,$MicrosoftProject,$ProjLicense,$ProjLicenseCount,$MicrosoftVisio,$Procore,$HH2,$TimberlineAccess,$LocationLabel,$FistNameTextBox,$LastNameTextBox,$EmailTextbox,$PhoneNumberTextBox,$JobTitleTextBox,$ComputerNameTextbox))

    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $Script:AustinSD = $Austin.Checked -eq $true
        $Script:HoustinSD = $Houston.Checked -eq $true
        $Script:UserSD = $User.Checked -eq $true
        $Script:ScannerSD = $Scanner.Checked -eq $true
        $Script:MarketingSD = $Marketing.Checked -eq $true
        $Script:TimberlineEstSD = $TimberlineEst.Checked -eq $true
        $Script:HRFolderSD = $HRFolder.Checked -eq $true
        $Script:DocumentFlowSD = $DocumentFlow.Checked -eq $true
        $Script:All_DL = $AllDL.Checked -eq $true
        $Script:Accounting_DL = $AccountingDL.Checked -eq $true
        $Script:Houston_DL = $HoustonDL.Checked -eq $true
        $Script:Austin_DL = $AustinDL.Checked -eq $true
        $Script:Estimating_DL = $EstimatingDL.Checked -eq $true
        $Script:Marketing_DL = $MarketingDL.Checked -eq $true
        $Script:MicrosoftOffice_App = $MicrosoftOffice.Checked -eq $true
        $Script:MicrosoftProjectViewer_App = $MicrosoftProjectViewer.Checked -eq $true
        $Script:AdobeDC_App = $AdobeDC.Checked -eq $true
        $Script:ProcoreSync_App = $ProcoreSync.Checked -eq $true
        $Script:OpenVPN_App = $OpenVPN.Checked -eq $true
        $Script:MicrosoftProject_App = $MicrosoftProject.Checked -eq $true
        $Script:MicrosoftVisio_App = $MicrosoftVisio.Checked -eq $true
        $Script:Procore_App = $Procore.Checked -eq $true
        $Script:HH2_App = $HH2.Checked -eq $true
        $Script:TimberlineAccess_App = $TimberlineAccess.Checked -eq $true
        $Script:AustinLargeConferenceRoom_Cal = $AustinLargeConferenceRoom.Checked -eq $true
        $Script:AustinSmallConferenceRoom_Cal = $AustinSmallConferenceRoom.Checked -eq $true
        $Script:AustinTrainingRoom_Cal = $AustinTrainingRoom.Checked -eq $true
        $Script:PTOCalendar_Cal = $PTOCalendar.Checked -eq $true
        $Script:FN = $FistNameTextBox.Text
        $Script:LN = $LastNameTextBox.Text
        $Script:Email = $EmailTextbox.Text
        $Script:PN = $PhoneNumberTextBox.Text
        $Script:JT = $JobTitleTextBox.Text
        $Script:CN = $ComputerNameTextbox.Text
        $Script:Location = $LocationComboBox.Text
        $script:password = $PasswordTextbox.Text
        $script:Location = $LocationComboBox.Text
    }
    elseif ($result -eq [System.Windows.Forms.DialogResult]::Cancel) {
        Write-Host "Goodbye!"
        Start-Sleep -Seconds 1
        Exit    
    }
  
}


#Function to process the userform
function process-userform
{
    #Calling userform function
    Prompt-userform
    #Formatting OU for AD Location
    $OU = "OU=$Location"
    #Formatting password to Secure String
    try {
        $SecurePassword = ConvertTo-SecureString $password -AsPlainText -Force -ErrorAction Stop
    }
    catch {
        Write-Host "Password field is required!"
        Return process-userform
    }
    #Formatting SAM account name
    $SAM = "$FN.$LN"

    #Create AD User
    try {
        New-ADUser -Name "$FN $LN" -GivenName "$FN" -Surname "$LN" -SamAccountName "$SAM" -UserPrincipalName "$Email" -Title "$JT" -Path "$OU,OU=Employees,DC=Pyroforces,DC=net" -AccountPassword $securepassword -HomeDirectory \\Pyro-DC\users\$SAM -HomeDrive U  -Enabled $true -ErrorAction Stop
        Write-Host "Sucessfully created user in Active Directory" -ForegroundColor Green
    }
    catch {
        Write-Host `n"Error when creating user in Active Directory" -ForegroundColor Red
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
        Write-Host `n"$ErrorMessage"
        Write-Host `n"$FailedItem"
        $WhatNow = Read-Host "Continue? (y/n) "
        if ($WhatNow -eq "y") {
            Write-Host "Continuing script..."
            Start-Sleep -Seconds 4
        }
        elseif ($WhatNow -eq "n") {
            Write-Host "Terminating script...Goodbye!"
            Start-Sleep -Seconds 2
            Exit
        }
    }

    #Add  AD Memberships
    #Add to Austin Share
    if ($AustinSD -eq "true") {
        try {
            Add-ADGroupMember -Identity Share_AustinOffice -Members $SAM -ErrorAction Continue
            Write-Host `n"Added $SAM to O: Austin Office Group" -ForegroundColor Green
        }
        catch {
            Write-Host `n"Unable to add $SAM to O: Austin Office Group" -ForegroundColor Red
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Write-Host `n"$ErrorMessage"
            Write-Host `n"$FailedItem"                
        }

    }
    Elseif ($AustinSD -eq "false") {
        Write-Host ""
    }

    #Add to Houston Share    
    if ($HoustinSD -eq "true") {
        try {
            Add-ADGroupMember -Identity Share_HoustonOffice -Members $SAM -ErrorAction Continue
            Write-Host `n"Added $SAM to H: Houstin Office Group" -ForegroundColor Green
        }
        catch {
            Write-Host `n"Unable to add $SAM to H: Houstin Office Group" -ForegroundColor Red
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Write-Host `n"$ErrorMessage"
            Write-Host `n"$FailedItem"                
        }

    }
    Elseif ($HoustinSD -eq "false") {
        Write-Host ""
    }


    #Add to Scanner Share
    if ($ScannerSD -eq "true") {
        try {
            Add-ADGroupMember -Identity Share_Scanner -Members $SAM -ErrorAction Continue
            Write-Host `n"Added $SAM to S: Scanner Group" -ForegroundColor Green
        }
        catch {
            Write-Host `n"Unable to add $SAM to S: Scanner Group" -ForegroundColor Red
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Write-Host `n"$ErrorMessage"
            Write-Host `n"$FailedItem"                
        }

    }
    Elseif ($ScannerSD -eq "false") {
        Write-Host ""
    }


    #Add to Marketing Share
    if ($MarketingSD -eq "true") {
        try {
            Add-ADGroupMember -Identity Share_Marketing -Members $SAM -ErrorAction Continue
            Write-Host `n"Added $SAM to M: Marketing Group" -ForegroundColor Green
        }
        catch {
            Write-Host `n"Unable to add $SAM to M: Marketing Group" -ForegroundColor Red
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Write-Host `n"$ErrorMessage"
            Write-Host `n"$FailedItem"               
        }

    }
    Elseif ($MarketingSD -eq "false") {
        Write-Host ""
    }


    #Add to Timberline Estimating Share
    if ($TimberlineEstSD -eq "true") {
        try {
            Add-ADGroupMember -Identity Share_TimberlineEstimating -Members $SAM -ErrorAction Continue
            Write-Host `n"Added $SAM to I: Timberline Estimating Group" -ForegroundColor Green
        }
        catch {
            Write-Host `n"Unable to add $SAM to I: Timberline Estimating Group" -ForegroundColor Red
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Write-Host `n"$ErrorMessage"
            Write-Host `n"$FailedItem"                
        }

    }
    Elseif ($TimberlineEstSD -eq "false") {
        Write-Host ""
    }


    #Add to Human Resources Share
    if ($HRFolderSD -eq "true") {
        try {
            Add-ADGroupMember -Identity "Human Resources" -Members $SAM -ErrorAction Continue
            Write-Host `n"Added $SAM to Human Resources Group" -ForegroundColor Green
        }
        catch {
            Write-Host `n"Unable to add $SAM to Human Resources Group" -ForegroundColor Red
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Write-Host `n"$ErrorMessage"
            Write-Host `n"$FailedItem"               
        }

    }
    Elseif ($HRFolderSD -eq "false") {
        Write-Host ""
    }


    #Add to Document Flow group
    if ($DocumentFlowSD -eq "true") {
        try {
            Add-ADGroupMember -Identity Share_DocumentFlow -Members $SAM -ErrorAction Continue
            Write-Host `n"Added $SAM to Document Flow Group" -ForegroundColor Green
        }
        catch {
            Write-Host `n"Unable to add $SAM to Document Flow Group" -ForegroundColor Red
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Write-Host `n"$ErrorMessage"
            Write-Host `n"$FailedItem"                
        }

    }
    Elseif ($DocumentFlowSD -eq "false") {
        Write-Host ""
    }

    #Add to OpenVPN group
    if ($OpenVPN_App -eq "true") {
        try {
            Add-ADGroupMember -Identity "VPN Users" -Members $SAM -ErrorAction Continue
            Write-Host `n"Added $SAM to VPN Users Group" -ForegroundColor Green
        }
        catch {
            Write-Host `n"Unable to add $SAM to VPN Users Group" -ForegroundColor Red
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Write-Host `n"$ErrorMessage"
            Write-Host `n"$FailedItem"              
        }

    }
    Elseif ($OpenVPN_App -eq "false") {
        Write-Host ""
    }


        #Add to Timberline Users group
    if ($TimberlineAccess_App -eq "true") {
        try {
            Add-ADGroupMember -Identity "Timberline Users" -Members $SAM -ErrorAction Continue
            Write-Host `n"Added $SAM to Timberline Users Group" -ForegroundColor Green
        }
        catch {
            Write-Host `n"Unable to add $SAM to Timberline Users Group" -ForegroundColor Red
                Write-Host `n"$ErrorMessage"
                Write-Host `n"$FailedItem"                
        }

    }
    Elseif ($TimberlineAccess_App -eq "false") {
        Write-Host ""
    }



            #Add to Timberline Users group
    if ($UserSD -eq "true") {
        try {
            New-Item -Path "\\pyro-dc\users\$SAM" -Name $UserName -ItemType Directory -ErrorAction Continue | Out-Null 
            $acl = Get-Acl \\pyro-dc\users\$SAM 
            $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("pyroforces.net\$sam","FullControl","Allow")
            $acl.SetAccessRule($AccessRule)
            $acl | Set-Acl "\\pyro-dc\users\$SAM"
                       
            Write-Host `n"Created $SAM U: User share" -ForegroundColor Green
        }
        catch {
            Write-Host `n"Issues creating $SAM User Share" -ForegroundColor Red
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Write-Host `n"$ErrorMessage"
            Write-Host `n"$FailedItem"                
        }

    }
    Elseif ($UserSD -eq "false") {
        Write-Host ""
    }
    

    
    #Forcing AD Sync
    try {
        Write-Host `n"Forcing AD Sync...please wait"
        Start-ADSyncSyncCycle -PolicyType Delta -ErrorAction Continue
        Write-Host `n""
        
        }
    catch {
        Write-Host "Issues running AD Sync..."
    }

    function Start-SyncTimer {
    $seconds = "120"
        $doneDT = (Get-Date).AddSeconds($seconds)
        while($doneDT -gt (Get-Date)) {
            $secondsLeft = $doneDT.Subtract((Get-Date)).TotalSeconds
            $percent = ($seconds - $secondsLeft) / $seconds * 100
            Write-Progress -Activity "Verifying $Email exists before adding O365 memberships" -Status "...waiting..." -SecondsRemaining $secondsLeft -PercentComplete $percent
            [System.Threading.Thread]::Sleep(500)
        }
        Write-Progress -Activity "Verifying $Email exists before adding O365 memberships" -Status "Waiting:" -SecondsRemaining 0 -Completed

        #Checking for mailbox
        $AzureADMembers = Get-AzureADUser | select -ExpandProperty userprincipalname
        if($AzureADMembers -contains $email ) {
            Write-Host "$email successfully sync'd to O365" 
                    
        }
        Elseif($AzureADMembers -notcontains $email) {
            Write-Host "$Email has still not been added to O365...please standby" -ForegroundColor Yellow
            Return Start-SyncTimer
        }

    }

    #Displays sleep timer
    Start-SyncTimer

    ####Adding Office 365 Memberships      
    #Adding user to All Distro List
 
    if ($All_DL -eq "true") {
        try {
            Add-DistributionGroupMember "All" -Member $Email
            Write-Host `n"Added $SAM to the All Distrobution List" -ForegroundColor Green

        }
    catch {
        Write-Host `n"Unable to add $Email to the 'All' Distrobution list" -ForegroundColor Red
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
        Write-Host `n"$ErrorMessage"
        Write-Host `n"$FailedItem"              
        }  
    }
    Elseif ($All_DL -eq "false") {
        Write-Host ""
    }


    #Adding user to Austin Distro List
    if ($Austin_DL -eq "true") {
        try {
            Add-DistributionGroupMember "Austin" -Member $Email
            Write-Host `n"Added $SAM to the Austin Distrobution List" -ForegroundColor Green

        }
        catch {
            Write-Host `n"Unable to add $Email to the 'Austin' Distrobution list" -ForegroundColor Red
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Write-Host `n"$ErrorMessage"
            Write-Host `n"$FailedItem"               
        }  
    }
    Elseif ($Austin_DL -eq "false") {
        Write-Host ""
    }


    #Adding user to Houston Distro List
    if ($Houston_DL -eq "true") {
        try {
            Add-DistributionGroupMember "Houston" -Member $Email
            Write-Host `n"Added $SAM to the Houston Distrobution List" -ForegroundColor Green

        }
        catch {
            Write-Host `n"Unable to add $Email to the 'Houston' Distrobution list" -ForegroundColor Red
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Write-Host `n"$ErrorMessage"
            Write-Host `n"$FailedItem"               
        }  
    }
    Elseif ($Houston_DL -eq "false") {
        Write-Host ""
    }

    #Adding user to Marketing Distro List
    if ($Marketing_DL -eq "true") {
        try {
            Add-DistributionGroupMember "Marketing" -Member $Email
            Write-Host `n"Added $SAM to the Marketing Distrobution List" -ForegroundColor Green

        }
        catch {
            Write-Host `n"Unable to add $Email to the 'Marketing' Distrobution list" -ForegroundColor Red
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Write-Host `n"$ErrorMessage"
            Write-Host `n"$FailedItem"                
        }  
    }
    Elseif ($Marketing_DL -eq "false") {
        Write-Host ""
    }

    #Adding user to Accounting Distro List
    if ($Accounting_DL -eq "true") {
        try {
            Add-DistributionGroupMember "Accounting" -Member $Email
            Write-Host `n"Added $SAM to the Accounting Distrobution List" -ForegroundColor Green

        }
        catch {
            Write-Host `n"Unable to add $Email to the 'Accounting' Distrobution list" -ForegroundColor Red
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Write-Host `n"$ErrorMessage"
            Write-Host `n"$FailedItem"                 
        }  
    }
    Elseif ($Accounting_DL -eq "false") {
        Write-Host ""
    }

    #Adding Microsoft License
    if ($MicrosoftOffice_App -eq "true") {
        try {
            Set-MsolUser -UserPrincipalName "$email" -UsageLocation "US" -ErrorAction Continue
            Set-MsolUserLicense -UserPrincipalName "$email" -AddLicenses "Pyroforces:ENTERPRISEPACK" 
            Write-Host `n"Successfully added Business Premium License" -ForegroundColor Green
        }
        catch {
            Write-Host "Error adding Office License"
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Write-Host `n"$ErrorMessage"
            Write-Host `n"$FailedItem" 

        }
    }
     elseif ($Microsoft_App -eq "false") {
        Write-Host ""
    }


    #Move application installers to $computer
    Start-Sleep -Seconds 2
    
    $Software = $AdobeDC_App, $ProcoreSync_App, $OpenVPN_App, $MicrosoftProject_App, $MicrosoftVisio_App, $Procore_App, $HH2_App, $TimberlineAccess_App

    $App1 = "\\pyro-dc\Installs\openvpn.exe"
    $App2 = "\\pyro-dc\Installs\ChromeSetup.exe"
    $App3 = "\\pyro-dc\Installs\AdobeAcrobat.exe"
    $App4 = "\\pyro-dc\Installs\MicrosoftOffice.exe"
    $App5 = "\\pyro-dc\Installs\MSProject.exe "
    $App6 = "\\pyro-dc\Installs\ProcoreSync.exe "
    $App7 = "\\pyro-dc\Installs\Agent_Austin.MSI"
    $App8 = "\\pyro-dc\Installs\Agent_Houston.MSI"
    $App9 = "\\pyro-dc\Installs\Agent_SanAntonio.MSI"
    
    if ($Software -contains "true") {
        Write-Host `n"Preparing to transfer requested install files to $CN C:\installs"
        Start-Sleep -Seconds 2
        try {
            Invoke-Command -ComputerName $CN -ScriptBlock {
                New-Item -Path "c:\" -Name "Installs" -ItemType "directory"
                Start-Sleep -Seconds 5
            } -ErrorAction Continue
            
            Write-Host `n"Created 'C:\Installs' Folder" -ForegroundColor Green
            
            Write-Host `n"Copying standard applications" -ForegroundColor Green
            
            #Copying Google Chrome Installer
            Copy-Item -Path $App2 -Destination "\\$CN\C$\Installs" 
            Write-Host `n"Copied Google Chrome Installer" -ForegroundColor Green               
            
            #Copying MITP Agent Installer
            if ($location -eq "Austin") {
                Copy-Item -Path $App7 -Destination "\\$CN\C$\Installs"
                Write-Host `n"Copied MITP Austin Agent" -ForegroundColor Green  
            }
            elseif ($Location -eq "Houston") {
                Copy-Item -Path $App8 -Destination "\\$CN\C$\Installs"
                Write-Host `n"Copied MITP Houston Agent"  -ForegroundColor Green
            }
            elseif ($Location -eq "San Antonio") {
                Copy-Item -Path $App9 -Destination "\\$CN\C$\Installs"
                Write-Host `n"Copied MITP San Antion Agent"  -ForegroundColor Green
            }


        }

        catch {
            Invoke-Command -ComputerName -ScriptBlock {
                Write-Host "Unable to create install folder"
                $ErrorMessage = $_.Exception.Message
                $FailedItem = $_.Exception.ItemName
                Write-Host `n"$ErrorMessage"
                Write-Host `n"$FailedItem" 

               Start-Sleep -Seconds 30
               exit
            }
            
        }


        ###Transferring requested software installs 

        #Copy OpenVPN
        if($OpenVPN_App -eq "true") {
            try {                
                Copy-Item -Path $App1 -Destination "\\$CN\C$\Installs" -ErrorAction Continue 
                Write-Host `n"Copied OpenVPN Installer" -ForegroundColor Green
            }

            catch {
                Write-Host "Issues transferring $App1 "
                $ErrorMessage = $_.Exception.Message
                $FailedItem = $_.Exception.ItemName
                Write-Host `n"$ErrorMessage"
                Write-Host `n"$FailedItem"        
            }

        }
        elseif($OpenVPN_App -eq "false") {
            Write-Host ""
        }

        #Copy ProCore Sync
        if($ProcoreSync_App -eq "true") {
            try {
                Copy-Item -Path $App2 -Destination "\\$CN\C$\Installs" -ErrorAction Continue 
                Write-Host `n"Copied Procore Sync Installer" -ForegroundColor Green
            }

            catch {
                Write-Host `n"Issues transferring $App2 "
                $ErrorMessage = $_.Exception.Message
                $FailedItem = $_.Exception.ItemName
                Write-Host `n"$ErrorMessage"
                Write-Host `n"$FailedItem"        
            }

        }
        elseif($ProcoreSync_App -eq "false") {
            Write-Host ""
        }

        #Copy Adobe install
        if($AdobeDC_App -eq "true") {
            try {
                Copy-Item -Path "$App3" -Destination "\\$CN\C$\Installs" -ErrorAction Continue 
                Write-Host `n"Copied Adobe Acrobat Installer" -ForegroundColor Green
            }

            catch {
                Write-Host "Issues transferring $App3 "
                $ErrorMessage = $_.Exception.Message
                $FailedItem = $_.Exception.ItemName
                Write-Host `n"$ErrorMessage"
                Write-Host `n"$FailedItem"         
            }

        }
        elseif($AdobeDC_App -eq "false") {
            Write-Host ""
        }

        #Copy Microsoft Office install
        if($MicrosoftOffice_App -eq "true") {
            try {             
                Copy-Item -Path "$App4" -Destination "\\$CN\C$\Installs" -ErrorAction Continue 
                Write-Host `n"Copied Microsoft Office Installer" -ForegroundColor Green
            }

            catch {
                Write-Host "Issues transferring $App4 "
                $ErrorMessage = $_.Exception.Message
                $FailedItem = $_.Exception.ItemName
                Write-Host `n"$ErrorMessage"
                Write-Host `n"$FailedItem"         
            }

        }
        elseif($MicrosoftOffice_App -eq "false") {
            Write-Host ""
        }

        #Copy Microsoft Project
        if($MicrosoftProject_App -eq "true") {
            try {
                Copy-Item -Path "$App5" -Destination "\\$CN\C$\Installs" -ErrorAction Continue 
                Write-Host `n"Copied Microsoft Project Installer" -ForegroundColor Green
            }

            catch {
                Write-Host "Issues transferring $App5 "
                $ErrorMessage = $_.Exception.Message
                $FailedItem = $_.Exception.ItemName
                Write-Host `n"$ErrorMessage"
                Write-Host `n"$FailedItem"         
            }

        }
        elseif($MicrosoftProject_App -eq "false") {
            Write-Host ""
        }

        <###Copy Template###
        if($AdobeDC_App -eq "true") {
            try {
                
                    Copy-Item -Path $App3 -Destination "\\$CN\C$\Installs" -ErrorAction Continue 
                    Write-Host "Copied Adobe Acrobat Installer"
            }

            catch {
            Write-Host "Issues transferring $App3 "        
            }

       }
       elseif($AdobeDC_App -eq "false") {
            Write-Host ""
       } #>



    #End of Software installs. 
    }

    elseif ($Software -notcontains "true") {
        Write-Host "No Applications requested."

    } 


#End of Process-userform function 
}

Prompt-warning
Connect-O365

process-userform


Write-Host "End of user creation"
Get-PSSession | Remove-PSSession

Start-Sleep -Seconds 60



 