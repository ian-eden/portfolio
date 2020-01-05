### Onsite User Create
### 12/23/2019
### Ian Eden
### Version 2.3




#Launch PowerShell As Adminsitrator
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
    $UserCredential = Get-Credential
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
    try {
        Import-PSSession $Session -DisableNameChecking -ErrorAction Stop
    }
    catch {
        Write-Host " `n Error connecting to O365...Please check username and password..." -ForegroundColor Red
        Start-Sleep -Seconds 1
        Clear-Host
        Return Connect-O365
    } 
    
    #Install AzureAD Module
    if (Get-Module -ListAvailable -Name "AzureAD") { 
        Write-Host "AzureAD Module Already installed" -ForegroundColor Green
    } 
    Else {
        Write-Host "Installing Module 'AzureAD'..." -ForegroundColor Yellow
        Install-Module AzureAD
    }
    
        if (Get-Module -ListAvailable -Name "msonline") { 
        Write-Host "MSonline Module Already installed" -ForegroundColor Green
    } 
    Else {
        Write-Host "Installing Module 'MSonline'..." -ForegroundColor Yellow
        Install-Module msonline
    }
            #Connect to Azure-AD and MSOL
    Connect-AzureAD -Credential $UserCredential
    Connect-MsolService -Credential $UserCredential

}

#Function to prompt Dialog box and take information
function Display-Dialog {
    #Get available Licenses
    $TotalLicense = Get-MsolAccountSku | Where-Object AccountSkuID -eq "reseller-account:STANDARDPACK" | Select-Object ActiveUnits | Out-String
    $ConsumedLicense = Get-MsolAccountSku | Where-Object AccountSkuID -eq "reseller-account:STANDARDPACK"| Select-Object ConsumedUnits | Out-String
    ### Enabled for testing only.
    #$TotalLicense = Get-MsolAccountSku | Where-Object AccountSkuID -eq "Pyroforces:ENTERPRISEPACK" | Select-Object ActiveUnits | Out-String
    #$ConsumedLicense = Get-MsolAccountSku | Where-Object AccountSkuID -eq "Pyroforces:ENTERPRISEPACK"| Select-Object ConsumedUnits | Out-String
    [int]$TrimTotalLicense = $TotalLicense.Trim("ActiveUnits
    ----------- ")
    [int]$TrimConsumedLicense = $ConsumedLicense.Trim("ConsumedUnits
    -------------
            ")
    #Subtracting Total Licenses from Consumed
    $script:AvailableLicense = $TrimTotalLicense - $TrimConsumedLicense

    #If there are no available licenses Show Warning dialog box and then exit script
    if ($AvailableLicense -eq  0)  { 
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        $warning = New-Object System.Windows.Forms.Form
        $warning.Text = '                      No Available Licenses!'
        $warning.BackColor = "DarkRed"
        $warning.Size = New-Object System.Drawing.Size(400,220)
        $warning.StartPosition = 'CenterScreen'

        $OKButton = New-Object System.Windows.Forms.Button
        $OKButton.Location = New-Object System.Drawing.Point(150,100)
        $OKButton.Size = New-Object System.Drawing.Size(75,23)
        $OKButton.BackColor = "white"
        $OKButton.Text = 'Exit'
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $warning.AcceptButton = $OKButton
        $warning.Controls.Add($OKButton)



        $label = New-Object System.Windows.Forms.Label
        $label.Location = New-Object System.Drawing.Point(60,20)
        $label.Size = New-Object System.Drawing.Size(280,40)
        $label.ForeColor = "White"
        $label.BorderStyle = "Fixed3D"
        $label.Text = `n'Please add E1 Licenses in Synnex before continuing'
        $warning.Controls.Add($label)

        $warning.Topmost = $true

        $result = $warning.ShowDialog()

        if ($result -eq [System.Windows.Forms.DialogResult]::OK)
            {   
            Write-Host "Exiting..."
            Exit
        }
    }
    Else 
    {
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        $form = New-Object System.Windows.Forms.Form
        $form.Text = 'Roscoe Properties User Creation'
        $form.Size = New-Object System.Drawing.Size(580,500)
        $form.StartPosition = 'CenterScreen'

        $OKButton = New-Object System.Windows.Forms.Button
        $OKButton.Location = New-Object System.Drawing.Point(450,400)
        $OKButton.Size = New-Object System.Drawing.Size(75,23)
        $OKButton.Text = 'Submit'
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.AcceptButton = $OKButton
        $form.Controls.Add($OKButton)

        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = New-Object System.Drawing.Point(350,400)
        $CancelButton.Size = New-Object System.Drawing.Size(75,23)
        $CancelButton.Text = 'Cancel'
        $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.CancelButton = $CancelButton
        $form.Controls.Add($CancelButton)

        $FirstName = New-Object System.Windows.Forms.Label
        $FirstName.Location = New-Object System.Drawing.Point(10,20)
        $FirstName.Size = New-Object System.Drawing.Size(280,20)
        $Firstname.ForeColor = "Red"
        $FirstName.Text = 'First name:'
        $form.Controls.Add($FirstName)

        $FirstNametextBox = New-Object System.Windows.Forms.TextBox
        $FirstNametextBox.Location = New-Object System.Drawing.Point(10,40)
        $FirstNametextBox.Size = New-Object System.Drawing.Size(240,20)
        $form.Controls.Add($FirstNametextBox)

        $LastName = New-Object System.Windows.Forms.Label
        $LastName.Location = New-Object System.Drawing.Point(300,20)
        $LastName.Size = New-Object System.Drawing.Size(280,20)
        $LastName.ForeColor = "Red"
        $LastName.Text = 'Last name:'
        $form.Controls.Add($LastName)


        $LastNametextBox = New-Object System.Windows.Forms.TextBox
        $LastNametextBox.Location = New-Object System.Drawing.Point(300,40)
        $LastNametextBox.Size = New-Object System.Drawing.Size(240,20)
        $form.Controls.Add($LastNametextBox)

        $Email = New-Object System.Windows.Forms.Label
        $Email.Location = New-Object System.Drawing.Point(10,80)
        $Email.Size = New-Object System.Drawing.Size(280,20)
        $Email.ForeColor = "Red"
        $Email.Text = 'Email:'
        $form.Controls.Add($Email)
        
        $EmailtextBox = New-Object System.Windows.Forms.TextBox
        $EmailtextBox.Location = New-Object System.Drawing.Point(10,100)
        $EmailtextBox.Size = New-Object System.Drawing.Size(100,20)
        $form.Controls.Add($EmailtextBox)

        $DomainDropDown = new-object System.Windows.Forms.ComboBox
        $DomainDropDown.Location = new-object System.Drawing.Size(120,100)
        $DomainDropDown.Size = new-object System.Drawing.Size(130,20)
        $DomainDropDown.Text = "@rpmliving.com"


        #[void] $DomainDropDown.Items.Add('@pyroforces.net')
        [void] $DomainDropDown.Items.Add('@rpmliving.com')
        [void] $DomainDropDown.Items.Add('@roscoeproperties.com')
        [void] $DomainDropDown.Items.Add('@fandbcapital.com')
        [void] $DomainDropDown.Items.Add('@oneeightyconstruction.com')
        [void] $DomainDropDown.Items.Add('@roscoeprop.com')
        [void] $DomainDropDown.Items.Add('@tuckerutilities.com')
        [void] $DomainDropDown.Items.Add('@tuckerutilities.com')

        $JobTitle = New-Object System.Windows.Forms.Label
        $JobTitle.Location = New-Object System.Drawing.Point(300,80)
        $JobTitle.Size = New-Object System.Drawing.Size(280,20)
        $JobTitle.Text = 'Job Title:'
        $form.Controls.Add($JobTitle)

        $JobTitletextBox = New-Object System.Windows.Forms.TextBox
        $JobTitletextBox.Location = New-Object System.Drawing.Point(300,100)
        $JobTitletextBox.Size = New-Object System.Drawing.Size(240,20)
        $form.Controls.Add($JobTitletextBox)

        $BusinessPhone = New-Object System.Windows.Forms.Label
        $BusinessPhone.Location = New-Object System.Drawing.Point(10,140)
        $BusinessPhone.Size = New-Object System.Drawing.Size(280,20)
        $BusinessPhone.Text = 'Business Phone:'
        $form.Controls.Add($BusinessPhone)

        $BusinessPhoneTextBox = New-Object System.Windows.Forms.TextBox
        $BusinessPhoneTextBox.Location = New-Object System.Drawing.Point(10,160)
        $BusinessPhoneTextBox.Size = New-Object System.Drawing.Size(240,20)
        $form.Controls.Add($BusinessPhoneTextBox)

        $CellPhone = New-Object System.Windows.Forms.Label
        $CellPhone.Location = New-Object System.Drawing.Point(300,140)
        $CellPhone.Size = New-Object System.Drawing.Size(280,20)
        $CellPhone.Text = 'Cell Phone:'
        $form.Controls.Add($CellPhone)

        $CellPhonetextBox = New-Object System.Windows.Forms.TextBox
        $CellPhonetextBox.Location = New-Object System.Drawing.Point(300,160)
        $CellPhonetextBox.Size = New-Object System.Drawing.Size(240,20)
        $form.Controls.Add($CellPhonetextBox)

        $Location = New-Object System.Windows.Forms.Label
        $Location.Location = New-Object System.Drawing.Point(10,200)
        $Location.Size = New-Object System.Drawing.Size(280,20)
        $Location.Text = 'Location:'
        $form.Controls.Add($Location)

        $LocationTextBox = New-Object System.Windows.Forms.TextBox
        $LocationTextBox.Location = New-Object System.Drawing.Point(10,220)
        $LocationTextBox.Size = New-Object System.Drawing.Size(240,20)
        $form.Controls.Add($LocationTextBox)


        $Department = New-Object System.Windows.Forms.Label
        $Department.Location = New-Object System.Drawing.Point(300,200)
        $Department.Size = New-Object System.Drawing.Size(280,20)
        $Department.Text = 'Department:'
        $form.Controls.Add($Department)

        $DepartmenttextBox = New-Object System.Windows.Forms.TextBox
        $DepartmenttextBox.Location = New-Object System.Drawing.Point(300,220)
        $DepartmenttextBox.Size = New-Object System.Drawing.Size(240,20)
        $DepartmenttextBox.Text = "Onsite"
        $form.Controls.Add($DepartmenttextBox)


        $Address = New-Object System.Windows.Forms.Label
        $Address.Location = New-Object System.Drawing.Point(10,260)
        $Address.Size = New-Object System.Drawing.Size(280,20)
        $Address.Text = 'Street Address:'
        $form.Controls.Add($Address)

        $AddressTextBox = New-Object System.Windows.Forms.TextBox
        $AddressTextBox.Location = New-Object System.Drawing.Point(10,280)
        $AddressTextBox.Size = New-Object System.Drawing.Size(240,20)
        $form.Controls.Add($AddressTextBox)


        $City = New-Object System.Windows.Forms.Label
        $City.Location = New-Object System.Drawing.Point(300,260)
        $City.Size = New-Object System.Drawing.Size(50,20)
        $City.Text = 'City:'
        $form.Controls.Add($City)

        $CitytextBox = New-Object System.Windows.Forms.TextBox
        $CitytextBox.Location = New-Object System.Drawing.Point(300,280)
        $CitytextBox.Size = New-Object System.Drawing.Size(100,20)
        $form.Controls.Add($CitytextBox)

        $State = New-Object System.Windows.Forms.Label
        $State.Location = New-Object System.Drawing.Point(425,260)
        $State.Size = New-Object System.Drawing.Size(50,20)
        $State.Text = 'State:'
        $form.Controls.Add($State)

        $StatetextBox = New-Object System.Windows.Forms.TextBox
        $StatetextBox.Location = New-Object System.Drawing.Point(425,280)
        $StatetextBox.Size = New-Object System.Drawing.Size(30,20)
        $form.Controls.Add($StatetextBox)

        $Zip = New-Object System.Windows.Forms.Label
        $Zip.Location = New-Object System.Drawing.Point(475,260)
        $Zip.Size = New-Object System.Drawing.Size(50,20)
        $Zip.Text = 'ZIP:'
        $form.Controls.Add($Zip)

        $ZiptextBox = New-Object System.Windows.Forms.TextBox
        $ZiptextBox.Location = New-Object System.Drawing.Point(475,280)
        $ZiptextBox.Size = New-Object System.Drawing.Size(65,20)
        $form.Controls.Add($ZiptextBox)

        #Shared Mailbox
        #$Group = New-Object System.Windows.Forms.Label
        #$Group.Location = New-Object System.Drawing.Point(10,310)
        #$Group.Size = New-Object System.Drawing.Size(280,20)
        #$Group.Text = 'Add to shared mailbox:'
        #$form.Controls.Add($Group)

        #$GroupTextBox = New-Object System.Windows.Forms.TextBox
        #$GroupTextBox.Location = New-Object System.Drawing.Point(10,330)
        #$GroupTextBox.Size = New-Object System.Drawing.Size(240,20)
        #$form.Controls.Add($GroupTextBox)


        $License = New-Object System.Windows.Forms.Label
        $License.Location = New-Object System.Drawing.Point(300,330)
        $License.Size = New-Object System.Drawing.Size(180,20)
        $License.ForeColor = "DarkRed"
        $License.Text =  "Number of Available E1 Licenses: "
        $form.Controls.Add($License)

        $LicenseCount = New-Object System.Windows.Forms.Label
        $LicenseCount.Location = New-Object System.Drawing.Point(480,330)
        $LicenseCount.Size = New-Object System.Drawing.Size(20,20)
        $LicenseCount.ForeColor = "Green"
        $LicenseCount.Text =  $AvailableLicense
        $form.Controls.Add($LicenseCount)


        $Required = New-Object System.Windows.Forms.Label
        $Required.Location = New-Object System.Drawing.Point(10,400)
        $Required.Size = New-Object System.Drawing.Size(180,20)
        $Required.ForeColor = "Red"
        $Required.Text =  "Fields in red are required."

        $form.Controls.Add($DomainDropDown)
        $form.Controls.Add($Required)


        $form.Topmost = $true

        $form.Add_Shown({$FirstNametextBox.Select()})
        $form.Add_Shown({$LastNametextBox.Select()})
        $form.Add_Shown({$EmailtextBox.Select()})
        $form.Add_Shown({$JobTitletextBox.Select()})
        $form.Add_Shown({$BusinessPhoneTextBox.Select()})
        $form.Add_Shown({$CellPhonetextBox.Select()})
        $form.Add_Shown({$LocationTextBox.Select()})
        $form.Add_Shown({$DepartmentTextBox.Select()})
        $form.Add_Shown({$AddressTextBox.Select()})
        $form.Add_Shown({$CityTextBox.Select()})
        $form.Add_Shown({$StateTextBox.Select()})
        $form.Add_Shown({$ZipTextBox.Select()})
        $form.Add_Shown({$DomainDropDown.Select()})
        #$form.Add_Shown({$GroupTextBox.Select()})


        $result = $form.ShowDialog()

        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            $script:FN = $FirstNametextBox.Text
            $script:LN = $LastNametextBox.Text
            $script:EM = $EmailtextBox.Text
            $script:JT = $JobTitletextBox.Text
            $script:BP = $BusinessPhoneTextBox.Text
            $script:CP = $CellPhoneTextBox.Text
            $script:LOC = $LocationTextBox.Text
            $script:Dep = $DepartmentTextBox.Text
            $script:Addr = $AddresstextBox.Text
            $script:City = $CitytextBox.Text
            $script:State = $StatetextBox.Text
            $script:Zip = $ZiptextBox.Text
            $script:EmailDomain = $DomainDropDown.Text
            #$script:Group = $GroupTextBox.Text


        }
        elseif ($result -eq [System.Windows.Forms.DialogResult]::Cancel) {
            Write-Host "Goodbye!"
            Start-Sleep -Seconds 1
            Exit    
        }
    }

}

#Function to process the entered data
function Process-User {
    #Executes function to gather information
    Display-Dialog
    
    #Stripping white spaces from variables
    $EM_S = $EM -replace '\s',''
    $EmailDomain_S = $EmailDomain -replace '\s',''


    $UPN = "$EM_S$EmailDomain_S"
    

    #Checking if mailbox already exists
    if ($Mailboxes -contains $UPN) {
        Write-Host "$UPN already exists" -ForegroundColor Yellow
        [System.Console]::Beep(900,800)
        return Process-User        
    }

    Elseif ($Mailboxes -notcontains $UPN) {
        #Create user in O365 with entered information
        try {
            New-MsolUser -DisplayName "$FN $LN" -FirstName "$FN" -LastName "$LN" -UserPrincipalName "$UPN" -Office "$LOC" -Department "$Dep" -UsageLocation "US" -StreetAddress "$Addr" -City "$City" -State "$State" -PostalCode "$Zip" -PhoneNumber "$BP" -MobilePhone "$CP" -Title "$JT" -LicenseAssignment "reseller-account:STANDARDPACK" -Password "*****" -ForceChangePassword $false -ErrorAction Stop

            Write-Host `n"Successfully created user" -ForegroundColor Green
        }
        catch {
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Write-Host `r"Error encountered while processing the user creation but the account still may have been created. Please verify and try again." -ForegroundColor DarkRed
            Write-Host `n"$ErrorMessage"
            Write-Host `n"$FailedItem"
            Return Process-User
        } 
        #Notifies user that the script is finished
        Write-Host `n"End of script. Enter another user or select 'Cancel' to exit" -ForegroundColor Yellow
        Start-Sleep -Seconds 1
        #Execute function to gather information and create the user in O365
        Return Process-User
    }
}
     

#Execute sign into 0365 PSSession function
Connect-O365
Write-Host `n"Loading...Please wait..." -ForegroundColor Yellow
$script:Mailboxes = Get-Mailbox | select -expand primarysmtpaddress
Write-Host "...complete" -ForegroundColor Green

#Execute function to gather information and create the user in O365
Process-User
