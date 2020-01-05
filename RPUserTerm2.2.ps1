### Roscoe Property Onsite Termination
### 12/23/2019
### MyITpros Ian Eden
### Version 2.2


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
        Write-Host " `n Error connecting to O365...Please check username and password..."
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

#Dialog box to gather termination info. 
function Prompt 
{
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Roscoe Termination'
    $form.Size = New-Object System.Drawing.Size(400,250)
    $form.StartPosition = 'CenterScreen'
    
    $SubmitButton = New-Object System.Windows.Forms.Button
    $SubmitButton.Location = New-Object System.Drawing.Point(200,150)
    $SubmitButton.Size = New-Object System.Drawing.Size(75,23)
    $SubmitButton.Text = 'Submit'
    $SubmitButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $SubmitButton
    $form.Controls.Add($SubmitButton)
    
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(100,150)
    $CancelButton.Size = New-Object System.Drawing.Size(75,23)
    $CancelButton.Text = 'Cancel'
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)

    $TermUser = New-Object System.Windows.Forms.Label
    $TermUser.Location = New-Object System.Drawing.Point(60,20)
    $TermUser.Size = New-Object System.Drawing.Size(280,20)
    $TermUser.ForeColor = "DarkRed"
    $TermUser.Text = "Please enter the terminated user's email address:"
    $form.Controls.Add($TermUser)
    
    $TermUsertextBox = New-Object System.Windows.Forms.TextBox
    $TermUsertextBox.Location = New-Object System.Drawing.Point(60,40)
    $TermUsertextBox.Size = New-Object System.Drawing.Size(260,20)
    $form.Controls.Add($TermUsertextBox)

    $AutoReply = New-Object System.Windows.Forms.Label
    $AutoReply.Location = New-Object System.Drawing.Point(60,80)
    $AutoReply.Size = New-Object System.Drawing.Size(280,20)
    $AutoReply.Text = "Please enter the Auto Reply email address:"
    $form.Controls.Add($AutoReply)
    
    $AutoReplytextBox = New-Object System.Windows.Forms.TextBox
    $AutoReplytextBox.Location = New-Object System.Drawing.Point(60,100)
    $AutoReplytextBox.Size = New-Object System.Drawing.Size(260,20)
    $form.Controls.Add($AutoReplytextBox)
    
    $form.Topmost = $true
    
    $form.Add_Shown({$TermUsertextBox.Select()})
    $form.Add_Shown({$AutoReplytextBox.Select()})
    $result = $form.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        $script:TermUserEmail = $TermUsertextBox.Text
        $script:AutoReplyEmail = $AutoReplytextBox.Text
    }
    elseif ($result -eq [System.Windows.Forms.DialogResult]::Cancel) {
        Write-Host `n"Goodbye!"
        Start-Sleep -Seconds 1
        Exit
    }
   
}


function ProcessTerm {
        #Calling function for variables.
    prompt

    #Password for term'd users (Modified for github)
    $password = (ConvertTo-SecureString -AsPlainText "******" -Force)
    
     if ($Mailboxes -contains $TermUserEmail) {
                 
        try {
            Get-Mailbox $TermUserEmail -ErrorAction Stop | Format-Table
            Write-Host "Successfully located mailbox" -ForegroundColor Green
            }
        catch {
            Write-Host "`n Could not find $TermUserEmail in O365. Please try again" -ForegroundColor Yellow
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Write-Host `n"$ErrorMessage"
            Write-Host `n"$FailedItem"
            Return ProcessTerm      
        }
        #Resetting password
        try {
            Set-AzureADUserPassword -ObjectId  "$TermUserEmail" -Password $password -EA Continue
            Write-Host `n"Password reset successfully" -ForegroundColor Green 
            }
         Catch {
             Write-Host `n"Error retting password" -ForegroundColor Red
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Write-Host `n"$ErrorMessage"
            Write-Host `n"$FailedItem"
            }

    
        #Blocking Sign-in
        try { 
            Set-AzureADUser -ObjectID $TermUserEmail -AccountEnabled $false -ErrorAction Continue
            Write-Host `n"Successfully blocked sign-in" -ForegroundColor Green
        }
        catch {
            Write-Host `n"Could not block user sign-in" -ForegroundColor Red
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Write-Host `n"$ErrorMessage"
            Write-Host `n"$FailedItem"

        }



        #Find all of the Office 365 Groups
        $O365Groups = Get-AzureADMSGroup -All:$true 
        #Who replaces the owner if the terminated user is the only owner.
        $OwnerSub = Get-AzureADUser -SearchString "myitpros@rpmliving.com"
        $ADTermuser = Get-AzureADUser -SearchString "$TermUserEmail"

        Write-Host "`n Preparing to remove user from groups...please wait" -ForegroundColor Yellow

        # Check each group for the user
        try {
            foreach ($group in $O365Groups) {
                $members = (Get-AzureADGroupMember -ObjectID $group.id).UserPrincipalName
                If ($members -contains $ADTermuser.UserPrincipalName) {
                    Remove-AzureADGroupMember -ObjectId $group.Id -MemberId $ADTermuser.ObjectId
                    Write-Host "`n Removed from $($group.DisplayName)" -ForegroundColor Green
                    $owners = Get-AzureADGroupOwner -ObjectId $group.Id
                    foreach ($owner in $owners) {
                        If ($ADTermuser.UserPrincipalName -eq $owner.UserPrincipalName) {
                            # If needed, add new owner to prevent orphaned group
                            If ($owner.count -lt 2){
                                Write-Host "$($OwnerSub.UserPrincipalName) was added as a new owner" -ForegroundColor Green
                                Add-AzureADGroupOwner -ObjectId $group.Id -RefObjectId $OwnerSub.ObjectId
                            }
            
                                # Remove the user as owner
                                Write-Host "Removed as owner of $($group.DisplayName)" -ForegroundColor Yellow
                                Remove-AzureADGroupOwner -ObjectId $group.Id -OwnerId $ADTermuser.ObjectId  
                        }  
                    }
                }   
            }
        }
        catch {
            Write-Host `n"No groups were removed" -ForegroundColor Green
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Write-Host `n"$ErrorMessage"
            Write-Host `n"$FailedItem"
        }

        #Hide from GAL
        try { 
            Set-Mailbox -Identity $TermUserEmail -HiddenFromAddressListsEnabled $true -ErrorAction Continue
            Write-Host `n"Hidden from the Global Address List" -ForegroundColor Green
        }
        catch {
            Write-Host `n"Unable to hide user from Global Address List" -ForegroundColor Red
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Write-Host `n"$ErrorMessage"
            Write-Host `n"$FailedItem"
        }

        #Set Automatic replies
        try {
            Set-MailboxAutoReplyConfiguration -Identity $TermUserEmail -AutoReplyState Enabled -InternalMessage "You have reached an account that is no longer active. Please send all requests to $AutoReplyEmail" -ExternalMessage "You have reached an account that is no longer active. Please send all requests to $AutoReplyEmail" -ErrorAction Continue

            Write-Host "`n The autoreply message has been enabled for internal and external replies with the follow response:  "
            Write-host "`n  You have reached an account that is no longer active. Please send all requests to $AutoReplyEmail" -ForegroundColor Gray
        }
        catch {
            Write-Host `n"Unable to set Automatic Replies" -ForegroundColor Red
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Write-Host `n"$ErrorMessage"
            Write-Host `n"$FailedItem"
        }

        #Clear department field
        try {
            Set-AzureADUser -ObjectID $TermUserEmail -Department $null -ErrorAction Continue
            Write-Host `n"Cleared Department Field" -ForegroundColor Green
        }
        catch {
            Write-Host `n"Unable to clear Department Field" -ForegroundColor Red
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Write-Host `n"$ErrorMessage"
            Write-Host `n"$FailedItem"
        }
        
        #Repeat process.
        Write-Host `n"Script complete. Please enter another user or select 'Cancel' to exit" -ForegroundColor Yellow 

        Return ProcessTerm    

    }
    Elseif ($Mailboxes -notcontains $TermUserEmail) {
        Write-Host "`n The email $TermUserEmail does not exist. Please verify in O365." -ForegroundColor Red
        Start-Sleep -Seconds 1  
        Return ProcessTerm  

    }
}








#Executing script.

Connect-O365

Write-Host "Loading mailboxes...please wait..." -ForegroundColor Yellow
$script:Mailboxes = Get-Mailbox | select -expand primarysmtpaddress
Write-Host "Complete!" -ForegroundColor Green
Clear-Host

Write-Warning "This script will perform the following actions:
`n
`r Reset Password
`r Block sign-in
`r Remove user from all groups
`r Hide user from the Global Address List
`r Set Auto-Reply"

ProcessTerm

