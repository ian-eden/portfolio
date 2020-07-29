### Author: Ian Eden 
### 7/20/2020



#Connecting to required O365 PowerShell
function connect-exo {
    try {

         
        Write-Host `n"Attempting to connecto to Exchange Online PowerShell...Please sign in"
        Start-Sleep 2
        Connect-ExchangeOnline -ErrorAction Stop
        Clear
        Write-Host `n"Attempting to connect to SharePoint Online PowerShell...Please sign in"
        Connect-SPOService -Url  -ErrorAction Stop 
        clear
        Write-Host `n"Attempting to connect to Azure AD PowerShell...Please sign in"
        Connect-AzureAD -ErrorAction Stop
    }

    catch {
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
        Write-Host `r"Error signing in, please check your credentials and try again" -ForegroundColor DarkRed
        Write-Host `n"$ErrorMessage"
        Write-Host `n"$FailedItem"
        Return connect-exo
    } 
}



function PROCESS-REPORT {

    Clear
    Write-Warning `n" This may take a very long time to run. DO NOT CLOSE THIS WINDOW "
    Write-Host `n"Please enter your username" -ForegroundColor Yellow
    Write-Host "Example: ianede01" -ForegroundColor Yellow
    $script:savepath = Read-Host
    Write-Host `n"Report will begin running in 15 seconds. Press ctrl+c to cancel" -ForegroundColor Green
    Start-Sleep 10
    Write-Host "5 Seconds remaining" -ForegroundColor Yellow
    Start-Sleep 5
    Write-Host `n"Starting report... Do not close this window"

    #Function to gather Blocked user accounts and SPO Owner Groups
    function get-SPGroups_Blocked {
        try {

            Write-Host `n"Step 1: Gathering disabled account..." -ForegroundColor Green
            $script:BlockedAccounts = Get-AzureADUser -Filter "AccountEnabled eq false" -ErrorAction Stop
            
                
        }
        catch {
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Write-Host `r"Encounted error in Step 1" -ForegroundColor DarkRed
            Write-Host `n"$ErrorMessage"
            Write-Host `n"$FailedItem"
            
        } 

    }

    ### Function to get SPO only Groups
    function SPO-GROUPS {

        try{
            Write-Host `n"Step 2: Gathering SharePoint Sites...This will probably take awhile..." -ForegroundColor Green
            #Getting all SPO Sites 
            $script:spogroups = Get-SPOSite -Limit All -ErrorAction Stop  
            $script:spogroups | Select Url | Out-File -FilePath C:\Users\$script:savepath\desktop\SPOSites.txt
           
        }

        catch {
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Write-Host `r"Encounted error in Step 2" -ForegroundColor DarkRed
            Write-Host `n"$ErrorMessage"
            Write-Host `n"$FailedItem"
            
        } 

        try{
            $ErrorActionPreference = "SilentlyContinue"
            Write-Host `n"Step 3: Processing SharePoint Sites...This will probably take awhile..." -ForegroundColor Green
            #Setting Variable to empty
            $script:spogroupids=@()

            #foreach loop to grab each GroupId, converto to string and then add to $spogroupids array
            foreach ($script:spogroup in $script:spogroups){
  
                  $script:GroupProperties = Get-SPOSite -Identity $spogroup | Select GroupId -ErrorAction SilentlyContinue
                  $script:GroupObject = $script:GroupProperties.GroupId.ToString()
                  $script:spogroupids +=  $script:GroupObject  | Where GroupId -NotContains "00000000-0000-0000-0000-000000000000" -ErrorAction SilentlyContinue 
             } 

          }
        catch {
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Write-Host `r"Encounted error in Step 3.0" -ForegroundColor DarkRed
            Write-Host `n"$ErrorMessage"
            Write-Host `n"$FailedItem"
            
        }

        try {
            $ErrorActionPreference = "SilentlyContinue"
            #foreach loop to get corrisponding O365 group that corrisponds with SPO Site GroupId 
            $script:O365Groups=@()
            foreach ($script:spogroupid in $script:spogroupids) {

                 $script:O365Properties = Get-UnifiedGroup -Filter "ExternalDirectoryObjectId -eq '$script:spogroupid'" 
                 $script:O365Object = $script:O365Properties
                 $script:O365Groups += $script:O365Object

            }
            
            $script:O365Groups | Select DisplayName, PrimarySmtpAddress | Out-File -FilePath C:\Users\$script:savepath\desktop\O365Groups.txt
            $script:spomangedby = $script:O365Groups | Select Managedby

        }
        catch {
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Write-Host `r"Encounted error in Step 3.1" -ForegroundColor DarkRed
            Write-Host `n"$ErrorMessage"
            Write-Host `n"$FailedItem"
            
        }

        
    }

    function SPO-OWNERS {
        $ErrorActionPreference = "Continue"
        try {
            Write-Host `n"Step 4: Processing SharePoint Group Owners" -ForegroundColor Green
            
            #Pulls Mangedby user accounts out of variable and runs Get AzureADuser
            $script:Sorted_SPOOwners = foreach ($script:SPOowner in $script:spomangedby.managedby){

                Get-AzureADUser -SearchString $script:SPOowner | Select UserPrincipalName

            }

            #Deduplicates SP Owners
            $script:UniqueSPowners = $script:Sorted_SPOOwners | Sort UserPrincipalName | Get-Unique -AsString

        }

        catch {
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Write-Host `r"Encounted error in Step 4.0" -ForegroundColor DarkRed
            Write-Host `n"$ErrorMessage"
            Write-Host `n"$FailedItem"
            
        }

    }

    #Function to Compare and display blocked users and groups that they are assigned to.
    function LOCATE-BLOCKED-USERS {
        
        $ErrorActionPreference = "Continue"

        try {

            Write-Host `n"Final Step 5: Locating disabled users with SharePoint Owner Permissions" -ForegroundColor Green
        
            #Compares list of blocked user accounts with group owners
            $script:BlockedUsers_Found = Compare-Object $script:BlockedAccounts $script:UniqueSPowners -Property UserPrincipalName -IncludeEqual -ExcludeDifferent | Select UserPrincipalName
            $script:BlockedUsers_Found | Out-File -FilePath C:\Users\$script:savepath\desktop\Blockedusers.txt

            $script:stringblockedusers=@()

            #Converting matched blocked users to strings and striping the @email.com from their account and then adding them to an array $stringblockedusers
            foreach ($script:BlockedUser_Found in $script:BlockedUsers_Found){

                $script:userProperties = $script:BlockedUser_Found.UserPrincipalName.ToString()
                $script:userObject = $script:userProperties.Split("@")[0]
                $script:stringblockedusers +=  $script:userObject
             }


            #Matching Listing blocked users and their Sites
            foreach ($script:stringblockeduser in $script:stringblockedusers) {

                $script:O365Groups | select DisplayName, ManagedBy | Where Managedby -Contains $script:stringblockeduser | FT Displayname, $stringblockeduser | Out-File -Append c:\users\$script:savepath\desktop\blocked_Users_Sites.csv
       
            }

       
       }

        catch {
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            Write-Host `r"Encounted error in Step 5.0" -ForegroundColor DarkRed
            Write-Host `n"$ErrorMessage"
            Write-Host `n"$FailedItem"
            
        }

    }

get-SPGroups_Blocked
SPO-GROUPS
SPO-OWNERS
LOCATE-BLOCKED-USERS
Write-Host `n"Multiple reports saved to this location 'c:\users\$script:savepath\desktop'" -ForegroundColor Yellow

}


connect-exo  
PROCESS-REPORT

Write-Host `n"End of script...You may now close this window" -ForegroundColor Yellow
Start-Sleep 60


