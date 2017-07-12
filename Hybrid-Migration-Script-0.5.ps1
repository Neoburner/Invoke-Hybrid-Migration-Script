###########################
# Current version: 0.5
#
# Change Notes:
# 30/05/2017 - Basic Script  Created and Tested
# 31/05/2017 - Added formatting, IF block for RG users
# 22/06/2017 - Added Hosted Voicemail Policy Assignment
# 23/06/2017 - Added voicemail if loop
# 07/07/2017 - Added Online User Validation

###########################
# Requirements:
# Lync Powershell Module
# LyncOnlineConnector Module
# Windows Azure AD Module
# Licences Assigned E3 / CPBX
# 
#
# Set Credential Password (Same Directory): read-host -assecurestring | convertfrom-securestring | out-file cred.txt
#
#

# Global Varibles
$targetFrontend = "XX"
$O365OverrideURL = "XX" # https://technet.microsoft.com/en-us/library/dn689115.aspx / To determine the Hosted Migration Service URL for your Office 365 tenant
$O365Username = "XX"
$hostedVoicemailPolicy = "XX"

# DO NOT EDIT BELOW

# Menu Function

Function ShowMenu {

    Write-Host " "
    Write-Host -ForegroundColor Cyan  "********************************************"
    Write-Host -ForegroundColor Cyan -BackgroundColor Red  "Skype for Business - Hybrid Migration Script"
    Write-Host -ForegroundColor Cyan  "********************************************"
    Write-Host " "
    Write-Host -ForegroundColor Cyan  "1: Migrate Users from Online to On-Premise"
    Write-Host -ForegroundColor Cyan  "2: Enable Users for EV / Apply Policys / Assign Response Groups"
    Write-Host -ForegroundColor Cyan  "3: Migrate Non Response Group Users from On-Premise to Online"
    Write-Host -ForegroundColor Cyan  "4: Apply EV / UM to Online Users"
    Write-Host " "
    Write-Host -ForegroundColor Cyan  "9: User Validation"
    Write-Host -ForegroundColor Cyan  "Q: Quit"
    Write-Host " "

}

# Migrate Users from Online to On-Premise

Function Invoke-MigrateOnlineOnprem () {
    
    Function Get-FileName($initialDirectory)
    {
        [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.initialDirectory = $initialDirectory
        $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
        $OpenFileDialog.ShowDialog() | Out-Null
        $OpenFileDialog.filename
    }
    
    # Input CSV
    Write-Host -ForegroundColor Cyan -BackgroundColor Red "Select CSV..."
    Start-Sleep (2);
    $inputfile = Get-FileName

    # Connect to O365
    Write-Host -ForegroundColor Cyan -BackgroundColor Red "Connecting to O365..."
    Start-Sleep (3);
    
    Import-Module LyncOnlineConnector
    $password = get-content cred.txt | convertto-securestring
    $credential = new-object -typename System.Management.Automation.PSCredential -argumentlist $O365Username,$password
    $session = New-CsOnlineSession -Credential $credential
    Import-PSSession $session -AllowClobber
    
    # Move users from CSV back to on-prem
    
        
                $Users = Import-Csv -Path $inputfile
                ForEach($User in $Users){
                    
                    $SipAddress = $($User.sipaddress)
                    Start-Sleep (3);
                    $Enable = Move-CsUser -Identity $SipAddress -Target $targetFrontend -Credential $credential -HostedMigrationOverrideURL $O365OverrideURL -Confirm:$False
                    $Enable | Out-File Logs\1-OnlinetoOnprem.txt -Append
                    Write-Host -ForegroundColor Cyan -BackgroundColor Red  $SipAddress "Moved to On-Premise"
                    Start-Sleep (2);
                }
               
    Write-Host -ForegroundColor Cyan -BackgroundColor Red  "Closing Sessions..."
    Start-Sleep (2);
    Get-PSSession | Remove-PSSession
    ShowMenu
}


# Enable Users for EV / Apply Policys / Assign Response Groups

Function Invoke-EVEnable () {
    
    # Import Lync Module
    Import-Module Lync

    Function Get-FileName($initialDirectory)
    {
        [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.initialDirectory = $initialDirectory
        $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
        $OpenFileDialog.ShowDialog() | Out-Null
        $OpenFileDialog.filename
    }

    # Input CSV
    Write-Host -ForegroundColor Cyan -BackgroundColor Red  "Select CSV..."
    Start-Sleep (2);
    $inputfile = Get-FileName

    # Apply DDI + Dial Plan + Local Voice Policy + Online Voice Routing Policy

        $Users = Import-Csv -Path $inputfile
        ForEach($User in $Users)
        {
            # Assign Varibles from CSV
            $sipAddress = $($User.sipaddress)
            $lineURI = $($User.lineuri)
            $dialPlan = $($User.dialplan)
            $localVoicePolicy = $($User.localvoiceroutingpolicy)
            $onlineVoicePolicy = $($User.onlinevoiceroutingpolicy)
            $rg = $($User.rg)

            Write-Host -ForegroundColor Cyan -BackgroundColor Red  "Current User:" $SipAddress
            
            Start-Sleep (1);

            Write-Host -ForegroundColor Cyan -BackgroundColor Red  "Enabling EV + Applying LineURI" $lineURI "for user" $sipAddress
            $Enable = Set-CsUser -Identity $sipAddress -EnterpriseVoiceEnabled $True -LineURI "tel:$lineURI"
            Write-Host -ForegroundColor Cyan -BackgroundColor Red  "Setting Dial Plan for user" $sipAddress
            $Enable = Grant-CsDialPlan -Identity $sipAddress -PolicyName $dialPlan
            Write-Host -ForegroundColor Cyan -BackgroundColor Red  "Setting Local Voice Policy for user" $sipAddress
            $Enable = Grant-CsVoicePolicy -Identity $sipAddress -PolicyName $localVoicePolicy
            Write-Host -ForegroundColor Cyan -BackgroundColor Red  "Applying Online Voice Policy:" $onlineVoicePolicy "for user" $sipAddress
            $Enable = Grant-CsVoiceRoutingPolicy -Identity $sipAddress -PolicyName $onlineVoicePolicy
            Write-Host -ForegroundColor Cyan -BackgroundColor Red  "Setting Hosted Voicemail Policy"
            $Enable = Grant-CsHostedVoicemailPolicy -Identity $sipAddress -PolicyName $hostedVoicemailPolicy

            If (!$rg) {
                Write-Host -ForegroundColor Cyan -BackgroundColor Red  "No response group assigned for user" $sipAddress
            }
                    
            Else {
                Write-Host -ForegroundColor Cyan -BackgroundColor Red  "Setting user as member of response group" $rg
                $Enable = Get-CsRgsAgentGroup -Identity service:ApplicationServer:$targetFrontend -Name $rg
                $Enable.AgentsByUri.Add($sipAddress)
                Set-CsRgsAgentGroup -Instance $Enable
                Write-Host -ForegroundColor Cyan -BackgroundColor Red $sipAddress "added to" $rg
                }
                
        }
    
            Write-Host -ForegroundColor Cyan -BackgroundColor Red  "EV Enabled / Policys Applied / RGs set for" $sipAddress
            Write-Host -ForegroundColor Cyan -BackgroundColor Red  "Perform SBC AD Cache Refresh"
            Write-Host -ForegroundColor Cyan -BackgroundColor Red  "Perform AAD Sync"
            Start-Sleep (5);
            ShowMenu
}




# Migrate Non Response Group Users from On-Premise to Online

Function Invoke-MigrateOnpremOnline () {

Function Get-FileName($initialDirectory)
    {
        [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.initialDirectory = $initialDirectory
        $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
        $OpenFileDialog.ShowDialog() | Out-Null
        $OpenFileDialog.filename
    }
    
    # Input CSV
    Write-Host -ForegroundColor Cyan -BackgroundColor Red  "Select CSV..."
    Start-Sleep (2);
    $inputfile = Get-FileName

    # Connect to O365
    Write-Host -ForegroundColor Cyan -BackgroundColor Red  "Connecting to O365..."
    Start-Sleep (3);
    
    Import-Module LyncOnlineConnector
    $password = get-content cred.txt | convertto-securestring
    $credential = new-object -typename System.Management.Automation.PSCredential -argumentlist $O365Username,$password
    $session = New-CsOnlineSession -Credential $credential
    Import-PSSession $session -AllowClobber

    # Move users from CSV to Online
 
    $Users = Import-Csv -Path $inputfile
    ForEach($User in $Users)
        {
            $SipAddress = $($User.sipaddress)
            $rg = $($User.rg)
            #Write-Host -ForegroundColor red  $SipAddress
            #$UPN = $SipAddress.replace("sip:", "")
            
            Start-Sleep (5);
            
            if (!$rg){
                $Enable = Move-CsUser -Identity $SipAddress -Target "sipfed.online.lync.com" -Credential $credential -HostedMigrationOverrideURL $O365OverrideURL -Confirm:$False -Verbose
                $Enable | Out-File Logs\3-OnpremToOnline.txt -Append
                Write-Host -ForegroundColor Cyan -BackgroundColor Red  $SipAddress "Migrated to Online"

            }
            else {
                Write-Host -ForegroundColor Cyan -BackgroundColor Red  $SipAddress "assigned to RG, skipping..."
            }
            
            Start-Sleep (1);
        
    }
                
    Write-Host -ForegroundColor Cyan -BackgroundColor Red  "Closing Sessions..."
    Start-Sleep (2);
    Get-PSSession | Remove-PSSession
    ShowMenu

}

# Apply EV / UM to Online Users

Function Invoke-EnableOnlineEVUser () {

Function Get-FileName($initialDirectory)
    {
        [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.initialDirectory = $initialDirectory
        $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
        $OpenFileDialog.ShowDialog() | Out-Null
        $OpenFileDialog.filename
    }
    
    # Input CSV
    Write-Host " "
    Write-Host -ForegroundColor Cyan -BackgroundColor Red  "Select CSV..."
    Write-Host " "
    Start-Sleep (2);
    $inputfile = Get-FileName

    # Connect to O365
    Write-Host " "
    Write-Host -ForegroundColor Cyan -BackgroundColor Red  "Connecting to O365..."
    Write-Host " "
    Start-Sleep (3);
    
    Import-Module LyncOnlineConnector
    $password = get-content cred.txt | convertto-securestring
    $credential = new-object -typename System.Management.Automation.PSCredential -argumentlist $O365Username,$password
    $session = New-CsOnlineSession -Credential $credential
    Import-PSSession $session -AllowClobber

    # Enable users for EV + UM Online
    $Users = Import-Csv -Path $inputfile
    ForEach($User in $Users)
        {
            $SipAddress = $($User.sipaddress)
            $vm = $($User.vm)
            #Write-Host -ForegroundColor red  $SipAddress
            #$UPN = $SipAddress.replace("sip:", "")
            
            Start-Sleep (2);
            
            $Enable = Set-CsUser -Identity $SipAddress -EnterpriseVoiceEnabled $true     
            Write-Host -ForegroundColor Cyan -BackgroundColor Red  $SipAddress "EV Enabled"

            if ($vm -eq "Y"){
                $Enable = Set-CsUser -Identity $SipAddress -HostedVoicemail $true
                Write-Host -ForegroundColor Cyan -BackgroundColor Red  $SipAddress "Voicemail Enabled"

            }
            else {
                $Enable = Set-CsUser -Identity $SipAddress -HostedVoicemail $false
                Write-Host -ForegroundColor Cyan -BackgroundColor Red $SipAddress "Voicemail Not Required - Skipping"
            }
            
            Start-Sleep (2);
        
        }

    Write-Host -ForegroundColor Cyan -BackgroundColor Red  "Closing Sessions..."
    Start-Sleep (2);
    Get-PSSession | Remove-PSSession
    Write-Host -ForegroundColor Cyan -BackgroundColor Red  "Run UC Migration Script"
    ShowMenu

}

Function Invoke-UserValidation{
    
    Function Get-FileName($initialDirectory)
    {
        [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.initialDirectory = $initialDirectory
        $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
        $OpenFileDialog.ShowDialog() | Out-Null
        $OpenFileDialog.filename
    }
    
    # Input CSV
    Write-Host -ForegroundColor Cyan -BackgroundColor Red "Select CSV..."
    Start-Sleep (2);
    $inputfile = Get-FileName

    # Connect to O365
    Write-Host -ForegroundColor Cyan -BackgroundColor Red "Connecting to O365..."
    Start-Sleep (3);
    
    Import-Module LyncOnlineConnector
    $password = get-content cred.txt | convertto-securestring
    $credential = new-object -typename System.Management.Automation.PSCredential -argumentlist $O365Username,$password
    $session = New-CsOnlineSession -Credential $credential
    Import-PSSession $session -AllowClobber

# Check if user is enabled or exists
    $Users = Import-Csv -Path $inputfile
    ForEach($User in $Users)
        {
            $SipAddress = $($User.sipaddress)           
            $Enable = Get-CsOnlineUser -Identity $SipAddress | Format-Table SipAddress,Enabled
            $Enable | Out-File Logs\9-UserValidation.txt -Append
        }

    Write-Host -ForegroundColor Cyan -BackgroundColor Red  "Closing Sessions..."
    Start-Sleep (2);
    Get-PSSession | Remove-PSSession
    ShowMenu

}


# Execute Menu
do
 {
     ShowMenu
     $selection = Read-Host "Please make a selection"
     switch ($selection)
     {
           '1' {
             Invoke-MigrateOnlineOnprem
         } '2' {
             Invoke-EVEnable
         } '3' {
             Invoke-MigrateOnpremOnline
         } '4' {
             Invoke-EnableOnlineEVUser
         } '9' {
             Invoke-UserValidation
         }
     }
     pause
 }
 until ($selection -eq 'q')