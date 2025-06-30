 <#
 .SYNOPSIS
    Enable users for Teams Shared Calling
 .DESCRIPTION
    This script performs bulk user enablement for Teams Shared Calling.
    
    -----------------------------------------------------------------------------------------------------------------------------------
	Script name : sharedCalling_enablement_v1.1.ps1
	Authors : Microsoft ISD
	Version : V1.1.0
	Dependencies :
                    Microsoft.Graph.Beta.Teams - v2.25.0
                    MicrosoftTeams - v6.4.0
	-----------------------------------------------------------------------------------------------------------------------------------
	-----------------------------------------------------------------------------------------------------------------------------------
	Version Changes:
	Date:           Version:        Changed By:         Info:
	02/06/2025      V1.1.0          Microsoft ISD       Removed all non required components and functions to only perform functions
                                                        related to the Teams Shared Calling user enablement.
                                                                                                     
	-----------------------------------------------------------------------------------------------------------------------------------
	DISCLAIMER
	THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED 'AS IS' WITHOUT WARRANTY OF ANY KIND.
	MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES
	OF MERCHANTABILITY OR OF FITNESS FOR A PARTICULAR PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR
	PERFORMANCE OF THE SAMPLES REMAINS WITH YOU. IN NO EVENT SHALL MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR
	ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS, BUSINESS
	INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR
	INABILITY TO USE THE SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES.
	BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION OF LIABILITY FOR CONSEQUENTIAL OR
	INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.
	-----------------------------------------------------------------------------------------------------------------------------------

.SCRIPT REQUIREMENTS
    Module Microsoft.Graph.Beta.Teams
    Module MicrosoftTeams
    
 .PARAMETER
    Global settings must be configured in the config.xml file

 .EXAMPLE
    .\sharedCalling_enablement_v1.1.0.ps1
 #>

$path = "$PSScriptRoot"
[string]$Title = "Teams Shared Calling Bulk User Enablement Tools"
[string]$Version = "1.1.0"
[int]$SplashDelay = 5 # Default delay
[string]$SplashForegroundColor = "Cyan"  # Default foreground color


 <# 
    Config Parameters 

    XML config file elements for each Teams policy has two options:
    1. GroupName - This is the name of the group that will be used to assign the policy to users if using Group based policy assignment.
    2. PolicyName - This is the name of the policy that will be assigned to the group when using Teams Policy based assignment.
    Only one of these options can be used at a time.
    If both have values, the script will use the PolicyName option.
#>
[xml]$Global:config = Get-Content "$path\config.xml"
[String]$Global:entraid = $Global:config.settings.EntraID
[String]$Global:spn = $Global:config.Settings.ServicePlanName
[String]$Global:SendChat = $Global:config.Settings.SendChat
[String]$Global:ChatImage = $Global:config.Settings.ChatImage

# Teams Policies
# Enterprise Voice
[String]$Global:tdp = $Global:config.Settings.TenantDialPlan.PolicyName
[String]$Global:tdpg = $Global:config.Settings.TenantDialPlan.GroupName
[String]$Global:ovrp = $Global:config.Settings.OnlineVoiceRoutingPolicy.PolicyName
[String]$Global:ovrpg = $Global:config.Settings.OnlineVoiceRoutingPolicy.GroupName
[String]$Global:cli = $Global:config.Settings.CallingLineIdentity.PolicyName
[String]$Global:clig = $Global:config.Settings.CallingLineIdentity.GroupName

# Calling Policies
[String]$Global:tcpu = $Global:config.Settings.TeamsCallingPolicyUsers.PolicyName
[String]$Global:tcpug = $Global:config.Settings.TeamsCallingPolicyUsers.GroupName

# Emergency Calling and Routing Policies
[String]$Global:ecp = $Global:config.Settings.TeamsEmergencyCallingPolicy.PolicyName
[String]$Global:ecpg = $Global:config.Settings.TeamsEmergencyCallingPolicy.GroupName
[String]$Global:ecrp = $Global:config.Settings.TeamsEmergencyCallRoutingPolicy.PolicyName
[String]$Global:ecrpg = $Global:config.Settings.TeamsEmergencyCallRoutingPolicy.GroupName

#Other Policies
[string]$Global:tscp = $Global:config.Settings.TeamsSharedCallingPolicy.PolicyName
[string]$Global:tscpg = $Global:config.Settings.TeamsSharedCallingPolicy.GroupName

# Teams AA CQ Policies
[String]$Global:tvap = $Global:config.Settings.TeamsVoiceApplicationPolicy.PolicyName
[String]$Global:tvapg = $Global:config.Settings.TeamsVoiceApplicationPolicy.GroupName

# Teams Chat Messages - Defaults
[String]$Global:MigrationSuccessMessage = $Global:config.Settings.MigrationSuccessMessage
[String]$Global:MigrationFailMessage = $Global:config.Settings.MigrationFailMessage

# Initialize the global variable to ensure it starts at 0
$Global:UsersInBatch = 0




###########################
#                         #
#        Functions        #
#                         #
########################### 



<#

    Function: Writes a splash screen to the console. When the script is run, it will display a splash screen with the specified title and version.
    The Splash screen will be displayed for a period of 5 seconds before loading the menu script.
    Parameters:
        $Title: The title of the splash screen.
        $Version: The version of the script.
        $Delay: The delay in seconds before the menu script is loaded.
#>

function Show-SplashScreen {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Title,

        [Parameter(Mandatory=$true)]
        [string]$Version,

        [Parameter(Mandatory=$false)]
        [int]$Delay = 5,

        [Parameter(Mandatory=$false)]
        [string]$ForegroundColor = "DarkMagenta"  
    )
    

    # Clear the console
    Clear-Host
    # Set the console window size
    $host.ui.RawUI.WindowSize = New-Object System.Management.Automation.Host.Size(120, 60)
    # Set the console window title
    $host.ui.RawUI.WindowTitle = $Title 


    Write-host -Object ""

    $AsciiArt = "                                                                                                  
                                       777777                                                                          
                                   7125555555537                                                                       
                                 77555555555555527                                                                     
                                 155555555555555557       777777                                                       
                                7555555555555555552     7466666647                                                     
                                7555555555555555552    166666666662                                                    
                                7555555555555555552   76666666666667                 73255555555555217                 
          12222222222222222222222544446444555555557   76666666666667              12555552311113255555537              
         24444444666666666666666666666666455555557     3666666666657           715555377          771555537            
         544466666666666666666666666666664555521       714666666427      771113555517                 7255517          
         566666666666666666666666666666433317             713337     72555555555517        735555537    755537         
         5666666647777777777777466666664                               3555377            3555555527     72553         
         566666664111117  1111166666666664555555555555556666666666641   755527          7255555551        75557        
         566666666666667  6666666666666664555555555555554666666666665    725557        725557  7           15557       
         566666666666667  6666666666666664555555555555554666666666665     72553       73552                75557       
         566666666666667  6666666666666664555555555555554666666666665     72551       75557                75551       
         566666666666667  6666666666666664555555555555554666666666665      2551       15557                75557       
         566666666666667  6666666666666664555555555555554666666666665      3553       35557                15557       
         566666666666667  6666666666666664555555555555554666666666665      75557      3555552             75557        
         566666666666667  6666666666666964555555555555554666666666665       15551     7555552            72553         
         566666666666666666666666666999964555555555555554666666666665        155517    725553           755537         
         566666666666666666666669999999964555555555555554666666666662         7555517     7           7255517          
         566666666666666666669999999999964555555555555554666666666661          715555377          771555527            
         146666666666666669999999999999964555555555555554666666666627             12555552311111255555537              
                          29999999999996645555555555555566666666643                  72555555555555217                 
                          744444444444444555555555555556666666647                                                      
                           75555555555555555555555555517777777                                                         
                            72555555555555555555555527                                                                 
                             77555555555555555555551                                                                   
                               7735555555555555537                                                                     
                                   771322223317"


    Write-Host -object $AsciiArt -ForegroundColor $ForegroundColor
    $centeredTitle = $Title.PadLeft(([math]::Floor((120 - $Title.Length) / 2)) + $Title.Length).PadRight(120)
    $centeredVersion = ("Version: $Version").PadLeft(([math]::Floor((120 - ("Version: $Version").Length) / 2)) + ("Version: $Version").Length).PadRight(120)
    $centeredLoading = "Loading...".PadLeft(([math]::Floor((120 - "Loading...".Length) / 2)) + "Loading...".Length).PadRight(120)

    Write-host -Object ""
    Write-host -Object ""
    Write-Host -object $centeredTitle -ForegroundColor $ForegroundColor
    Write-host -Object ""
    Write-Host -object $centeredVersion -ForegroundColor $ForegroundColor
    Write-host -Object ""
    Write-Host -object $centeredLoading -ForegroundColor $ForegroundColor
    Write-host -Object ""
    Write-host -Object ""
    Write-host -Object ""
    Write-host -Object ""
    Write-host -Object ""
    Write-host -Object ""
    Write-host -Object ""
    Write-host -Object ""
    Write-host -Object ""

    Write-Host -Object "DISCLAIMER
	THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED 'AS IS' WITHOUT WARRANTY OF ANY KIND.
	MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES
	OF MERCHANTABILITY OR OF FITNESS FOR A PARTICULAR PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR
	PERFORMANCE OF THE SAMPLES REMAINS WITH YOU. IN NO EVENT SHALL MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR
	ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS, BUSINESS
	INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR
	INABILITY TO USE THE SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES.
	BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION OF LIABILITY FOR CONSEQUENTIAL OR
	INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU." -ForegroundColor DarkGray

    # Wait for the specified delay
    Start-Sleep -Seconds $Delay

    Write-Host ""
    $continue = Read-Host "Press [C] to continue"
    If ($continue.ToLower() -eq "c") {
        Return $true
    }

    # Clear Host
    Clear-Host

}



<# 
    Auth2
    This function will check if the user is already authenticated to Microsoft Teams and Graph.
    If not, it will prompt the user to authenticate.
    It will also import the required modules for the script to run.
    Ensure all modules are installed before running the script.
    e.g. Install-Module -Name Microsoft.Graph -Scope CurrentUser -AllowClobber -Force
    Additionally if already installed, use the Fix-Graph.ps1 script
#>


Function Auth2 {
    $Modules = Get-Module -ListAvailable MicrosoftTeams, Microsoft.Graph.Teams, Microsoft.Graph.Authentication
        $scopes = [string[]]@(
                'User.Read.All'
                'User.ReadWrite.All'
                'Chat.ReadWrite'
                'ChatMessage.Send'
                'Chat.Create'
                'GroupMember.ReadWrite.All'
                'TeamworkDevice.ReadWrite.All'
            )
    
    Try {
        
        # Option Microsoft Graph - Graph
        
            Write-Host "Connecting to Microsoft Graph, provide credential when prompted..."
            Connect-MgGraph -ContextScope Process -Scopes $scopes -NoWelcome | Out-Null

        
        # Option Microsoft Teams - Teams
          
            Write-Host "Connecting to Microsoft Teams, provide credential when prompted..."
            Connect-MicrosoftTeams | Out-Null
        
        $InitialState = [InitialSessionState]::CreateDefault2()
        $Modules.ForEach({ 
                Write-Verbose "Importing module $($_.Name)"
                $InitialState.ImportPSModule($_)
            })
    }    
    Catch {
        ('Error - {0}' -f $_.Exception.Message) | Out-File "$path\logs\auth_err.txt" -append
        Clear-Host
        Exit
    }
    Finally {
        Clear-Host
    }
}


<# 
    Initialize
#>
Auth2



Function Validate-BatchFile {
    
    $targetusers = Import-Csv "$batchfile" -ErrorAction SilentlyContinue
    CLS
    if (-not $targetusers -or $targetusers.Count -eq 0) {
        Write-Warning "No Batch file has been selected or the selected batch file is empty or invalid. 
        Please run 'Option B' and select a batch file which contains users to be enabled"
        
        Write-Host -Object ''
		
        $continue = Read-Host "Press [C] to continue"
        If ($continue.ToLower() -eq "c")
		{
			Return $false
		}
    }

    return $true
}
<#
    Is Group Based Policy Assignment
    check if the policy to be assigned is via group based policy assignment or not from the config.xml file
    if the policy is group based, return true, else return false
    if both group and policy name are present, return true
#>
Function isGroupBasedPolicyAssignment ($policyName) {

    $policy = $Global:config.settings.$policyName
    If ($policy.GroupName -eq $null -or $policy.GroupName -eq "") {
        return $false
    } elseif ($policy.GroupName -ne $null -and $policy.PolicyName -ne $null) {
        return $true
    } else {
        return $false
    }   
}



<#
    Send-Chat-Message
    This function will send a chat message to the user with the specified UPN.
    The message is sent from the user running the script.
    The message will include the migration status and the new Teams number assigned.
    The message is sent as Important
    If the chat fails, run the Fix-Graph.ps1 script
#>

Function Send-Chat2 {
    param(
        [string]
        $To,
    
        [string]
        $MigrationStatus,

        [string]
        $PhoneNumber,

        [string]
        $Extn
    )
    # Get details of the signed-in user
    $SendingUser = (Get-MgContext).Account
    $SendingUserId = (Get-MgUser -UserId $SendingUser).Id
    #Write-Host ("Chats will be sent by {0}" -f $SendingUser)

    $WebImage = $Global:ChatImage
    # Download the icon we want to use if it's not already available - use your own image if you want
    $ContentFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\ChatImage.jpg"
    $UrlCacheFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\ChatImageUrl.txt"
    
    # Check if we need to refresh the image (if URL changed or file doesn't exist)
    $refreshImage = $true
    if (Test-Path -Path $ContentFile -PathType Leaf) {
        if (Test-Path -Path $UrlCacheFile -PathType Leaf) {
            $cachedUrl = Get-Content -Path $UrlCacheFile -Raw
            if ($cachedUrl -eq $WebImage) {
                $refreshImage = $false
            }
        }
    }
    
    # Download new image if needed
    if ($refreshImage) {
        # Remove old files if they exist
        if (Test-Path -Path $ContentFile -PathType Leaf) {
            Remove-Item -Path $ContentFile -Force
        }
        # Download new image
        Invoke-WebRequest $WebImage -OutFile $ContentFile
        # Save the URL to cache file
        Set-Content -Path $UrlCacheFile -Value $WebImage -Force
    }

    # Get image dimensions (width and height) of the downloaded file
    Add-Type -AssemblyName System.Drawing
    $image = [System.Drawing.Image]::FromFile($ContentFile)
    $width = $image.Width
    $height = $image.Height
    $image.Dispose()

    # Define the content of the chat message, starting with the inline image
    $Content = '<img height="' + $height + '" src="../hostedContents/1/$value" width="' + $width + '" style="vertical-align:bottom; width:' + $width + 'px; height:' + $height + 'px">'
    $Content = $Content + '<p><b>Teams Shared Calling Enablement Status</b></p>'
    $Content = $Content + "<p>$MigrationStatus</p>"
    if (-not [string]::IsNullOrEmpty($PhoneNumber)) {
        $Content = $Content + "<p>Your Teams Shared Calling Telephone Number is: <b>$PhoneNumber</b></p>"
    }
    If (-not [string]::IsNullorEmpty($Extn)){
        $Content = $Content + "<p>Your Extension Numbr is: <b>$Extn</b></p>"
    }

    # Create a hash table to hold the image content that's used with the HostedContents parameter
    $ContentDataDetails = @{}
    $ContentDataDetails.Add("@microsoft.graph.temporaryId", "1")
    $ContentDataDetails.Add("contentBytes", [System.IO.File]::ReadAllBytes("$ContentFile"))
    $ContentDataDetails.Add("contentType", "image/jpeg")
    [array]$ContentData = $ContentDataDetails

    # Define the body of the chat message
    $Body = @{}
    $Body.add("content", $Content)
    $Body.add("contentType", 'html')

    # Get the user details for the recipient
    $User = Get-MgUser -Filter "userPrincipalName eq '$To'" -Property id, displayName, userprincipalName, userType
    If (-not $User) {
        echo ('Timestamp: {0}' -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) ('User Not Found: {0}' -f $To) | Out-File "$path\logs\Send_Chat_debugs.txt" -Append
        return
    }

    # No need to chat with the sender, so ignore if the recipient is the sender
    If ($User.Id -eq $SendingUserId) {
        echo ('Timestamp: {0}' -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) ('Skipping sending chat to self: {0}' -f $To) | Out-File "$path\logs\Send_Chat_debugs.txt" -Append
        return
    }

    [array]$MemberstoAdd = $SendingUserId, $User.Id
    [array]$Members = $null
    ForEach ($Member in $MemberstoAdd) {
        $MemberId = ("https://graph.microsoft.com/v1.0/users('{0}')" -f $Member)
        $MemberDetails = @{}
        [array]$MemberRole = "owner"
        If ($User.userType -eq "Guest") {
            [array]$MemberRole = "guest"
        }
        $MemberDetails.Add("roles", $MemberRole.trim())
        $MemberDetails.Add("@odata.type", "#microsoft.graph.aadUserConversationMember")
        $MemberDetails.Add("user@odata.bind", $MemberId.trim())
        $Members += $MemberDetails
    }

   # Add the members to the chat body
   $OneOnOneChatBody = @{}
   $OneOnOneChatBody.Add("chattype", "oneOnOne")
   $OneOnOneChatBody.Add("members", $Members)

   # Debugging - Uncomment to see the chat body details
   # write-host -Object "Chat Body: $($OneOnOneChatBody | ConvertTo-Json -Depth 5)"


    # Set up the chat - if one already exists between these two participants, Teams returns the id for that chat
    Try {
        $NewChat = New-MgChat -BodyParameter $OneOnOneChatBody
    } Catch {
        echo ('Timestamp: {0}' -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) ('Error: Failed to create chat - {0}' -f $_.Exception.Message) | Out-File "$path\logs\Send_Chat_debugs.txt" -Append
        return
    }
    If ($NewChat) {
        #Write-Host ("Chat {0} available" -f $NewChat.id)
        echo ('Timestamp: {0}' -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) ('Chat body {0} created' -f $NewChat.id -f $User.DisplayName) | Out-File "$path\exp\Send_Chat_Log.txt" -Append
    } Else {
        echo ('Timestamp: {0}' -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) ('Failed to create chat' -f $User.DisplayName) | Out-File "$path\logs\Send_Chat_debugs.txt" -Append
        return
    }

    # Send the message to the chat
    echo ('Timestamp: {0}' -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) ('Sending chat to {0}' -f $User.DisplayName) | Out-File "$path\exp\Send_Chat_Log.txt" -Append
    $ChatMessage = New-MgChatMessage -ChatId $NewChat.Id -Body $Body `
        -HostedContents $ContentData -Importance High
    If ($ChatMessage) {
        echo ('Timestamp: {0}' -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) ('Chat sent to {0}' -f $User.DisplayName) | Out-File "$path\exp\Send_Chat_Log.txt" -Append
    } Else {
        echo ('Timestamp: {0}' -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) ('Failed to send chat message to {0}' -f $User.DisplayName) | Out-File "$path\logs\Send_Chat_debugs.txt" -Append
    }
}



<# 
    Option-A
    Archive Files
#>
Function Archive-Files {
    Try {
        Get-ChildItem -Path "$path\exp\*", "$path\logs\*" | Compress-Archive -DestinationPath "$path\arc\$($batchid)_files_$((Get-Date).ToString('dd-MM-yyyy')).zip" -Force -ErrorAction Stop

    }
    catch {
        ('Error: Archive Files - {0}' -f $_.Exception.Message) | Out-File "$path\logs\Archive_Files_debugs.txt" -Append
    }
}

Function Clear-Cache {
Remove-Item "$path\exp\*", "$path\logs\*", "$path\tmp\*"
}


<# 
    Option-B
    Batch File Import
    Select CSV file when using the Import-Csv
    This will only allow csv files to be selected.
#>
Function Select-csvFile
{
  
	param([string]$Title="Please select the batch file",[string]$Directory=$("$path\data\"),[string]$Filter='Comma Separated Values File (.csv) |*.csv')
	[System.Reflection.Assembly]::LoadWithPartialName("PresentationFramework") | Out-Null
	$objForm = New-Object Microsoft.Win32.OpenFileDialog
	$objForm.InitialDirectory = $Directory
	$objForm.Filter = $Filter
	$objForm.Title = $Title
	$show = $objForm.ShowDialog()

	If ($show -eq $true)
	{
		#Get file path
        [String]$csvFile = $objForm.FileName

		#Check file
        Write-Warning "The following file was selected: $($csvFile.Split("\")[$_.Length-2])"
        $targetusers = Import-Csv -Path $csvFile
        Write-Host -Object "There are $($targetusers.Count) users in the selected $($csvFile.Split('\')[-1]) file"

        $Global:UsersInBatch = $targetusers.Count
		$continue = Read-Host "Press [C] to continue"

		If ($continue.ToLower() -eq "c")
		{
			Return $csvFile
		}
    }
    
    Return 0
}



<#
    Option-C
    Show-Config
    This function will show the current config that is loaded into the script

#>

Function Show-Config {
    Cls
    # Initialize the table with a header row
    $configTable = @()
    $configTable += [PSCustomObject]@{ Name = 'Attribute Name'; Value = 'Value' }

    $configTable += [PSCustomObject]@{ Name = 'EntraID'; Value = $Global:entraid }
    $configTable += [PSCustomObject]@{ Name = 'ServicePlanName'; Value = $Global:spn }
    $configTable += [PSCustomObject]@{ Name = 'SendChat'; Value = $Global:SendChat }
    $configTable += [PSCustomObject]@{ Name = 'ChatImage'; Value = $Global:ChatImage }

    $configTable += [PSCustomObject]@{ Name = 'TenantDialPlan.PolicyName'; Value = $Global:tdp }
    $configTable += [PSCustomObject]@{ Name = 'TenantDialPlan.GroupName'; Value = $Global:tdpg }
    $configTable += [PSCustomObject]@{ Name = 'OnlineVoiceRoutingPolicy.PolicyName'; Value = $Global:ovrp }
    $configTable += [PSCustomObject]@{ Name = 'OnlineVoiceRoutingPolicy.GroupName'; Value = $Global:ovrpg }
    $configTable += [PSCustomObject]@{ Name = 'CallingLineIdentity.PolicyName'; Value = $Global:cli }
    $configTable += [PSCustomObject]@{ Name = 'CallingLineIdentity.GroupName'; Value = $Global:clig }

    $configTable += [PSCustomObject]@{ Name = 'TeamsCallingPolicyUsers.PolicyName'; Value = $Global:tcpu }
    $configTable += [PSCustomObject]@{ Name = 'TeamsCallingPolicyUsers.GroupName'; Value = $Global:tcpug }
    
    $configTable += [PSCustomObject]@{ Name = 'TeamsEmergencyCallingPolicy.PolicyName'; Value = $Global:ecp }
    $configTable += [PSCustomObject]@{ Name = 'TeamsEmergencyCallingPolicy.GroupName'; Value = $Global:ecpg }
    $configTable += [PSCustomObject]@{ Name = 'TeamsEmergencyCallRoutingPolicy.PolicyName'; Value = $Global:ecrp }
    $configTable += [PSCustomObject]@{ Name = 'TeamsEmergencyCallRoutingPolicy.GroupName'; Value = $Global:ecrpg }

    $configTable += [PSCustomObject]@{ Name = 'TeamsSharedCallingPolicy.PolicyName'; Value = $Global:tscp }
    $configTable += [PSCustomObject]@{ Name = 'TeamsSharedCallingPolicy.GroupName'; Value = $Global:tscpg }

    $configTable += [PSCustomObject]@{ Name = 'TeamsVoiceApplicationPolicy.PolicyName'; Value = $Global:tvap }
    $configTable += [PSCustomObject]@{ Name = 'TeamsVoiceApplicationPolicy.GroupName'; Value = $Global:tvapg }

    $configTable += [PSCustomObject]@{ Name = 'MigrationSuccessMessage'; Value = $Global:MigrationSuccessMessage }
    $configTable += [PSCustomObject]@{ Name = 'MigrationFailMessage'; Value = $Global:MigrationFailMessage }

    $configTable | Format-Table -AutoSize

    Write-Host ""
    $continue = Read-Host "Press [C] to continue"
    If ($continue.ToLower() -eq "c") {
        Return $true
    }

}



<#
    Option-0

    Pre-Check - MgGraph Status
    Checks if more than one MgGraph version is installed 
    Checks the current versions
    If more than one or wrong version it will offer the Fix-MgGraph script to be run
#>

function Check-MgGraph-Status {
    
    # Check if the MgGraph module is installed
    $mgGraphModules = Get-Module -ListAvailable -Name Microsoft.Graph

    if ($mgGraphModules.Count -eq 0) {
        Write-Warning "Microsoft.Graph module is not installed. Please run the Fix-MgGraph.ps1 script before continuing." -
        return
    }

    # Check the installed versions of the MgGraph module
    $installedVersions = $mgGraphModules | Select-Object -ExpandProperty Version | Sort-Object -Descending

    if ($installedVersions.Count -gt 1) {
        Write-Warning "Multiple versions of the Microsoft.Graph module are installed:"
        $installedVersions | ForEach-Object { Write-Host "Version: $_" }
        Write-Warning "Please run the Fix-MgGraph.ps1 script to resolve this issue before continuing. 
        Run in an elevated PowerShell session. It will take between 1-3hrs to complete"
        return
    }

    # Check if the installed version is 2.25.0
    if ($installedVersions[0] -ne [Version]"2.25.0") {
        Write-Warning "The installed version of Microsoft.Graph is $($installedVersions[0])."
        Write-Warning "Please run the Fix-MgGraph.ps1 script to update to version 2.25.0 before continuing.
        Run in an elevated PowerShell session. It will take between 1-3hrs to complete"
        return
    }

    Write-Host "Microsoft.Graph module is installed and the version is 2.25.0. No remediation needed." -ForegroundColor Green

    Write-Host ""

    $continue = Read-Host "Press [C] to continue"
    If ($continue.ToLower() -eq "c") {
        Return $true
    }
}


<# 
    Option 1
    Pre-Check Licensing Status
    This will check the licensing status of the users and export the results to a CSV file.
    Use the config.xml file to set the ServicePlanName for the license you want to check. 
#>
Function Licensing-status {
    # Check if the batch file has been selected first, if not the user is prompted to run option B first.
    if (-not (Validate-BatchFile)) {
        Write-Warning "Batch file validation failed. Exiting function."
        return
    }

    $expfile = "$path\exp\$($sitecode)-Licensing-status.csv"

    If (Test-Path -path $expfile -PathType Leaf) {
        Clear-Content $expfile
    }
    
    # Split the license types to check
    $LicensesToCheck = $Global:spn -split ',' | ForEach-Object { $_.Trim() }
    
    # Initialize tracking for CSV output
    $licenseData = @()
    
    # Initialize counters for each license type
    $licenseCounts = @{}
    foreach ($license in $LicensesToCheck) {
        $licenseCounts[$license] = @{
            Assigned = 0
            NotAssigned = 0
        }
    }
    
    $targetusers = Import-Csv "$batchFile"
    $userarrayq = $targetusers | Measure
    $file = $userarrayq.count
    $i = 0
    cls
    
    ForEach ($user in $targetusers) {
        Try {
            $percentcomplete = ($i / $file) * 100; $i++
            Write-Progress -Activity ('Licensing status for: ' + $user.UPN) -Status ('file ' + $i + ' of ' + $file) -PercentComplete $percentcomplete

            $userLicenses = (Get-MgUserLicenseDetail -UserId $user.UPN -ErrorAction Stop).ServicePlans

            foreach ($license in $LicensesToCheck) {
                $hasLicense = $userLicenses | Where-Object { $_.ServicePlanName -eq $license }
                
                if ($hasLicense) {
                    $licenseCounts[$license].Assigned++
                    $status = "$license - Assigned"
                } else {
                    $licenseCounts[$license].NotAssigned++
                    $status = "$license - Not assigned"
                }
                
                # Add to data for CSV export
                $licenseData += [PSCustomObject]@{
                    'Timestamp' = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                    'User principal name' = $user.UPN
                    'Voice license' = $status
                }
            }
        }
        catch {
            echo ('Timestamp: {0}' -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) ('No Cigar: {0}' -f $user.UPN) ('Error: Licensing Status - {0}' -f $_.Exception.Message) | Out-File "$path\logs\License_Status_debugs.txt" -Append
        }
    }
    
    # Export the detailed license data to CSV
    $licenseData | Export-Csv "$path\exp\$($sitecode)-Licensing-status_$((Get-Date).ToString('dd-MM-yyyy')).csv" -NoTypeInformation
    
    Write-Progress -Activity "Sleep" -Completed
    
    # Build a summary table for display
    $totalsTable = @()
    foreach ($license in $LicensesToCheck) {
        $totalsTable += [PSCustomObject]@{ 
            License = $license
            Status = 'Assigned'
            Count = $licenseCounts[$license].Assigned 
        }
        $totalsTable += [PSCustomObject]@{ 
            License = $license
            Status = 'Not Assigned'
            Count = $licenseCounts[$license].NotAssigned 
        }
    }
    
    # Output the totals table
    Cls
    Write-Host -object "License Assignment Status Overview for $file users"
    
    $totalsTable | Format-Table -AutoSize
    Write-Host -object ""
    Write-Host -Object "See the $path\exp\$($sitecode)-Licensing-status_$((Get-Date).ToString('dd-MM-yyyy')).csv file for more details"
    Write-Host -Object ""

    $continue = Read-Host "Press [C] to continue"
    If ($continue.ToLower() -eq "c") {
        Return $true
    }
}


<#
    Option 2
    Pre-Check Policy Assignments
    This function will check the assignments of Teams policies and Entra groups for the users in the batch file.
    It will prompt the user to select which assignments to check and call the appropriate function.
#>

Function Check-Assignments {
    
    $confirm = Read-Host "Which assignments do you want to check? 
    [1] Teams Policies 
    [2] Entra Groups
    [C] Cancel

    Selection "
    Cls

    switch ($confirm) {
        "1" {
            Check-TeamsPolicyAssignments
        }
        "2" {
            Check-GroupBasedPolicyAssignment
        }
        default {
            Write-Warning "Invalid selection. Please choose [1] or [2]."
        }
    }
}


<#
    Pre-Check Check-TeamsPolicyAssignments
    Option 2.1
    This function will check the policy name from within the config.xml file and check if the policy exists in Teams.
    The function will return true if the policy exists, otherwise it will return false, outputting the policy name and the status in a tabel format.
#>

Function Check-TeamsPolicyAssignments {

    $policyStatusTable = @()
    Write-Host "Getting Teams Policy Status"
    Write-Host "This may take a few minutes, please be patient..."
    Write-Host -Object ''
    foreach ($policy in $Global:config.settings.ChildNodes) {
        if ($null -ne $policy.PolicyName -and $policy.PolicyName -ne "") {
            try {
                $policyType = $policy.Name
                
                # Initialize a variable to track if the policy exists
                $policyExists = $null

                # Check the policy type and call the appropriate cmdlet
                switch ($policyType) {
                    "TenantDialPlan" {
                        $policyExists = Get-CsTenantDialPlan -Identity $policy.PolicyName -ErrorAction SilentlyContinue
                    }
                    "OnlineVoiceRoutingPolicy" {
                        $policyExists = Get-CsOnlineVoiceRoutingPolicy -Identity $policy.PolicyName -ErrorAction SilentlyContinue
                    }
                    "CallingLineIdentity" {
                        $policyExists = Get-CsCallingLineIdentity -Identity $policy.PolicyName -ErrorAction SilentlyContinue
                    }
                    "TeamsCallingPolicyUsers" {
                        $policyExists = Get-CsTeamsCallingPolicy -Identity $policy.PolicyName -ErrorAction SilentlyContinue
                    }
                    "TeamsSharedCallingPolicy" {
                        $policyExists = Get-CsTeamsSharedCallingRoutingPolicy -Identity $policy.PolicyName -ErrorAction SilentlyContinue
                    }
                    "TeamsEmergencyCallingPolicy" {
                        $policyExists = Get-CsTeamsEmergencyCallingPolicy -Identity $policy.PolicyName -ErrorAction SilentlyContinue
                    }
                    "TeamsEmergencyCallRoutingPolicy" {
                        $policyExists = Get-CsTeamsEmergencyCallRoutingPolicy -Identity $policy.PolicyName -ErrorAction SilentlyContinue
                    }
                    "TeamsVoiceApplicationPolicy" {
                        $policyExists = Get-CsTeamsVoiceApplicationsPolicy -Identity $policy.PolicyName -ErrorAction SilentlyContinue
                    }
                    default {
                        throw "Unsupported policy type: $policyType"
                    }
                }

                if ($policyExists) {
                    $policyStatusTable += [PSCustomObject]@{
                        PolicyType     = $policyType
                        PolicyName     = $policy.PolicyName
                        Status         = "Exists"
                    }
                } else {
                    $policyStatusTable += [PSCustomObject]@{
                        PolicyType     = $policyType
                        PolicyName     = $policy.PolicyName
                        Status         = "Not Exists"
                    }
                }
            } catch {
                $policyStatusTable += [PSCustomObject]@{
                    PolicyType     = $policyType
                    PolicyName     = $policy.PolicyName
                    Status         = "Error"
                }
                Write-Output ('Timestamp: {0}' -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) ('Policy Not Found: {0}' -f $policy.PolicyName) ('Error: {0}' -f $_.Exception.Message) | Out-File "$path\logs\TPA_Status_debugs.txt" -Append
            }
        }
    }

    cls
    $policyStatusTable | Format-Table -Property PolicyType, PolicyName, Status -AutoSize

    Write-Host -Object ''
        
    $continue = Read-Host "Press [C] to continue"
    If ($continue.ToLower() -eq "c") {
        Return $true
    }
}


<#
    Option 2.2
    Pre-Check Group Based Policy Assignments
    Check each Group Based Policy Assignment from the config.xml file and check if the group exists in Azure AD.
    The function will return true if the group exists, otherwise it will return false, outputting the group name and the status in the debug logs.
#>

Function Check-GroupBasedPolicyAssignment {
    Cls
    $groupStatusTable = @()
    Write-Host "Getting Groups for Teams Group based Policy assignment Status"
    Write-Host "This may take a while, please be patient..."
    Write-Host -Object ''
    foreach ($policy in $Global:config.settings.ChildNodes) {
        if ($null -ne $policy.GroupName -and $policy.GroupName -ne "") {
            try {
                $group = Get-MgGroup -Filter "DisplayName eq '$($policy.GroupName)'" -ErrorAction Stop
                $groupMembersCount = (Get-MgGroupMember -GroupId $group.Id -All).Count
                $groupStatusTable += [PSCustomObject]@{
                    GroupName     = $policy.GroupName
                    Status        = "Exists"
                    GroupMembers  = $groupMembersCount
                }
            } catch {
                $groupStatusTable += [PSCustomObject]@{
                    GroupName     = $policy.GroupName
                    Status        = "Not Exists"
                    GroupMembers  = "N/A"
                }
                Write-Output ('Timestamp: {0}' -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) ('Group Not Found: {0}' -f $policy.GroupName) ('Error: {0}' -f $_.Exception.Message) | Out-File "$path\logs\GBPA_Status_debugs.txt" -Append
            }
        }
    }
    
    Cls
    $groupStatusTable | Format-Table -AutoSize

    Write-Host -Object ''
        
    $continue = Read-Host "Press [C] to continue"
    If ($continue.ToLower() -eq "c") {
        Return $true
    }
}



<# 
    Option-3
    Teams user status
    This will export the Teams user status to a CSV file, including the following attributes: UserPrincipalName, LineUri, CallingLineIdentity, TeamsCallingPolicy, TenantDialPlan, OnlineVoiceRoutingPolicy
    Note: Tweak the select statement to include any other attributes you want to export.
#>
Function Teams-user-status {

    # Check if the batch file has been selected first, if not the user is prompted to run option B first.
    if (-not (Validate-BatchFile)) {
        Write-Warning "Batch file validation failed. Exiting function."
        return
    }

    $expfile = "$path\exp\$($sitecode)-Teams-user-status.csv"

    If (Test-Path -path $expfile -PathType Leaf) {
        Clear-Content $expfile
    } 

    $targetusers = Import-Csv "$batchfile"
    $userarrayq = $targetusers | Measure
    $file = $userarrayq.count
    $i = 0

    ForEach ($user in $targetusers) {
        Try {
            $percentcomplete = ($i / $file) * 100; $i++
            Write-Progress -Activity ('Exporting... ' + $user.UPN) -Status ('file ' + $i + ' of ' + $file) -PercentComplete $percentcomplete

            #Get-CsOnlineUser -Identity $user.UPN -ErrorAction Stop | Select  | Export-Csv "$path\exp\$($sitecode)-Teams-user-status.csv" -Append -NoTypeInformation
            Get-CsOnlineUser -Identity $user.UPN -ErrorAction Stop | Select @{Name="Timestamp";Expression={(Get-Date).ToString("yyyy-MM-dd HH:mm:ss")}}, UserPrincipalName, AccountType, HostingProvider, InterpretedUserType, TeamsUpgradeEffectiveMode, UsageLocation, EnterpriseVoiceEnabled, LineUri, TeamsCallingPolicy, TenantDialPlan, OnlineVoiceRoutingPolicy, CallingLineIdentity, TeamsCallHoldPolicy, TeamsCallParkPolicy, TeamsEmergencyCallRoutingPolicy, TeamsEmergencyCallingPolicy, TeamsEnhancedEncryptionPolicy, TeamsIPPhonePolicy, TeamsMeetingPolicy, TeamsMobilityPolicy, TeamsSharedCallingRoutingPolicy, TeamsVdiPolicy, TeamsVoiceApplicationsPolicy | Export-Csv "$path\exp\$($sitecode)-Teams-user-voice-status_$((Get-Date).ToString('dd-MM-yyyy')).csv" -Append -NoTypeInformation
            #Get-CsOnlineVoicemailUserSettings -Identity $user.UPN | Select @{Name="Timestamp";Expression={(Get-Date).ToString("yyyy-MM-dd HH:mm:ss")}}, VoicemailEnabled, PromptLanguage | Export-Csv "$path\exp\$($sitecode)-Teams-user-vm-status.csv" -Append -NoTypeInformation
        }
        catch {
            echo ('Timestamp: {0}' -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) ('No Cigar: {0}' -f $user.UPN) ('Error: Teams user status - {0}' -f $_.Exception.Message) | Out-File "$path\logs\Teams_User_Status_debugs.txt" -Append
        }
    }
    Write-Progress -Activity "Sleep" -Completed
}



<#

    Option-10
    Funciton to enable users for Enterprise Voice

#>

Function Enable-UserForEV {

# Check if the batch file has been selected first, if not the user is prompted to run option B first.
    if (-not (Validate-BatchFile)) {
        Write-Warning "Batch file validation failed. Exiting function."
        return
    }

    $targetusers = Import-Csv "$batchfile"
    $userarrayq = $targetusers | Measure
    $file = $userarrayq.count
    $i = 0

     # Loop through each user in the batch file and apply the policy
    # If the policy is group based, add the user to the group instead of applying the policy directly
    ForEach ($user in $targetusers) {

        Try {
            $percentcomplete = ($i / $file) * 100; $i++
            Write-Progress -Activity ('Applying policy to... ' + $user.UPN) -Status ('file ' + $i + ' of ' + $file) -PercentComplete $percentcomplete

            Set-CsPhoneNumberAssignment -EnterpriseVoiceEnabled $true -Identity $user.UPN -ErrorAction Stop

        }
        catch {
        echo ('Timestamp: {0}' -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) ('No Cigar: {0}' -f $user.UPN) ('Error: Apply Tenant Dial Plan Policy - {0}' -f $_.Exception.Message) | Out-File "$path\logs\Tenant_Dial_Plan_debugs.txt" -Append
        }
    }
    Write-Progress -Activity "Sleep" -Completed

}


<#

    Option-11
    Function to assign Phone Extension to user
    This function is dependant on the PhoneNumber extension being in the CSV file
    The user must be cloud only for this funciton to work. If the user is Hybrid then the Attribute must be updated via On_premises AD

#>

Function Assign-XTN-Numbers {

    # Check if the batch file has been selected first, if not the user is prompted to run option B first.
    if (-not (Validate-BatchFile)) {
        Write-Warning "Batch file validation failed. Exiting function."
        return
    }

    $expfile = "$path\exp\$($sitecode)-XTN-Number-Assignment-Status.csv"

    # Ensure the export file exists and clear its content if it does
    If (Test-Path -Path $expfile -PathType Leaf) {
        Clear-Content $expfile
    } Else {
        # Create the file with headers if it doesn't exist
        "Timestamp,UPN,XTN,Status" | Out-File $expfile
    }

    $targetusers = Import-Csv "$batchfile"
    $userarrayq = $targetusers | Measure
    $file = $userarrayq.count
    $i = 0

    ForEach ($user in $targetusers) {
        Try {
            $percentcomplete = ($i / $file) * 100; $i++
            Write-Progress -Activity ('Assigning Extension number to... ' + $user.UPN) -Status ('file ' + $i + ' of ' + $file) -PercentComplete $percentcomplete

            Update-MgUser -UserId $user.UPN -BusinessPhones $user.DDI -ErrorAction Stop

            # Log success
            $logEntry = "{0},{1},{2},Success" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $user.UPN, $user.DDI
            $logEntry | Out-File $expfile -Append

        }
        catch {
            # Log failure
            $logEntry = "{0},{1},,Failed" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $user.UPN
            $logEntry | Out-File $expfile -Append

            # Log detailed error to debug file
            echo ('Timestamp: {0}' -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) ('No Cigar: {0}' -f $user.UPN) ('Error: Assign Extension Numbers - {0}' -f $_.Exception.Message) | Out-File "$path\logs\Assign_XTN_Numbers_debugs.txt" -Append

        }
    }
    Write-Progress -Activity "Sleep" -Completed
}



<# 
    
    Option-12
    Apply Tenant Dial Plan Policy
    
#>
Function Apply-Tenant-Dial-Plan-Policy {
    
    # Check if the batch file has been selected first, if not the user is prompted to run option B first.
    if (-not (Validate-BatchFile)) {
        Write-Warning "Batch file validation failed. Exiting function."
        return
    }

    $targetusers = Import-Csv "$batchfile"
    $userarrayq = $targetusers | Measure
    $file = $userarrayq.count
    $i = 0

    # Check if the policy is group based or not by checking the config.xml file
    # If true, get the Group Object ID
    $gbpa = isGroupBasedPolicyAssignment -policyName "TenantDialPlan"
    If ($gbpa -eq $true) {
        $group = Get-MgGroup -Filter "DisplayName eq '$Global:tdpg'" -ErrorAction Stop
        $groupId = $group.Id
    }
    Else {
        
    }

    # Loop through each user in the batch file and apply the policy
    # If the policy is group based, add the user to the group instead of applying the policy directly
    ForEach ($user in $targetusers) {

        Try {
            $percentcomplete = ($i / $file) * 100; $i++
            Write-Progress -Activity ('Applying policy to... ' + $user.UPN) -Status ('file ' + $i + ' of ' + $file) -PercentComplete $percentcomplete

            if ($gbpa -eq $false) { 
                Grant-CsTenantDialPlan -Identity $user.UPN -PolicyName $Global:tdp -ErrorAction Stop
            }
            else {
                # Get user ID from the UPN and add user as a member to the group
                $userId = (Get-MgUser -Filter "UserPrincipalName eq '$($user.UPN)'" -ErrorAction Stop).Id
                New-MgGroupMember -GroupId $groupId -DirectoryObjectId $userId -ErrorAction Stop
            }

        }
        catch {
        echo ('Timestamp: {0}' -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) ('No Cigar: {0}' -f $user.UPN) ('Error: Apply Tenant Dial Plan Policy - {0}' -f $_.Exception.Message) | Out-File "$path\logs\Tenant_Dial_Plan_debugs.txt" -Append
        }
    }
    Write-Progress -Activity "Sleep" -Completed
}



<# 
    Option-13
    Apply Online Voice Routing Policy
#>

Function Apply-Online-Voice-Routing-Policy {
    
    # Check if the batch file has been selected first, if not the user is prompted to run option B first.
    if (-not (Validate-BatchFile)) {
        Write-Warning "Batch file validation failed. Exiting function."
        return
    }

    $targetusers = Import-Csv "$batchfile"
    $userarrayq = $targetusers | Measure
    $file = $userarrayq.count
    $i = 0

    $gbpa = isGroupBasedPolicyAssignment -policyName "OnlineVoiceRoutingPolicy"
    If ($gbpa -eq $true) {
        $group = Get-MgGroup -Filter "DisplayName eq '$Global:ovrpg'" -ErrorAction Stop
        $groupId = $group.Id
    }
    Else {
        
    }

    # Loop through each user in the batch file and apply the policy
    # If the policy is group based, add the user to the group instead of applying the policy directly
    ForEach ($user in $targetusers) {

        Try {
            $percentcomplete = ($i / $file) * 100; $i++
            Write-Progress -Activity ('Applying policy to... ' + $user.UPN) -Status ('file ' + $i + ' of ' + $file) -PercentComplete $percentcomplete

            if ($gbpa -eq $false) { 
                Grant-CsOnlineVoiceRoutingPolicy -Identity $user.UPN -PolicyName $Global:ovrp -ErrorAction Stop
            }
            else {
                # Get user ID from the UPN and add user as a member to the group
                $userId = (Get-MgUser -UserId $user.UPN).Id
                New-MgGroupMember -GroupId $groupId -DirectoryObjectId $userId -ErrorAction Stop
            }
            
        }
        catch {
            echo ('Timestamp: {0}' -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) ('No Cigar: {0}' -f $user.UPN) ('Error: Apply Online Voice Routing Policy - {0}' -f $_.Exception.Message) | Out-File "$path\logs\Online_Voice_Routing_debugs.txt" -Append
        }
    }
    Write-Progress -Activity "Sleep" -Completed
}


<# 
    Option-14
    Apply Teams Calling Policy
#>
Function Apply-Teams-Calling-Policy {

    # Check if the batch file has been selected first, if not the user is prompted to run option B first.
    if (-not (Validate-BatchFile)) {
        Write-Warning "Batch file validation failed. Exiting function."
        return
    }

    Cls
    $targetusers = Import-Csv "$batchFile"
    $userarrayq = $targetusers | Measure
    $file = $userarrayq.count
    $i = 0

        ForEach ($user in $targetusers) {

            Try {
                $percentcomplete = ($i / $file) * 100; $i++
                Write-Progress -Activity ('Applying policy to... ' + $user.UPN) -Status ('file ' + $i + ' of ' + $file) -PercentComplete $percentcomplete

                
                    # Check if the policy is group based or not by checking the config.xml file
                    $gbpa = isGroupBasedPolicyAssignment -policyName "TeamsCallingPolicyUsers"
                
                    If ($gbpa -eq $false) { 
                        Grant-CsTeamsCallingPolicy -Identity $user.UPN -PolicyName $Global:tcpu -ErrorAction Stop
                        Set-CsOnlineVoicemailUserSettings -Identity $user.UPN -PromptLanguage "en-GB" -ErrorAction Stop
                    }
                    Else {
                        # Get Group ID from the config file and add user as a member to the group
                        $group = Get-MgGroup -Filter "DisplayName eq '$Global:tcpug'" -ErrorAction Stop
                        $groupId = $group.Id
                        $userId = (Get-MgUser -UserId $user.UPN).Id
                        New-MgGroupMember -GroupId $groupId -DirectoryObjectId $userId -ErrorAction Stop
                        Set-CsOnlineVoicemailUserSettings -Identity $user.UPN -PromptLanguage "en-GB" -ErrorAction Stop
                    }

                }
            catch {
                echo ('Timestamp: {0}' -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) ('No Cigar: {0}' -f $user.UPN) ('Error: Apply Teams Calling Policy - {0}' -f $_.Exception.Message) | Out-File "$path\logs\Calling_Policy_debugs.txt" -Append
            }
        }
    
    Write-Progress -Activity "Sleep" -Completed
}



<# 
    Option-15
    Apply Caller ID Policy
#>
Function Apply-Caller-ID-Policy {

    # Check if the batch file has been selected first, if not the user is prompted to run option B first.
    if (-not (Validate-BatchFile)) {
        Write-Warning "Batch file validation failed. Exiting function."
        return
    }

    $targetusers = Import-Csv "$batchfile"

    $userarrayq = $targetusers | Measure
    $file = $userarrayq.count
    $i = 0

    # Check if the policy is group based or not by checking the config.xml file
    # If true, get the Group Object ID
    $gbpa = isGroupBasedPolicyAssignment -policyName "CallingLineIdentity"
    If ($gbpa -eq $true) {
        $group = Get-MgGroup -Filter "DisplayName eq '$Global:clig'" -ErrorAction Stop
        $groupId = $group.Id
    }
    Else {
        
    }

    ForEach ($user in $targetusers) {

        Try {
            $percentcomplete = ($i / $file) * 100; $i++
            Write-Progress -Activity ('Applying policy to... ' + $user.UPN) -Status ('file ' + $i + ' of ' + $file) -PercentComplete $percentcomplete

            if ($gbpa -eq $false) { 
                Grant-CsCallingLineIdentity -Identity $user.UPN -PolicyName $Global:cli -ErrorAction Stop
            }
            else {
                # Get the User ID from the UPN and add them to the group
                $userId = (Get-MgUser -UserId $user.UPN).Id
                New-MgGroupMember -GroupId $groupId -DirectoryObjectId $userId -ErrorAction Stop
            }

        }
        catch {
            echo ('Timestamp: {0}' -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) ('No Cigar: {0}' -f $user.UPN) ('Error: Apply Caller ID Policy - {0}' -f $_.Exception.Message) | Out-File "$path\logs\Caller_ID_Policy_debugs.txt" -Append
        }
    }
    Write-Progress -Activity "Sleep" -Completed
}



<# 
    Option-16
    Apply Shared Calling Policy
#>
Function Apply-Teams-Shared-Calling-Policy {

    # Check if the batch file has been selected first, if not the user is prompted to run option B first.
    if (-not (Validate-BatchFile)) {
        Write-Warning "Batch file validation failed. Exiting function."
        return
    }

    $targetusers = Import-Csv "$batchfile"
    $userarrayq = $targetusers | Measure
    $file = $userarrayq.count
    $i = 0

    $gbpa = isGroupBasedPolicyAssignment -policyName "TeamsSharedCallingPolicy"
    If ($gbpa -eq $true) {
        $group = Get-MgGroup -Filter "DisplayName eq '$Global:tscpg'" -ErrorAction Stop
        $groupId = $group.Id
    }
    Else {
        
    }

    # Loop through each user in the batch file and apply the policy
    # If the policy is group based, add the user to the group instead of applying the policy directly
    ForEach ($user in $targetusers) {

        Try {
            $percentcomplete = ($i / $file) * 100; $i++
            Write-Progress -Activity ('Applying policy to... ' + $user.UPN) -Status ('file ' + $i + ' of ' + $file) -PercentComplete $percentcomplete

            if ($gbpa -eq $false) { 
                Grant-CsTeamsSharedCallingRoutingPolicy -Identity $user.UPN -PolicyName $Global:tscp -ErrorAction Stop
            }
            else {
                # Get user ID from the UPN and add user as a member to the group
                $userId = (Get-MgUser -UserId $user.UPN).Id
                New-MgGroupMember -GroupId $groupId -DirectoryObjectId $userId -ErrorAction Stop
            }
            
        }
        catch {
            echo ('Timestamp: {0}' -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) ('No Cigar: {0}' -f $user.UPN) ('Error: Apply Shared Calling Policy - {0}' -f $_.Exception.Message) | Out-File "$path\logs\Teams_Shared_Calling_debugs.txt" -Append
        }
    }
    Write-Progress -Activity "Sleep" -Completed
}


<# 
    Option-17
    Apply Emergency Calling Policy
#>
Function Apply-Teams-Emergency-Calling-Policy {

    # Check if the batch file has been selected first, if not the user is prompted to run option B first.
    if (-not (Validate-BatchFile)) {
        Write-Warning "Batch file validation failed. Exiting function."
        return
    }

    $targetusers = Import-Csv "$batchfile"
    $userarrayq = $targetusers | Measure
    $file = $userarrayq.count
    $i = 0

    $gbpa = isGroupBasedPolicyAssignment -policyName "TeamsEmergencyCallingPolicy"
    If ($gbpa -eq $true) {
        $group = Get-MgGroup -Filter "DisplayName eq '$Global:ecpg'" -ErrorAction Stop
        $groupId = $group.Id
    }
    Else {
        
    }

    # Loop through each user in the batch file and apply the policy
    # If the policy is group based, add the user to the group instead of applying the policy directly
    ForEach ($user in $targetusers) {

        Try {
            $percentcomplete = ($i / $file) * 100; $i++
            Write-Progress -Activity ('Applying policy to... ' + $user.UPN) -Status ('file ' + $i + ' of ' + $file) -PercentComplete $percentcomplete

            if ($gbpa -eq $false) { 
                Grant-CsTeamsEmergencyCallingPolicy -Identity $user.UPN -PolicyName $Global:ecp -ErrorAction Stop
            }
            else {
                # Get user ID from the UPN and add user as a member to the group
                $userId = (Get-MgUser -UserId $user.UPN).Id
                New-MgGroupMember -GroupId $groupId -DirectoryObjectId $userId -ErrorAction Stop
            }
            
        }
        catch {
            echo ('Timestamp: {0}' -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) ('No Cigar: {0}' -f $user.UPN) ('Error: Apply Emergency Calling Policy - {0}' -f $_.Exception.Message) | Out-File "$path\logs\Teams_Emergency_Calling_debugs.txt" -Append
        }
    }
    Write-Progress -Activity "Sleep" -Completed
}


<# 
    Option-18
    Apply Emergency Call Routing Policy
#>
Function Apply-Teams-Emergency-Call-Routing-Policy {

    # Check if the batch file has been selected first, if not the user is prompted to run option B first.
    if (-not (Validate-BatchFile)) {
        Write-Warning "Batch file validation failed. Exiting function."
        return
    }

    $targetusers = Import-Csv "$batchfile"
    $userarrayq = $targetusers | Measure
    $file = $userarrayq.count
    $i = 0

    $gbpa = isGroupBasedPolicyAssignment -policyName "TeamsEmergencyCallRoutingPolicy"
    If ($gbpa -eq $true) {
        $group = Get-MgGroup -Filter "DisplayName eq '$Global:ecrpg'" -ErrorAction Stop
        $groupId = $group.Id
    }
    Else {
        
    }

    # Loop through each user in the batch file and apply the policy
    # If the policy is group based, add the user to the group instead of applying the policy directly
    ForEach ($user in $targetusers) {

        Try {
            $percentcomplete = ($i / $file) * 100; $i++
            Write-Progress -Activity ('Applying policy to... ' + $user.UPN) -Status ('file ' + $i + ' of ' + $file) -PercentComplete $percentcomplete

            if ($gbpa -eq $false) { 
                Grant-CsTeamsEmergencyCallRoutingPolicy -Identity $user.UPN -PolicyName $Global:ecrp -ErrorAction Stop
            }
            else {
                # Get user ID from the UPN and add user as a member to the group
                $userId = (Get-MgUser -UserId $user.UPN).Id
                New-MgGroupMember -GroupId $groupId -DirectoryObjectId $userId -ErrorAction Stop
            }
            
        }
        catch {
            echo ('Timestamp: {0}' -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) ('No Cigar: {0}' -f $user.UPN) ('Error: Apply Emergency Call Routing Policy - {0}' -f $_.Exception.Message) | Out-File "$path\logs\Teams_Emergency_Call_Routing_debugs.txt" -Append
        }
    }
    Write-Progress -Activity "Sleep" -Completed
}



<# 
    Option-19
    Apply Teams Voice Application Policy
#>
Function Apply-Teams-Voice-Applications-Policy {

    # Check if the batch file has been selected first, if not the user is prompted to run option B first.
    if (-not (Validate-BatchFile)) {
        Write-Warning "Batch file validation failed. Exiting function."
        return
    }

    $targetusers = Import-Csv "$batchfile"
    $userarrayq = $targetusers | Measure
    $file = $userarrayq.count
    $i = 0

    $gbpa = isGroupBasedPolicyAssignment -policyName "TeamsVoiceApplicationPolicy"
    If ($gbpa -eq $true) {
        $group = Get-MgGroup -Filter "DisplayName eq '$Global:tvapg'" -ErrorAction Stop
        $groupId = $group.Id
    }
    Else {
        
    }

    # Loop through each user in the batch file and apply the policy
    # If the policy is group based, add the user to the group instead of applying the policy directly
    ForEach ($user in $targetusers) {

        Try {
            $percentcomplete = ($i / $file) * 100; $i++
            Write-Progress -Activity ('Applying policy to... ' + $user.UPN) -Status ('file ' + $i + ' of ' + $file) -PercentComplete $percentcomplete

            if ($gbpa -eq $false) { 
                Grant-CsTeamsVoiceApplicationsPolicy -Identity $user.UPN -PolicyName $Global:tvap -ErrorAction Stop
            }
            else {
                # Get user ID from the UPN and add user as a member to the group
                $userId = (Get-MgUser -UserId $user.UPN).Id
                New-MgGroupMember -GroupId $groupId -DirectoryObjectId $userId -ErrorAction Stop
            }
            
        }
        catch {
            echo ('Timestamp: {0}' -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) ('No Cigar: {0}' -f $user.UPN) ('Error: Apply Teams Voice Application Policy - {0}' -f $_.Exception.Message) | Out-File "$path\logs\Teams_Voice_Application_debugs.txt" -Append
        }
    }
    Write-Progress -Activity "Sleep" -Completed
}


<#
    Option-20
    Send-Bulk-Chat Function
    Use this function to send a status message directly to the user via Teams Chat in bulk at the end of the enablement.
    It will sent the user a message detailing their assigned Shared Calling Numnber and their Extension Number if assigned.

#>
Function Send-Bulk-Chat {

    # Check if the batch file has been selected first, if not the user is prompted to run option B first.
    if (-not (Validate-BatchFile)) {
        Write-Warning "Batch file validation failed. Exiting function."
        return
    }

    $targetusers = Import-Csv "$batchfile"
    $userarrayq = $targetusers | Measure
    $file = $userarrayq.count
    $i = 0

    ForEach ($user in $targetusers) {
        Try {
            $percentcomplete = ($i / $file) * 100; $i++
            Write-Progress -Activity ('Sending Chat to... ' + $user.UPN) -Status ('file ' + $i + ' of ' + $file) -PercentComplete $percentcomplete
            

            <# This needs to be updated, however waiting on appropriate licensing in my tenant to test! #>
            try {
                $SharedCallingPolicyAssigned = Get-CsOnlineUser -Identity $user.UPN | Select-Object TeamsSharedCallingRoutingPolicy
                if (-not $SharedCallingPolicyAssigned) {
                    throw "Failed to retrieve Teams Shared Calling Routing Policy for user $($user.UPN)"
                }
            } catch {
                Write-Output ('Timestamp: {0}' -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) ('Error getting shared calling policy for: {0}' -f $user.UPN) ('Error: {0}' -f $_.Exception.Message) | Out-File "$path\logs\Bulk_Chat_debugs.txt" -Append
                return # Exit the function
            }

            try {
                $SharedCallingAssignedRA = Get-CsTeamsSharedCallingRoutingPolicy -Identity $SharedCallingPolicyAssigned.TeamsSharedCallingRoutingPolicy | Select-Object ResourceAccount
                if (-not $SharedCallingAssignedRA) {
                    throw "Failed to retrieve Resource Account for policy $($SharedCallingPolicyAssigned.TeamsSharedCallingRoutingPolicy)"
                }
            } catch {
                Write-Output ('Timestamp: {0}' -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) ('Error getting resource account for: {0}' -f $user.UPN) ('Error: {0}' -f $_.Exception.Message) | Out-File "$path\logs\Bulk_Chat_debugs.txt" -Append
                return # Exit the function
            }

            try {
                $SharedCallingPhoneNumber = Get-CsOnlineApplicationInstance -Identity $SharedCallingAssignedRA.ResourceAccount | Select-Object PhoneNumber
                if (-not $SharedCallingPhoneNumber) {
                    throw "Failed to retrieve phone number for Resource Account $($SharedCallingAssignedRA.ResourceAccount)"
                }
            } catch {
                Write-Output ('Timestamp: {0}' -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) ('Error getting phone number for: {0}' -f $user.UPN) ('Error: {0}' -f $_.Exception.Message) | Out-File "$path\logs\Bulk_Chat_debugs.txt" -Append
                return # Exit the function
            }

            try {
                $AssignedExtension = Get-MgUser -UserId $user.UPN | Select-Object BusinessPhones
                if (-not $AssignedExtension) {
                    throw "Failed to retrieve business phones for user $($user.UPN)"
                }
            } catch {
                Write-Output ('Timestamp: {0}' -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) ('Error getting extension for: {0}' -f $user.UPN) ('Error: {0}' -f $_.Exception.Message) | Out-File "$path\logs\Bulk_Chat_debugs.txt" -Append
                return # Exit the function
            }

            # If the phone number is not null or empty, send a chat message to the user
            if ($null -ne $SharedCallingPhoneNumber.PhoneNumber -and $SharedCallingPhoneNumber.PhoneNumber -ne "") {
                # If the phone number starts with "tel:", remove it for display purposes
                if ($AssignedExtension.BusinessPhones -like "x*") {
                    # Send chat message to notify the user of their successful migration and new phone number
                    Send-Chat2 -To $user.UPN -MigrationStatus $Global:MigrationSuccessMessage -PhoneNumber ($SharedCallingPhoneNumber.PhoneNumber -replace '^tel:', '') -Extn $AssignedExtension.BusinessPhones
                }
                else {
                    # Send chat message to notify the user of their successful migration and new phone number
                    Send-Chat2 -To $user.UPN -MigrationStatus $Global:MigrationSuccessMessage -PhoneNumber ($SharedCallingPhoneNumber.PhoneNumber -replace '^tel:', '') -Extn "No Extension Assigned"
                }
            }
            else {
                # Send chat message to user to state migration status failure
                Send-Chat2 -To $user.UPN -MigrationStatus $Global:MigrationFailMessage
            }
            
        }
        catch {
            
            # Log detailed error to debug file
            Write-Output ('Timestamp: {0}' -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) ('No Cigar: {0}' -f $user.UPN) ('Error: Assign Phone Numbers - {0}' -f $_.Exception.Message) | Out-File "$path\logs\Bulk_Chat_debugs.txt" -Append
            
        }
    }
    Write-Progress -Activity "Sleep" -Completed
}

<#
    Help
    Option H
    This function is to provide help to the user on how to use the script and what each option does.
#>

Function Helpme {

    Write-Host -Object "Help - Teams Shared Calling Bulk User Enablement Script" -ForegroundColor Yellow
    Write-Host -Object ""
    Write-Host -Object "This script is designed to assist with bulk enabling users for Teams Shared Calling"
    Write-Host -Object "and managing related Teams voice and policy assignments. Below is a guide to"
    Write-Host -Object "the menu options available in the script:"
    Write-Host -Object ""

    Write-Host -Object "General Options:" -ForegroundColor Cyan
    Write-Host -Object "A. Archive: Archives all current logs and exported files into a zip file"
    Write-Host -Object "   located in the path\arc folder, and clears the cache."
    Write-Host -Object "B. Batch File: Allows you to select a batch file containing user data for"
    Write-Host -Object "   processing. The batch file must be imported prior to any enablement tasks."
    Write-Host -Object "   The batch file should be in CSV format with the following headers: UPN, DDI."
    Write-Host -Object "   The DDI column is optional and should be used to assign an extension (e.g., +67890)."
    Write-Host -Object "C. Show Config File: Displays the currently loaded configuration from config.xml."
    Write-Host -Object "D. Clear Debugs: Clears all debug logs located in the path\logs folder."
    Write-Host -Object ""

    Write-Host -Object "Pre-Check Options:" -ForegroundColor Cyan
    Write-Host -Object "0. Microsoft Graph Status: Checks the current version of Microsoft.Graph and"
    Write-Host -Object "   if remediation is needed by running Fix-MgGraph.ps1."
    Write-Host -Object "1. Licensing Status: Checks and exports the licensing status of users in the batch."
    Write-Host -Object "2. Check Policy Assignments: Verifies if Teams policies or group-based policy"
    Write-Host -Object "   assignments exist in Teams / Entra ID."
    Write-Host -Object "3. Teams User Status: Exports Teams user status, including policies and"
    Write-Host -Object "   attributes, to a CSV file."
    Write-Host -Object ""

    Write-Host -Object "Enterprise Voice Enablement:" -ForegroundColor Cyan
    Write-Host -Object "10. Enable Users for Enterprise Voice: Enables Enterprise Voice for users."
    Write-Host -Object "11. Apply User Extension Numbers: Assigns extension numbers to users (cloud only)."
    Write-Host -Object ""

    Write-Host -Object "Teams Policy Assignment Options:" -ForegroundColor Cyan
    Write-Host -Object "12. Apply Tenant Dial Plan Policy: Assigns Tenant Dial Plan policy or adds users"
    Write-Host -Object "    to the group specified in config.xml."
    Write-Host -Object "13. Apply Online Voice Routing Policy: Assigns Online Voice Routing policy or adds"
    Write-Host -Object "    users to the group specified in config.xml."
    Write-Host -Object "14. Apply Teams Calling Policy: Assigns Teams Calling policy or adds users to the"
    Write-Host -Object "    group specified in config.xml."
    Write-Host -Object "15. Apply Caller ID Policy: Assigns Caller ID policy or adds users to the group."
    Write-Host -Object "16. Apply Teams Shared Calling Policy: Assigns Shared Calling Routing policy or"
    Write-Host -Object "    adds users to the group specified in config.xml."
    Write-Host -Object "17. Apply Teams Emergency Calling Policy: Assigns Emergency Calling policy or adds"
    Write-Host -Object "    users to the group specified in config.xml."
    Write-Host -Object "18. Apply Teams Emergency Call Routing Policy: Assigns Emergency Call Routing"
    Write-Host -Object "    policy or adds users to the group specified in config.xml."
    Write-Host -Object "19. Apply Teams Voice Application Policy: Assigns Voice Application policy or adds"
    Write-Host -Object "    users to the group specified in config.xml."
    Write-Host -Object ""

    Write-Host -Object "User Notification:" -ForegroundColor Cyan
    Write-Host -Object "20. Send Bulk Chat: Sends a Teams chat message to each user in the batch with"
    Write-Host -Object "    their Shared Calling number and extension (if assigned)."
    Write-Host -Object ""

    Write-Host -Object "Additional Options:" -ForegroundColor Cyan
    Write-Host -Object "H. Help: Shows this help information."
    Write-Host -Object "Q. Quit: Exits the script and disconnects from all services."
    Write-Host -Object ""

    Write-Host -Object "Usage Instructions:" -ForegroundColor Green
    Write-Host -Object "1. Ensure you have the necessary permissions and required PowerShell modules:"
    Write-Host -Object "   MicrosoftTeams, Microsoft.Graph.Beta.Teams, Microsoft.Graph.Authentication."
    Write-Host -Object "2. Update the config.xml file in the script directory with your policy and group"
    Write-Host -Object "   names as needed for your environment."
    Write-Host -Object "3. Run the script, which will load the main menu."
    Write-Host -Object "4. Start by selecting Option B to import your batch file."
    Write-Host -Object "5. Run pre-checks (Options 0, 1, 2, 3) to validate environment and user readiness."
    Write-Host -Object "6. Use the enablement and policy assignment options as required for your migration."
    Write-Host -Object "7. Use Option 20 to notify users of their new Shared Calling number and extension."
    Write-Host -Object "8. Use Option A to archive logs and exports after completing your tasks."
    Write-Host -Object "9. Use Option D to clear debug logs if needed."
    Write-Host -Object "10. To exit the script, select option Q."
    Write-Host -Object ""

    Write-Host -Object "Note: Group-based policy assignment is supported. If both PolicyName and GroupName"
    Write-Host -Object "      are set in config.xml, group assignment will be used."
    Write-Host -Object ""
    Write-Host -Object "All errors encountered will be written to the path\logs folder, with a debug"
    Write-Host -Object "file per option selected."
    Write-Host -Object ""

    $continue = Read-Host "Press [C] to continue"
    If ($continue.ToLower() -eq "c") {
        Return $false
        Cls
    }
}

Function Reset-GlobalVariables {

    # Set all declared global variables to $null
    $Global:config = $null
    $Global:entraid = $null
    $Global:spn = $null
    $Global:SendChat = $null
    $Global:ChatImage = $null
    $Global:tdp = $null
    $Global:tdpg = $null
    $Global:ovrp = $null
    $Global:ovrpg = $null
    $Global:cli = $null
    $Global:clig = $null
    $Global:tcpu = $null
    $Global:tcpug = $null
    $Global:ecp = $null
    $Global:ecpg = $null
    $Global:ecrp = $null
    $Global:ecrpg = $null
    $Global:tvmp = $null
    $Global:tvmpg = $null
    $Global:tscp = $null
    $Global:tscpg = $null
    $Global:tvap = $null
    $Global:tvapg = $null
    $Global:MigrationSuccessMessage = $null
    $Global:MigrationFailMessage = $null
    $Global:UsersInBatch = $null
}

Function DisconnectServices {
    #Clear Global Variables and Cache
    Reset-GlobalVariables
    Clear-Cache
    cls

    # Disconnect from all connected services to clean up resources
    Write-Host "Disconnecting from services..."
    if (Get-Module Microsoft.Graph.Teams -ErrorAction SilentlyContinue) {
        Disconnect-MgGraph
        Write-Host "Disconnected from Microsoft Graph."
    }
    if (Get-Module MicrosoftTeams -ErrorAction SilentlyContinue) {
        Disconnect-MicrosoftTeams
        Write-Host "Disconnected from Microsoft Teams."
    }
    
    # wait for 5 seconds before exiting
    sleep 5
    Cls
}


###########################
#                         #
#      Sub Functions      #
#                         #
###########################

<# 
    Menu Options
#>
Function Option-A {
    Cls
    Write-Host 'Archiving files...'
    Archive-Files
    Clear-Cache
}


Function Option-B { 
    Cls
    $sitecode = Read-Host -Prompt "Enter a Site or Batch Code (XXXXX)"
    $batchfile = Select-csvFile
    $filename = ($($batchfile.Split("\")[$_.Length-2]))

    Options_Menu_GUI
}


Function Option-C {
    Cls
    Show-Config

    Options_Menu_GUI
}


Function Option-D {
    Cls
    Get-ChildItem -Path "$path\logs\*" | ForEach-Object { Clear-Content $_.FullName }
}

<# 
    Pre-Check Menu Options
#>


Function Option-0 {
    Cls
    Check-MgGraph-Status
}


Function Option-1 {
    Cls
    Licensing-status
}


Function Option-2 {
    Cls
    Check-Assignments
}


Function Option-3 {
    Cls
    Teams-user-status
}

<#
    Enable EV and Assign Extensions
#>

Function Option-10 {
    Cls
    Enable-UserForEV
}

Function Option-11 {
    Cls
    Assign-XTN-Numbers
}

<# 
    Assign Policy Menu Items
#>


Function Option-12 {
    Cls
    Apply-Tenant-Dial-Plan-Policy
}


Function Option-13 {
    Cls
    Apply-Online-Voice-Routing-Policy 
}


Function Option-14 {
    Cls
    Apply-Teams-Calling-Policy
}


Function Option-15 {
    Cls
    Apply-Caller-ID-Policy
}


Function Option-16 {
    Cls
    Apply-Teams-Shared-Calling-Policy
}

Function Option-17 {
    Cls
    Apply-Teams-Emergency-Calling-Policy
}

Function Option-18 {
    Cls
    Apply-Teams-Emergency-Call-Routing-Policy
}

Function Option-19 {
    Cls
    Apply-Teams-Voice-Applications-Policy
}

Function Option-20 {
    Cls
    Write-Host -Object 'Sending Bulk Chat message to user...'
    Send-Bulk-Chat
    sleep 5
}

Function Option-H {
    Cls
    Helpme
}

Function Option-Q {
    Cls
    Write-Host -Object 'Exiting script...'
    DisconnectServices
    Exit
}

###########################
#                         #
#         Runtime         #
#                         #
###########################

<# 
    GUI Menu
#>
Function Options_Menu_GUI
{
    Do
    {
        Clear-Host
        If (([string]::IsNullOrEmpty($sitecode))){
            $sitecode = 'Default'
        }

                    $headerLength = [math]::Ceiling([math]::Max(
                        [math]::Max(50 + $sitecode.Length, ("Batch File Imported: $Global:UsersInBatch Users").Length + 4),
                        "No Batch File Imported".Length + 4
                    ))

                    $border = '*' * $headerLength
                    $padding = ' ' * [math]::Max([math]::Floor(($headerLength - 2 - $sitecode.Length - 17) / 2), 0)

                    Write-Host -Object $border
                    Write-Host -Object ('*' + (' ' * ($headerLength - 2)) + '*')
                    Write-Host -Object ('*' + $padding + $sitecode + ' - Teams Migration' + $padding + '*')
                    Write-Host -Object ('*' + (' ' * ($headerLength - 2)) + '*')
                    If ($Global:UsersInBatch -and $Global:UsersInBatch -gt 0) {
                        $batchMessage = "Batch File Imported: $Global:UsersInBatch Users"
                        $batchPadding = ' ' * [math]::Max([math]::Floor(($headerLength - 2 - $batchMessage.Length) / 2), 0)
                        Write-Host -Object ('*' + $batchPadding + $batchMessage + $batchPadding + '*')
                    } Else {
                        $noBatchMessage = "No Batch File Imported"
                        $noBatchPadding = ' ' * [math]::Max([math]::Floor(($headerLength - 2 - $noBatchMessage.Length) / 2), 0)
                        Write-Host -Object ('*' + $noBatchPadding + $noBatchMessage + $noBatchPadding + '*')
                    }
                    Write-Host -Object ('*' + (' ' * ($headerLength - 2)) + '*')
                    Write-Host -Object $border
        Write-Host -Object ''
        Write-Host -Object ' Choose an option' -ForegroundColor Yellow
        #Write-Host -Object ''
        Write-Host -Object ''
        Write-Host -Object ' A.  Archive (Current)' 
        #Write-Host -Object ''
        Write-Host -Object (' B.  Batch File ' + '(' + $filename + ')')
        #Write-Host -Object ''
        Write-Host -Object ' C.  Show Config File'
        #Write-Host -Object ''
        Write-Host -Object ' D.  Clear debugs'
        Write-Host -Object ''
        Write-Host -Object ' -------- Pre-Checks --------'
        Write-Host -Object ''
        Write-Host -Object ' 0.  Check Microsoft Graph Status'
        #Write-Host -Object ''
        Write-Host -Object ' 1.  Licensing status' 
        #Write-Host -Object ''
        Write-Host -object ' 2.  Check Policy Assignment *'
        #Write-Host -Object ''
        Write-Host -Object ' 3.  Teams user status' 
        Write-Host -Object ''
        Write-Host -Object ' ----- Enterprise Voice -----'
        Write-Host -Object ''
        Write-Host -Object ' 10.  Enable Users for Enterprise Voice' 
        #Write-Host -Object ''
        Write-Host -Object ' 11.  Apply User Extension Numbers' 
        Write-Host -Object ''
        Write-Host -Object ' ----- Assign Teams Policies -----'
        Write-Host -Object ''
        Write-Host -Object ' 12.  Apply Tenant Dial Plan Policy' 
        #Write-Host -Object ''
        Write-Host -Object ' 13.  Apply Online Voice Routing Policy'
        #Write-Host -Object ''
        Write-Host -Object ' 14.  Apply Teams Calling Policy'
        #Write-Host -Object ''
        Write-Host -Object ' 15.  Apply Caller ID Policy'
        #Write-Host -Object ''
        Write-Host -Object ' 16.  Apply Teams Shared Calling Policy'
        #Write-Host -Object ''
        Write-Host -Object ' 17.  Apply Teams Emergency Calling Policy'
        #Write-Host -Object ''
        Write-Host -Object ' 18.  Apply Teams Emergency Call Routing Policy'
        #Write-Host -Object ''
        Write-Host -object ' 19.  Apply Teams Voice Application Policy'
        Write-Host -Object ''
        Write-Host -Object ' ------- Inform Users -------'
        Write-Host -Object ''
        Write-Host -object ' 20.  Send Bulk Chat - Update Message to Users with Shared Calling Number and Extension'
        Write-Host -Object ''
        Write-Host -Object ' -------     Misc     -------'
        Write-Host -Object ''
        Write-Host -Object ' H.  Help'
        Write-Host -Object ''
        Write-Host -Object ' Q.  Quit'
        Write-Host -Object ''
        #Write-Host -Object ''
        $Menu = Read-Host -Prompt ' (A, B, C, D, 1-25 or H or Q to Quit)'

 
        Switch ($Menu)
        {
                A
            {
                Option-A
            }
                B
            {
                Option-B
            }
                C
            {
                Option-C
            }
                0
            {
                Option-0
            }
                1
            {
                Option-1
            }
                2
            {
                Option-2
            }
                3
            {
                Option-3
            }
                10
            {
                Option-10
            }
                11
            {
                Option-11
            }
                12
            {
                Option-12
            }
                13
            {
                Option-13
            }
                14
            {
                Option-14
            }
                15
            {
                Option-15
            }
                16
            {
                Option-16
            }
                17
            {
                Option-17
            }
                18
            {
                Option-18
            }
            19
            {
                Option-19
            }
                20
            {
                Option-20
            }
                H
            {            
                Option-H
            }
                Q
            {
                Option-Q
                Cls
                Exit
            }
        }
    }
    until ($Menu -eq 'Q')
}

# Call the function to show the splash screen and then the menu
if (Show-SplashScreen -Title $Title -Version $Version -Delay $SplashDelay -ForegroundColor $SplashForegroundColor) {
    Options_Menu_GUI
}

