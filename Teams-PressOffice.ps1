###############################################################################################################
#
# ABOUT THIS PROGRAM
#
#   Teams-PressOffice.ps1
#   https://github.com/Headbolt/Teams-PressOffice
#
#   This script was designed to Create a Teams AutoAttendant and corresponding Call Queues
#
###############################################################################################################
#
# HISTORY
#
#   Version: 1.0 - 22/03/2021
#
#   - 22/03/2021 - V1.0 - Created by Headbolt
#
###############################################################################################################
#
#   CUSTOMISABLE FUNCTIONS SECTION
#
###############################################################################################################
#
#   CUSTOMISABLE VARIABLES FUNCTION
#
function CustomisableVariables
{
	$global:LoggingEnabled="YES" # Enable Logging if needed
	$global:LogFileLocation="....\TeamsPressOfficeCreation-Log.log" # Location Of LogFile if Enabled
	$global:AAD_Sync_Server="Server.domain.com" # FQDN of the AAD Sync Server
	$global:AAD_Sync_Server_Sync_Script="C:\Scripts\ManualADSync.ps1" # Location on AAD Sync Server of a Manual AAD Sync Script
	$global:OnPremConnectionUrl="http://Server.domain.com/PowerShell/" # URL For On Prem Hybrid Exchange Server
	$global:AzureADlicenseGroup="O365-Licenses_AZURE_Microsoft_365_Phone_System"  # AD Group to Add Resource Accounts to
	#                                                                             # In Order to Obtain relevant Licenses
	$global:GALsearchScope=". All Company All Users" # For Limiting the Voice Search Scope If Needed
	$global:Language = "en-GB"
	$global:Timezone = "GMT Standard Time"
	$global:GreetingText = ".
	Thank you for calling the $ClientName Press Office.
	Please hold while we connect your call"
#
}
#
###############################################################################################################
#
#   ORGANISATION UNIT FUNCTION
#
function OrgUnit
{
	# Assigning SPecific OU's to Specific Company Prefixes eg.
	#if ( "Company" -eq $CompanyPrefix )
	#{
	#    $global:OU="OU=Teams Phone System,OU=Distribution Groups,OU=Site,DC=Company,DC=com"
	#}
	#
}
#
###############################################################################################################
#
#   HOLIDAYS FUNCTION
#
function Holidays
{
	Write-Host 'Xmas 2020'
	$Xmas2020CallFlow = New-CsAutoAttendantCallFlow -Name "Christmas 2020" -Greetings @($greetingPrompt) -Menu $afterHoursMenu
	$Xmas2020CallHandlingAssociation = New-CsAutoAttendantCallHandlingAssociation -Type Holiday -ScheduleId 849edeee-2ce4-4485-9e91-1eb9552fbdfa -CallFlowId $Xmas2020CallFlow.Id
	Write-Host 'New Year 2021'
	$NewYear2021CallFlow = New-CsAutoAttendantCallFlow -Name "New Year 2021" -Greetings @($greetingPrompt) -Menu $afterHoursMenu
	$NewYear2021CallHandlingAssociation = New-CsAutoAttendantCallHandlingAssociation -Type Holiday -ScheduleId 745ed586-02e2-4661-b342-d04d946cda0e -CallFlowId $NewYear2021CallFlow.Id
	Write-Host 'Easter 2021'
	$Easter2021CallFlow = New-CsAutoAttendantCallFlow -Name "Easter 2021" -Greetings @($greetingPrompt) -Menu $afterHoursMenu
	$Easter2021CallHandlingAssociation = New-CsAutoAttendantCallHandlingAssociation -Type Holiday -ScheduleId 4380e982-68ef-4284-8996-4e1d3b460df6 -CallFlowId $Easter2021CallFlow.Id
	Write-Host 'May Day Bank 2021'
	$MayDay2021CallFlow = New-CsAutoAttendantCallFlow -Name "May Day 2021" -Greetings @($greetingPrompt) -Menu $afterHoursMenu
	$MayDay2021CallHandlingAssociation = New-CsAutoAttendantCallHandlingAssociation -Type Holiday -ScheduleId 73a738ea-4a04-4319-9987-398323d4c256 -CallFlowId $MayDay2021CallFlow.Id
	Write-Host 'Spring Bank 2021'
	$SpringBank2021CallFlow = New-CsAutoAttendantCallFlow -Name "Spring Bank Holiday 2021" -Greetings @($greetingPrompt) -Menu $afterHoursMenu
	$SpringBank2021CallHandlingAssociation = New-CsAutoAttendantCallHandlingAssociation -Type Holiday -ScheduleId 3553d98d-9ba1-46d2-a277-6758601c160b -CallFlowId $SpringBank2021CallFlow.Id
	Write-Host 'Summer 2021'
	$Summer2021CallFlow = New-CsAutoAttendantCallFlow -Name "Summer Bank Holiday 2021" -Greetings @($greetingPrompt) -Menu $afterHoursMenu
	$Summer2021CallHandlingAssociation = New-CsAutoAttendantCallHandlingAssociation -Type Holiday -ScheduleId 08f479c3-5a31-496e-8abc-c9eb022d1ffe -CallFlowId $Summer2021CallFlow.Id
	Write-Host 'Xmas 2021'
	$Xmas2021CallFlow = New-CsAutoAttendantCallFlow -Name "Christmas 2021" -Greetings @($greetingPrompt) -Menu $afterHoursMenu
	$Xmas2021CallHandlingAssociation = New-CsAutoAttendantCallHandlingAssociation -Type Holiday -ScheduleId a59a749e-a808-46b8-9897-c9d39b4232b6 -CallFlowId $Xmas2021CallFlow.Id
	Write-Host 'New Year 2022'
	$NewYear2022CallFlow = New-CsAutoAttendantCallFlow -Name "New Year 2022" -Greetings @($greetingPrompt) -Menu $afterHoursMenu
	$NewYear2022CallHandlingAssociation = New-CsAutoAttendantCallHandlingAssociation -Type Holiday -ScheduleId 4d3da958-12cb-419f-aec3-e12f7a4902df -CallFlowId $NewYear2022CallFlow.Id
	#
	# Set Up Call Flows and Call Handling $AfterHoursCallFlow and $AfterHoursCallHandlingAssociation MUST always
	# Remain and be first in their respective lists when Holidays are Updated
	#
	$global:CallFlows=@($AfterHoursCallFlow, $Xmas2020CallFlow, $NewYear2021CallFlow, $Easter2021CallFlow, $MayDay2021CallFlow, $SpringBank2021CallFlow, $Summer2021CallFlow, $Xmas2021CallFlow, $NewYear2022CallFlow)
	$global:CallHandlingAssociations=@($AfterHoursCallHandlingAssociation, $Xmas2020CallHandlingAssociation, $NewYear2021CallHandlingAssociation, $Easter2021CallHandlingAssociation, $MayDay2021CallHandlingAssociation, $SpringBank2021CallHandlingAssociation, $Summer2021CallHandlingAssociation, $Xmas2021CallHandlingAssociation, $NewYear2022CallHandlingAssociation)
	#
}
#
###############################################################################################################
#
#   END OF CUSTOMISABLE FUNCTIONS
#
###############################################################################################################
#
#   START FUNCTION
#
function Logging
{
	if ( $global:LoggingEnabled -eq "YES" )
	{
		Start-Transcript $global:LogFileLocation # Start the logging
		Clear-Host #Clear Screen
		Write-Output "Logging to $global:LogFileLocation"
	}     
}
#
###############################################################################################################
#
#   END FUNCTION
#
function End
{
	if ( $global:LoggingEnabled -eq "YES" )
	{
		Stop-Transcript # Stop Logging
	}
	Write-Host "END !!"
	Exit
}
#
###############################################################################################################
#
#   CONNECTIONS FUNCTION
#
function Connections
{
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host "Connecting To On-Prem Systems"
	Write-Host '' # Output To Make Screen Easier for User to read.
	$OnPremUserCredential = Get-Credential -Credential $global:AdminUser
	# Connect To On-Prem Exchange
	$OnPremSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $global:OnPremConnectionUrl -Authentication Kerberos -Credential $OnPremUserCredential
	Import-PSSession $OnPremSession -DisableNameChecking -AllowClobber -WarningAction SilentlyContinue
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host '-------------------------------------------------------------------------------------------------------------------' # Output To Make Screen Easier for User to read.
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host "Connecting To Azure AD"
	Write-Host '' # Output To Make Screen Easier for User to read.
	$AzureUserCredential = Get-Credential -Credential $global:AdminUser
	# Connect to Azure AD
	Connect-AzureAD -Credential $AzureUserCredential
	Write-Host '-------------------------------------------------------------------------------------------------------------------' # Output To Make Screen Easier for User to read.
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host "Connecting To Teams Admin"
	Write-Host '' # Output To Make Screen Easier for User to read.
	# Connect to Teams
	Connect-MicrosoftTeams -Credential $AzureUserCredential # Needs Teams Module https://www.powershellgallery.com/packages/MicrosoftTeams/2.0.0
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host '-------------------------------------------------------------------------------------------------------------------' # Output To Make Screen Easier for User to read.
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host "Connecting To MS Online Services"
	Write-Host '' # Output To Make Screen Easier for User to read.
	# Connect to MSOL
	Connect-MsolService -Credential $AzureUserCredential
	Write-Host '-------------------------------------------------------------------------------------------------------------------' # Output To Make Screen Easier for User to read.
}
#
###############################################################################################################
#
#   USER INPUT FUNCTION
#
function UserInput 
{
	Write-Host '' # Output To Make Screen Easier for User to read.
	$global:Input = '' # Ensure Input Variable is Blank
	$global:Input=Read-Host -Prompt "Input the $InputVarible . eg. $InputVaribleExplanation" # Grab the Variable
	if ( "" -ne $global:Input )
	{
#		Write-Host '' # Output To Make Screen Easier for User to read.
		Write-Host $InputVarible Value gathered is "'$global:Input'"
	}
	else
	{
		Write-Host Input was Blank, Ending Script
		End
	}
}
#
###############################################################################################################
#
#   COLLECT VARIABLES FUNCTION
#
function CollectVariables 
{
	# Setting Up Global Variables
	$global:AAD_Sync_Needed=""
	$global:DistListCreatedDAY=""
	$global:DistListCreatedNIGHT=""
	$global:CompanyPrefix=""
	$global:CompanyDomainNamePrefix=""
	$global:ClientName=""
	$global:ClientName_NoSpace = $global:ClientName -replace '\s',''
	$global:OnpremPhoneNumber=""
	$global:OU=""
	#
	Write-Host '-------------------------------------------------------------------------------------------------------------------' # Output To Make Screen Easier for User to read.
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host 'Gathering Required Data'
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host '-------------------------------------------------------------------------------------------------------------------' # Output To Make Screen Easier for User to read.
	#
	$InputVarible="Admin User Account You Want To Use"
	$InputVaribleExplanation="fredsmithBO@domain.com"
	UserInput
	$global:AdminUser=$global:Input
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host '-------------------------------------------------------------------------------------------------------------------' # Output To Make Screen Easier for User to read.
	#
	$InputVarible="CompanyPrefix"
	$InputVaribleExplanation="COM1"
	UserInput
	$global:CompanyPrefix=$global:Input.ToUpper()
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host '-------------------------------------------------------------------------------------------------------------------' # Output To Make Screen Easier for User to read.
	#
	$InputVarible="CompanyDomainNamePrefix"
	$InputVaribleExplanation="company.com"
	UserInput
	$global:CompanyDomainNamePrefix=$global:Input
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host '-------------------------------------------------------------------------------------------------------------------' # Output To Make Screen Easier for User to read.
	#
	$InputVarible="ClientName"
	$InputVaribleExplanation="AA Test Client"
	UserInput
	$global:ClientName=$global:Input
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host '-------------------------------------------------------------------------------------------------------------------' # Output To Make Screen Easier for User to read.
	#
	$InputVarible="OnpremPhoneNumber"
	$InputVaribleExplanation="441112223333"
	UserInput
	$global:OnpremPhoneNumber=$global:Input
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host '-------------------------------------------------------------------------------------------------------------------' # Output To Make Screen Easier for User to read.
	#
	$global:ClientName_NoSpace = $global:ClientName -replace '\s',''
	#
	OrgUnit
}
#
###############################################################################################################
#
#   DISTRIBUTION LIST CHECK FUNCTION
#
function DistListCheck 
{
	$OnPremDistList =""
	$OnPremDistList = (Get-DistributionGroup -Identity TCQ-$CompanyPrefix-"$ClientName_NoSpace"PO_$global:DistListToCheck -ErrorAction SilentlyContinue)
	if ( "" -ne ($OnPremDistList -replace '\s','') )
	{
		Write-Host '' # Output To Make Screen Easier for User to read.
		Write-Host Found DistributionGroup TCQ-$global:CompanyPrefix-"$ClientName_NoSpace"PO_$global:DistListToCheck
	}
	else
	{
		Write-Host '' # Output To Make Screen Easier for User to read.
		Write-Host NOT Found DistributionGroup TCQ-$global:CompanyPrefix-"$global:ClientName_NoSpace"PO_$global:DistListToCheck
		Write-Host Creating it ...
		Write-Host '' # Output To Make Screen Easier for User to read.
		Write-Host Running Command '"New-DistributionGroup -Name 'TCQ-$CompanyPrefix-"$ClientName_NoSpace"PO_$DistListToCheck' -DisplayName 'TCQ-$CompanyPrefix-"$ClientName_NoSpace"PO_$global:DistListToCheck' -Alias 'TCQ-$CompanyPrefix-"$ClientName_NoSpace"PO_$global:DistListToCheck' -OrganizationalUnit '"$OU"' | Out-Null"'
		New-DistributionGroup -Name TCQ-$CompanyPrefix-"$ClientName_NoSpace"PO_$global:DistListToCheck -DisplayName TCQ-$CompanyPrefix-"$ClientName_NoSpace"PO_$global:DistListToCheck -Alias TCQ-$CompanyPrefix-"$ClientName_NoSpace"PO_$global:DistListToCheck -OrganizationalUnit "$OU" | Out-Null
		Write-Host '' # Output To Make Screen Easier for User to read.
		Write-Host '-------------------------------------------------------------------------------------------------------------------' # Output To Make Screen Easier for User to read.
		#
		$global:AAD_Sync_Needed="YES"
		if ( $global:DistListToCheck -eq "D" )
		{
			$global:DistListCreatedDAY="Y"
		}
		if ( $global:DistListToCheck -eq "N" )
		{
		$global:DistListCreatedNIGHT="Y"
		}
	} 
}
#
###############################################################################################################
#
#   DISTRIBUTION LIST CHECKER FUNCTION
#
function DistListChecker
{
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host 'Checking For Distribution Lists and Creating Where Needed'
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host '-------------------------------------------------------------------------------------------------------------------' # Output To Make Screen Easier for User to read.
	#
	$global:DistListToCheck="D"
	DistListCheck
	$global:DistListToCheck="N"
	DistListCheck
}
#
function HasDistListBeenCreated
{
	if ( $DistListCreatedDAY -eq "Y" )
	{
		$global:DistListToCheck="D"
		$CheckType="Distribution List"
		$global:CheckTypePrefix="TCQ"
		AADsyncCheck-DL
	}
	#
	if ( $DistListCreatedNIGHT -eq "Y" )
	{
		$global:DistListToCheck="N"
		$CheckType="Distribution List"
		$global:CheckTypePrefix="TCQ"
		AADsyncCheck-DL
	}
}
#
###############################################################################################################
#
#   AZURE ACTIVE DIRECTORY SYNC FUNCTION
#
function AADsync 
{   
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host 'New Dist List/Lists have Been Created, AAD Sync is required'
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host 'Forcing AAD Sync'
	$ScriptBlock = [scriptblock]::create("$global:AAD_Sync_Server_Sync_Script")
	Invoke-Command -ComputerName $global:AAD_Sync_Server -ScriptBlock $ScriptBlock -Authentication Kerberos | Out-Null
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host '-------------------------------------------------------------------------------------------------------------------' # Output To Make Screen Easier for User to read.
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host 'New Dist List/Lists have Been Created and AAD Sync Run, checking List/Lists have Appeared in Azure AD'
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host '-------------------------------------------------------------------------------------------------------------------' # Output To Make Screen Easier for User to read.
}
#
###############################################################################################################
#
#   AZURE ACTIVE DIRECTORY SYNC CHECK DISTRIBUTION LIST CREATION FUNCTION
#
function AADsyncCheck-DL
{
	Write-Host '' # Output To Make Screen Easier for User to read.
	$Counter = 0
	do 
	{ 
		Write-Host Waiting For $CheckType $global:CheckTypePrefix-$CompanyPrefix-"$ClientName_NoSpace"PO_$global:DistListToCheck to Appear in Azure AD - $Counter Seconds Elapsed
		Start-Sleep -Seconds 10
		$Counter = $Counter + 10
	} until (Get-AzureADGroup -SearchString (Write-Output $global:CheckTypePrefix-$CompanyPrefix-"$ClientName_NoSpace"PO_$global:DistListToCheck) -ErrorAction SilentlyContinue)
}
#
###############################################################################################################
#
#   AZURE ACTIVE DIRECTORY SYNC CHECK RESOURCE ACCOUNT CREATION FUNCTION
#
function AADsyncCheck-RA
{
	Write-Host '' # Output To Make Screen Easier for User to read.
	$Counter = 0
	do 
	{ 
		Write-Host Waiting For Resource Account TRA-$CompanyPrefix-"$ClientName_NoSpace"PO-$ResAccType@$CompanyDomainNamePrefix to Appear in Azure AD - $Counter Seconds Elapsed
		Start-Sleep -Seconds 10
		$Counter = $Counter + 10
	} until (Get-CsOnlineUser (Write-Output TRA-$CompanyPrefix-"$ClientName_NoSpace"PO-$ResAccType@$CompanyDomainNamePrefix) -ErrorAction SilentlyContinue)
}
#
###############################################################################################################
#
#   AZURE ACTIVE DIRECTORY CREATE RESOURCE ACCOUNT FUNCTION
#
function CreateResourceAccount 
{
	Write-Host '' # Output To Make Screen Easier for User to read.
	#
	if ( $global:ResourceAccountToCreate -eq "CQD" )
	{
		$global:AppID="11cd3e2e-fccb-42ad-ad00-878b93575e07"
	}
	#
	if ( $global:ResourceAccountToCreate -eq "CQN" )
	{
		$global:AppID=“11cd3e2e-fccb-42ad-ad00-878b93575e07"
	}
	#
	if ( $global:ResourceAccountToCreate -eq "AA" )
	{
		$global:AppID=“ce933385-9390-45d1-9512-c8d228074e07"
	}
	#
	Write-Host 'Creating Resource Account'
	Write-Host TRA-$CompanyPrefix-"$ClientName_NoSpace"PO-$global:ResourceAccountToCreate@$CompanyDomainNamePrefix
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host Running Command '"New-CsOnlineApplicationInstance -UserPrincipalName'(Write-Output TRA-$CompanyPrefix-"$ClientName_NoSpace"PO-$global:ResourceAccountToCreate@$CompanyDomainNamePrefix)'-DisplayName'(Write-Output "$ClientName PO – $global:ResourceAccountType")' -ApplicationId '"$global:AppID"'| Out-Null"'
	New-CsOnlineApplicationInstance -UserPrincipalName (Write-Output TRA-$CompanyPrefix-"$ClientName_NoSpace"PO-$global:ResourceAccountToCreate@$CompanyDomainNamePrefix) -DisplayName (Write-Output "$ClientName PO – $global:ResourceAccountType") -ApplicationId "$global:AppID" | Out-Null
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host '-------------------------------------------------------------------------------------------------------------------' # Output To Make Screen Easier for User to read.
}
#
###############################################################################################################
#
#   AZURE ACTIVE DIRECTORY RESOURCE ACCOUNTS FUNCTION
#
function ResourceAccounts 
{
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host '-------------------------------------------------------------------------------------------------------------------' # Output To Make Screen Easier for User to read.
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host 'Creating Resource Accounts'
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host '-------------------------------------------------------------------------------------------------------------------' # Output To Make Screen Easier for User to read.
	#
	$global:ResourceAccountToCreate="CQD"
	$global:ResourceAccountType="Call Queue Day"
	CreateResourceAccount
	$global:ResourceAccountToCreate="CQN"
	$global:ResourceAccountType="Call Queue Night"
	CreateResourceAccount
	$global:ResourceAccountToCreate="AA"
	$global:ResourceAccountType="Auto Attendant"
	CreateResourceAccount
	#
	$ResAccType="AA"
	AADsyncCheck-RA
	#
	# Grab The Azure Object ID of the Resource Account
	$global:AAresourceAccount=(Get-AzureADUser -ObjectId TRA-$CompanyPrefix-"$ClientName_NoSpace"PO-$global:ResourceAccountToCreate@$CompanyDomainNamePrefix).ObjectId
	#
	# Grab The Azure Object ID of the Group Assigning The Relevant Licenses 
	$global:AzureADgroupAdd=(Get-AzureADGroup -SearchString $global:AzureADlicenseGroup | Where-Object { $_.DisplayName.EndsWith($global:AzureADlicenseGroup) }).ObjectId
	#
	# Add The User To The Group So it gets a License
	Add-AzureADGroupMember -ObjectId $global:AzureADgroupAdd -RefObjectId $global:AAresourceAccount
}
#
###############################################################################################################
#
#   TEAMS CREATE CALL QUEUES FUNCTION
#
function CreateCallQueue 
{
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host Creating Call Queue (Write-Output "$ClientName PO – Call Queue $global:DistListType")
	$CallQueueDistList = (Get-AzureADGroup -SearchString (Write-Output TCQ-$CompanyPrefix-"$ClientName_NoSpace"PO_$global:DistListToCreate)).ObjectId
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host Running Command '"New-CsCallQueue -Name'(Write-Output "$ClientName PO – Call Queue $global:DistListType")'-Tenant 7768bca1-9fc1-4f38-b39e-e75e9aebb498 -RoutingMethod Attendant -DistributionLists'$CallQueueDistList' -AllowOptOut $False -ConferenceMode $True -PresenceBasedRouting $False -AgentAlertTime 30 -LanguageId en-GB -OverflowThreshold 50 -OverflowAction DisconnectWithBusy -EnableOverflowSharedVoicemailTranscription $True -TimeoutThreshold 30 -TimeoutAction SharedVoiceMail -TimeoutActionTarget '$CallQueueDistList'-TimeoutSharedVoicemailTextToSpeechPrompt "We’re sorry but we have not been able to connect your call.
	Write-Host 'Please leave a message and a member of the Team will get back to you as soon as possible." -EnableTimeoutSharedVoicemailTranscription $True -UseDefaultMusicOnHold $True -WarningAction SilentlyContinue | Out-Null"'
	#
	New-CsCallQueue -Name (Write-Output "$ClientName PO – Call Queue $global:DistListType") -Tenant 7768bca1-9fc1-4f38-b39e-e75e9aebb498 -RoutingMethod Attendant -DistributionLists $CallQueueDistList -AllowOptOut $False -ConferenceMode $True -PresenceBasedRouting $False -AgentAlertTime 30 -LanguageId en-GB -OverflowThreshold 50 -OverflowAction DisconnectWithBusy -EnableOverflowSharedVoicemailTranscription $True -TimeoutThreshold 30 -TimeoutAction SharedVoiceMail -TimeoutActionTarget $CallQueueDistList -TimeoutSharedVoicemailTextToSpeechPrompt "We’re sorry but we have not been able to connect your call.
	Please leave a message and a member of the Team will get back to you as soon as possible." -EnableTimeoutSharedVoicemailTranscription $True -UseDefaultMusicOnHold $True -WarningAction SilentlyContinue | Out-Null
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host '-------------------------------------------------------------------------------------------------------------------' # Output To Make Screen Easier for User to read.
	# 
	AADsyncCheck-RA
	#
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host '-------------------------------------------------------------------------------------------------------------------' # Output To Make Screen Easier for User to read.
	Write-Host '' # Output To Make Screen Easier for User to read.
	#
	$ResourceAccountId = (Get-CsOnlineUser (Write-Output TRA-$CompanyPrefix-"$ClientName_NoSpace"PO-$ResAccType@$CompanyDomainNamePrefix)).ObjectId
	$CallQueue = (Get-CsCallQueue -NameFilter (Write-Output "$ClientName PO – Call Queue $global:DistListType") -WarningAction SilentlyContinue).Identity
	Write-Host Associating Call Queue '"'$ClientName PO – Call Queue $global:DistListType '"' To Resource Account '"'TRA-$CompanyPrefix-"$ClientName_NoSpace"PO-$ResAccType@$CompanyDomainNamePrefix '"'
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host Running Command '"New-CsOnlineApplicationInstanceAssociation -Identities'@($ResourceAccountId) '-ConfigurationId'$CallQueue'-ConfigurationType CallQueue"'
	New-CsOnlineApplicationInstanceAssociation -Identities @($ResourceAccountId) -ConfigurationId $CallQueue -ConfigurationType CallQueue | Out-Null
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host '-------------------------------------------------------------------------------------------------------------------' # Output To Make Screen Easier for User to read.
}
#
###############################################################################################################
#
#   TEAMS CALL QUEUES FUNCTION
#
function CallQueues
{
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host '-------------------------------------------------------------------------------------------------------------------' # Output To Make Screen Easier for User to read.
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host 'Creating Call Queues'
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host '-------------------------------------------------------------------------------------------------------------------' # Output To Make Screen Easier for User to read.
	#
	$global:DistListToCreate="D"
	$global:DistListType="Day"
	$ResAccType="CQD"
	CreateCallQueue
	#
	$global:DistListToCreate="N"
	$global:DistListType="Night"
	$ResAccType="CQN"
	CreateCallQueue
}
#
###############################################################################################################
#
#   TEAMS CREATE AUTO ATTENDANT FUNCTION
#
function AutoAttendant
{
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host 'Setting Up Auto Attendant'
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host '-------------------------------------------------------------------------------------------------------------------' # Output To Make Screen Easier for User to read.
	Write-Host '' # Output To Make Screen Easier for User to read.
	#
	$AutoAttendantName = (Write-Output "$ClientName PO – Auto Attendant")
	$GroupIds = Find-CsGroup -SearchQuery $global:GALsearchScope | % { $_.Id }
	$DialScope = New-CsAutoAttendantDialScope -GroupScope -GroupIds $groupIds
	#
	Write-Host 'Setting Call Queue Routing Targets'
	Write-Host '' # Output To Make Screen Easier for User to read.
	$daytargetCQappinstance = (Get-CsOnlineUser -Identity (Write-Output TRA-$CompanyPrefix-"$ClientName_NoSpace"PO-CQD@$CompanyDomainNamePrefix)).ObjectId 
	$nighttargetCQappinstance = (Get-CsOnlineUser -Identity (Write-Output TRA-$CompanyPrefix-"$ClientName_NoSpace"PO-CQN@$CompanyDomainNamePrefix)).ObjectId 
	#
	Write-Host 'Setting Business Hours Menu Options'
	Write-Host '' # Output To Make Screen Easier for User to read.
	$daytarget = New-CsAutoAttendantCallableEntity -Identity $daytargetCQappinstance -Type applicationendpoint
	$automaticMenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Automatic -CallTarget $daytarget
	#
	Write-Host 'Setting After Hours Menu Options'
	Write-Host '' # Output To Make Screen Easier for User to read.
	$nighttarget = New-CsAutoAttendantCallableEntity -Identity $nighttargetCQappinstance -Type applicationendpoint
	$afterHoursMenuOption = New-CsAutoAttendantMenuOption -Action TransferCallToTarget -DtmfResponse Automatic -CallTarget $nighttarget
	#
	Write-Host 'Setting Greetings Prompts'
	Write-Host '' # Output To Make Screen Easier for User to read.
	$greetingPrompt = New-CsAutoAttendantPrompt -TextToSpeechPrompt $global:GreetingText
	#
	Write-Host 'Setting Up Business Hours Menu'
	Write-Host '' # Output To Make Screen Easier for User to read.
	$BusinessHoursMenu = New-CsAutoAttendantMenu -Name "AA menu2" -EnableDialByName -MenuOptions @($automaticMenuOption)
	#
	Write-Host 'Setting Up Business Hours Call Flow'
	Write-Host '' # Output To Make Screen Easier for User to read.
	$BusinessHoursCallFlow = New-CsAutoAttendantCallFlow -Name "Default" -Menu $BusinessHoursMenu -Greetings $greetingPrompt
	#
	Write-Host 'Setting Up After Hours Menu'
	Write-Host '' # Output To Make Screen Easier for User to read.
	$tr1 = New-CsOnlineTimeRange -Start 09:00 -End 17:30
	$afterHoursSchedule = New-CsOnlineSchedule -Name (Write-Output "After Hours - $ClientName PO - Auto Attendant") -WeeklyRecurrentSchedule -MondayHours @($tr1) -TuesdayHours @($tr1) -WednesdayHours @($tr1) -ThursdayHours @($tr1) -FridayHours @($tr1) -Complement
	$afterHoursMenu = New-CsAutoAttendantMenu -Name "AA menu1" -MenuOptions @($afterHoursMenuOption)
	#
	Write-Host 'Setting Up After Hours Call Flow and Call Handling Associations'
	Write-Host '' # Output To Make Screen Easier for User to read.
	$AfterHoursCallFlow = New-CsAutoAttendantCallFlow -Name "After Hours" -Menu $afterHoursMenu -Greetings @($greetingPrompt)
	$AfterHoursCallHandlingAssociation = New-CsAutoAttendantCallHandlingAssociation -Type AfterHours -ScheduleId $afterHoursSchedule.Id -CallFlowId $afterHoursCallFlow.Id
	#
	Write-Host '-------------------------------------------------------------------------------------------------------------------' # Output To Make Screen Easier for User to read.
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host 'Setting Up Holidays Call Flow and Call Handling Associations'
	Write-Host '' # Output To Make Screen Easier for User to read.
	#
	Holidays
	#
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host '-------------------------------------------------------------------------------------------------------------------' # Output To Make Screen Easier for User to read.
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host 'Creating AutoAttendant'
	#
	New-CsAutoAttendant -Name $AutoAttendantName -LanguageId $global:Language -DefaultCallFlow $BusinessHoursCallFlow -CallFlows @($global:CallFlows) -TimeZoneId $global:Timezone -Operator $operator -CallHandlingAssociations @($global:CallHandlingAssociations) -InclusionScope $DialScope | Out-Null
	#
	$ResAccType="AA"
	AADsyncCheck-RA
	#
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host '-------------------------------------------------------------------------------------------------------------------' # Output To Make Screen Easier for User to read.
	Write-Host '' # Output To Make Screen Easier for User to read.
	#
	$ResourceAccountId = (Get-CsOnlineUser (Write-Output TRA-$CompanyPrefix-"$ClientName_NoSpace"PO-$ResAccType@$CompanyDomainNamePrefix)).ObjectId
	$AutoAttendant = (Get-CsAutoAttendant -NameFilter (Write-Output "$ClientName PO – Auto Attendant") -WarningAction SilentlyContinue).Identity
	Write-Host Associating Auto Attendant "$ClientName PO – Auto Attendant" To Resource Account '"'TRA-$CompanyPrefix-"$ClientName_NoSpace"PO-$ResAccType@$CompanyDomainNamePrefix '"'
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host Running Command '"New-CsOnlineApplicationInstanceAssociation -Identities'@($ResourceAccountId) '-ConfigurationId'$AutoAttendant'-ConfigurationType AutoAttendant"'
	New-CsOnlineApplicationInstanceAssociation -Identities @($ResourceAccountId) -ConfigurationId $AutoAttendant -ConfigurationType AutoAttendant | Out-Null
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host '-------------------------------------------------------------------------------------------------------------------' # Output To Make Screen Easier for User to read.
	Write-Host '' # Output To Make Screen Easier for User to read.
	#
	$Counter = 0
	do 
	{ 
		Write-Host Waiting For Resource Account TRA-$CompanyPrefix-"$ClientName_NoSpace"PO-AA@$CompanyDomainNamePrefix to Report A Phone System License in Azure AD - $Counter Seconds Elapsed
		Start-Sleep -Seconds 10
		$Counter = $Counter + 10
	} until ((Get-MsolUser -ErrorAction SilentlyContinue –UserPrincipalName (Write-Output TRA-$CompanyPrefix-"$ClientName_NoSpace"PO-AA@$CompanyDomainNamePrefix)).Licenses[0].ServiceStatus | SELECT-OBJECT "MCOV")
	#
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host '-------------------------------------------------------------------------------------------------------------------' # Output To Make Screen Easier for User to read.
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host 'Assigning Phone Number to Resource Account'
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host Running Command '"Set-CsOnlineApplicationInstance -Identity (Write-Output 'TRA-$CompanyPrefix-"$ClientName_NoSpace"PO-AA@$CompanyDomainNamePrefix') -OnpremPhoneNumber '$OnpremPhoneNumber'"'
	Set-CsOnlineApplicationInstance -Identity (Write-Output TRA-$CompanyPrefix-"$ClientName_NoSpace"PO-AA@$CompanyDomainNamePrefix) -OnpremPhoneNumber $OnpremPhoneNumber
	Write-Host '' # Output To Make Screen Easier for User to read.
	Write-Host '-------------------------------------------------------------------------------------------------------------------' # Output To Make Screen Easier for User to read.
}
#
###############################################################################################################
#
#   END OF FUNCTION DEFENITION
#
###############################################################################################################
#
#   BEGIN PROCESSING
#
###############################################################################################################
#
Write-Host '' # Output To Make Screen Easier for User to read.
#
CustomisableVariables
#
Logging
#
CollectVariables
#
Connections
#
DistListChecker
#
if ( $AAD_Sync_Needed -eq "YES" )
{
	AADsync
}
#
HasDistListBeenCreated
#
ResourceAccounts
#
CallQueues
#
AutoAttendant
#
End
