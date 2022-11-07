

Connect-MicrosoftTeams
Connect-ExchangeOnline
Connect-AzureAD
Connect-MsolService


# #############################Teams Calling Config ######################################################################################################################
## Call Queue Name ##
$cqName = 'Call_Queue_Main'
##Auto Attendant Name
$aaName = "Main"
## Auto Attendant Dial-in Number ##
$AutoAttendantPhoneNo = "+61xxxxxxxxx"
## 365 Tenancy domain name (use the FQDN not default onmicrosoft) ##
$resourceAccountDomain = 'tenancy.onmicrosoft.com'
## Prompt used for all voicemail for call time-outs and overflow ##
$voicemailprompt = "Thank you for calling, All our staff are currently unavailable, please leave a message and we will return your call as soon as possible"
## Auto Attendant Business Hours Greeting ##
$aaGreeting = "Thank you for calling, please hold the line and your call will be answered by the next available staff member"
## After Hours Voicemail Greeting ##
$vmGreeting = "Thank you for calling, you called outside business hours"
## Timezone - Supported timezones can be found by running command "Get-CsAutoAttendantSupportedTimeZone" ##
$timezone = "E. Australia Standard Time"
## Greeting voice Gender  ##
$gender = "Female"
##Business hours Mon-Fri
$starttime = "7:30"
$endtime = "20:00"
########################
$aaLanguage = 'en-AU'
$aaTimezone = $timezone
##########################################################################################################################################################################

################################### Create 365 Call queue User Group   #############################
cls
## Create Voicemail Group/Call Queue group and assign the Guid to $voicemailguid
Write-host "Creating Call Queue Unified Group"
New-UnifiedGroup -DisplayName "$($cqName)" -Alias "$($cqName)" ` -EmailAddresses "$($cqName)@$($resourceAccountDomain)" -AccessType Public
$voicemailguid = Get-AzureADGroup | Where-Object {$_.DisplayName -eq "$($cqName)"} | ForEach-Object {$_.ObjectId}


sleep 20



################################### Call Queue   ########################################################################################################################

# Create resource account of call queue type
$cqRaParams = @{
	UserPrincipalName = "RA_CQ_$cqName@$resourceAccountDomain"
	ApplicationId = '11cd3e2e-fccb-42ad-ad00-878b93575e07'
	DisplayName = "RA_CQ_$cqName"
    

}
$newCqRa = New-CsOnlineApplicationInstance @cqRaParams
Write-host "Creating Call Queue Resource Group"

sleep 20

################   Licence Call Queue Resource Account####################################################################################################################

$licence = Get-MsolAccountSku | Where-Object {$_.AccountSkuId -match "PHONESYSTEM_VIRTUALUSER"} | ForEach-Object {$_.AccountSkuId}
Write-host "Setting Call Queue Users Usage Location"
Set-MsolUser -UserPrincipalName "RA_CQ_$cqName@$resourceAccountDomain" -UsageLocation AU
sleep 60
Write-host "Assigning license to Call Queue Resource Group"
Set-MsolUserLicense -UserPrincipalName "RA_CQ_$cqName@$resourceAccountDomain" -AddLicenses "$licence" 
sleep 20

##################### Create call queue and configure ########################################################################################

Write-host -background Black "Building Call Queue parameters"


$newCqParams = @{
    #General Info
	Name = "CQ_$($cqName)"
    LanguageId = "en-AU"
    #Greetings and Music
	UseDefaultMusicOnHold = $true
    #Call Answering
    ConferenceMode = $True
    DistributionLists = "$voicemailguid"
    #Agent Selection
    RoutingMethod = "Attendant"
    AgentAlertTime = "60"
    AllowOptOut = $False
    PresenceBasedRouting = $False
    #Call Overflow Handling
    OverflowAction = "SharedVoicemail"
    OverflowActionTarget = "$voicemailguid"
    OverflowSharedVoicemailTextToSpeechPrompt = $voicemailprompt
    EnableOverflowSharedVoicemailTranscription = $true
    OverflowThreshold = "5"
    #Call timeout Handling
    TimeoutAction = "SharedVoicemail"
    TimeoutActionTarget = "$voicemailguid"
    TimeoutSharedVoicemailTextToSpeechPrompt = "$voicemailprompt"
    TimeoutThreshold = "120"
    EnableTimeoutSharedVoicemailTranscription = $True

}
Write-host "Creating call Queue"
$newCq = New-CsCallQueue @newCqParams
Write-host -background Black "Call queue Created successfully"
sleep 20

# Associate resource account with call queue
$newCqAppInstanceParams = @{
	# Requires array of strings
	# Use array sub-expression operator
	Identities = @($newCqRa.ObjectId)
	ConfigurationId = $newCq.Identity
	ConfigurationType = 'CallQueue'
	ErrorAction = 'Stop'
}
Write-host "Associate resource account with call queue"
$associationRes = New-CsOnlineApplicationInstanceAssociation @newCqAppInstanceParams


################################## Auto Attendant Resource Account Setup  ############################################################################################

#Create Resource Account / License & Assign Number
$newAaRaParams = @{
	UserPrincipalName = "RA_AA_$aaName@$resourceAccountDomain"
	ApplicationId = 'ce933385-9390-45d1-9512-c8d228074e07'
	DisplayName = "RA_AA_$aaName"
}
Write-host "Creating AA Resource Account"
$newAaRa = New-CsOnlineApplicationInstance @newAaRaParams
sleep 20

$licence = Get-MsolAccountSku | Where-Object {$_.AccountSkuId -match "PHONESYSTEM_VIRTUALUSER"} | ForEach-Object {$_.AccountSkuId}
Write-host "Setting Auto Attendant User Usage Location"
Set-MsolUser -UserPrincipalName "RA_AA_$aaName@$resourceAccountDomain" -UsageLocation AU
sleep 60
Write-host "Setting Auto Attendant Licence"
Set-MsolUserLicense -UserPrincipalName "RA_AA_$aaName@$resourceAccountDomain" -AddLicenses "$licence" 

sleep 120
Write-host "Assigning Auto Attendant Resource Account Phone Number"
Set-CsPhoneNumberAssignment -Identity "RA_AA_$aaName@$resourceAccountDomain" -PhoneNumber $AutoAttendantPhoneNo -PhoneNumberType CallingPlan


################################## Auto Attendant Configuration ##################################################

Write-host -background Black "Building Auto Attendant parameters"



# Callable entity
$callableEntityParams = @{
	# Point to resource account, not call queue
	Identity = $newCqRa.ObjectId
	Type = 'ApplicationEndpoint'
}
$targetCqEntity = New-CsAutoAttendantCallableEntity @callableEntityParams


# AA Menu options
$menuOptionParams = @{
	Action = 'TransferCallToTarget'
	DtmfResponse = 'Automatic'
	CallTarget = $targetCqEntity

}
$menuOptionZero = New-CsAutoAttendantMenuOption @menuOptionParams

# AA Menu
$menuParams = @{
	Name = "$aaName Default Menu"
	# Accepts list, so use array sub-expression operator
	MenuOptions = @($menuOptionZero)    
}
$menu = New-CsAutoAttendantMenu @menuParams

#Auto Attendant GreetingParams
$greetingParams = @{
    TextToSpeechPrompt = "$aaGreeting"
    
}
$greetingPrompt = New-CsAutoAttendantPrompt @greetingParams

# And the call flow
$defaultCallFlowParams = @{
	Name = "RA_AA_$aaName Default Call Flow"
    Greetings = @($greetingPrompt)
	Menu = $menu
    
}
$defaultCallFlow = New-CsAutoAttendantCallFlow @defaultCallFlowParams
Write-host "Created Default Auto Attendant greeting"


# Callable entity
$vmcallableEntityParams = @{
	# Point to resource account, not call queue
	Identity = $voicemailguid
	Type = 'SharedVoicemail'
}
$targetvmEntity = New-CsAutoAttendantCallableEntity @vmcallableEntityParams


# AA Menu options
$menuOptionvmParams = @{
	Action = 'TransferCallToTarget'
	DtmfResponse = 'Automatic'
	CallTarget = $targetvmEntity

}
$menuOptionvm = New-CsAutoAttendantMenuOption @menuOptionvmParams

# AA Menu
$menuvmParams = @{
	Name = "Voicemail Menu"
	# Accepts list, so use array sub-expression operator
	MenuOptions = @($menuOptionvm)    
}
$menuvm = New-CsAutoAttendantMenu @menuvmParams

#Voicemail Auto Attendant GreetingParams
$vmgreetingParams = @{
    TextToSpeechPrompt = "$vmGreeting"
    
}
$vmgreetingPrompt = New-CsAutoAttendantPrompt @vmgreetingParams

# And the call flow
$vmCallFlowParams = @{
	Name = "$Voicemail Call Flow"
    Greetings = @($vmgreetingPrompt)
	Menu = $menuvm
    
}
$vmCallFlow = New-CsAutoAttendantCallFlow @vmCallFlowParams
Write-host "Created Voicemail Call flow"

#### Setup formatted properly

# AfterHours Time Range
$timerangeMoFrParams = @{
	Start = "$starttime"
    end = "$endtime"
}
$timerangeMoFr = New-CsOnlineTimeRange @timerangeMoFrParams

# AfterHours Schedule
$afterHoursScheduleParams = @{
    Name = "After Hours Schedule"
	MondayHours = @($timerangeMoFr)
    TuesdayHours = @($timerangeMoFr)
    WednesdayHours = @($timerangeMoFr) 
    ThursdayHours = @($timerangeMoFr)
    FridayHours = @($timerangeMoFr)

}
$afterHoursSchedule = New-CsOnlineSchedule -WeeklyRecurrentSchedule @afterHoursScheduleParams -Complement

# AfterHours Schedule
$afterHoursCallHandlingAssociationParams = @{
    Type = "AfterHours"
    ScheduleId = $afterHoursSchedule.Id
    CallFlowId = $vmCallFlow.Id

}
$afterHoursCallHandlingAssociation = New-CsAutoAttendantCallHandlingAssociation @afterHoursCallHandlingAssociationParams
Write-host "Created After Hours Schedule"

# You have all the objects
# Creating an auto attendant
$autoAttendantParams = @{
	Name = "$aaName Auto Attendant"
	LanguageId = $aaLanguage
	TimeZoneId = $timezone
	DefaultCallFlow = $defaultCallFlow
    CallFlows = $vmCallFlow
    CallHandlingAssociations = $afterHoursCallHandlingAssociation
    VoiceId = "$gender"
	ErrorAction = 'Stop'
}
$newAA = New-CsAutoAttendant @autoAttendantParams

Write-host "Creating $aaName Auto Attendant"

# Resource account association to Auto Attendant
$aaAssociationParams = @{
	# As the previous association, array expected
	Identities = @($newAARA.ObjectId)
	ConfigurationId = $newAA.Identity
	ConfigurationType = 'AutoAttendant'
	ErrorAction = 'Stop'
}

Write-host "Assigning $($newAARA.DisplayName) resource account to $aaName Auto Attendant"
$associationRes = New-CsOnlineApplicationInstanceAssociation @aaAssociationParams




