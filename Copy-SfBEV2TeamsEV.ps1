<#
	.SYNOPSIS
		A script to copy a Skype for Business on-prem Enterprise Voice configuration to a Teams Enterprise Voice configuration.

	.DESCRIPTION
		A script to automatically copy a Skype for Business on-prem Enterprise Voice configuration to a Teams Enterprise Voice configuration. Will copy the following items:
		- Dialplans and associated normalization rules
		- Voice routes
		- Voice policies
		- PSTN usages
		- Outbound translation rules
		
		The script must be run from an on-prem Skype for Business server.

		User running the script must have the following roles at minimum:
		- at least CSViewOnlyAdministrator rights in SfB
		- full admin rights in Teams

	.PARAMETER KeepExisting
		OPTIONAL. Will not erase existing Teams Enterprise Voice configuration before restoring.

	.PARAMETER OverrideAdminDomain
		OPTIONAL: The FQDN your Office365 tenant. Use if your admin account is not in the same domain as your tenant (ie. doesn't use a @tenantname.onmicrosoft.com address)
		
	.NOTES
		Version 1.00
		Build: Feb 06, 2020
		Copyright © 2020  Ken Lasko
		klasko@ucdialplans.com
		https://www.ucdialplans.com
#>

[CmdletBinding(ConfirmImpact = 'Medium',
SupportsShouldProcess)]
param (
	[switch]
	$KeepExisting,
	[string]
	$OverrideAdminDomain
)

# Make sure there isn't an active PSSession to O365, because the O365 Get-CsDialplan command overrides the on-prem one
Get-PSSession | Where {$_.ComputerName -like '*online.lync.com'} | Remove-PSSession

# Closing a broken PSSession often leaves the O365 PS commands in the session. If it is, forcibly remove them
$PSSource = (Get-Command Get-CsDialplan).Source
If ($PSSource -ne 'SkypeForBusiness') {
	Remove-Module -Name $PSSource
}

Write-Host -Object 'Getting Skype for Business Enterprise Voice details'
Write-Verbose -Message 'Getting SfB dialplans'
$Dialplans = Get-CsDialplan
Write-Verbose -Message 'Getting SfB voice routes'
$VoiceRoutes = Get-CsVoiceRoute
Write-Verbose -Message 'Getting SfB PSTN usages'
$PSTNUsages = (Get-CsPSTNUsage).Usage
Write-Verbose -Message 'Getting SfB voice policies'
$VoicePolicies = Get-CsVoicePolicy
Write-Verbose -Message 'Getting SfB PSTN gateways'
$PSTNGateways = (Get-CsTrunk).PoolFQDN
Write-Verbose -Message 'Getting SfB outbound calling number translation rules'
$OutboundCallingTransRules = Get-CsOutboundCallingNumberTranslationRule
Write-Verbose -Message 'Getting SfB outbound called number translation rules'
$OutboundCalledTransRules = Get-CsOutboundTranslationRule

# Connect to O365
Write-Host -Object 'Logging into Office 365...'

If ($OverrideAdminDomain) {
	$O365Session = (New-CsOnlineSession -OverrideAdminDomain $OverrideAdminDomain)
}
Else {
	$O365Session = (New-CsOnlineSession)
}
$null = (Import-PSSession -Session $O365Session -AllowClobber)

# Try to find matching PSTN Gateways in Teams
$PSTNGWMatch = @{}
$TeamsPSTNGW = (Get-CsOnlinePSTNGateway).Identity

# For each SfB PSTN gateway, try to match against an existing Teams gateway with the same name. If not, then allow to manually match or create one.
ForEach ($PSTNGateway in $PSTNGateways) {
	$ValidSelection = $FALSE
	Write-Host
	Do {
		If (Get-CsOnlinePSTNGateway -Identity $PSTNGateway -ErrorAction SilentlyContinue) {
			Write-Host -Object 'Found a matching Teams PSTN gateway for ' -NoNewLine
			Write-Host $PSTNGateway -ForegroundColor Yellow
			$PSTNGWMatch.Add($PSTNGateway, $PSTNGateway)
			$ValidSelection = $TRUE
		}
		Else {
			Write-Host
			Write-Host 'Could not find a matching PSTN gateway for ' -NoNewLine
			Write-Host $PSTNGateway -ForegroundColor Yellow -NoNewLine
			Write-Host '.'
			Write-Host 'Please select an existing Teams PSTN gateway from the below list, or opt to create one:'
			$TeamsPSTNGWList = @()
			Write-Host
			Write-Host '#     Teams PSTN Gateway'
			Write-Host '=     ==================='
			For ($i=0; $i -lt $TeamsPSTNGW.Count; $i++) {
				$a = $i + 1
				Write-Host ($a, $TeamsPSTNGW[$i]) -Separator '     '
			}
			
			If ($PSTNGateway -match "[A-Za-z]") { # If there are no letters in the gateway name, its not going to be a valid FQDN (likely an IP address), so don't offer to create a Teams GW with that name
				$a = $a + 1 
				Write-Host ($a, "Create a PSTN gateway called $PSTNGateway") -Separator '     '
			}
			
			$a = $a + 1 
			Write-Host ($a, 'Create a new PSTN gateway') -Separator '     '
			$a = $a + 1 
			Write-Host ($a, "Don't attempt match") -Separator '     '

			$Range = '(1-' + $a + ')'
			Write-Host
			$Select = Read-Host "Select a Teams PSTN gateway to match with the SfB GW $PSTNGateway" $Range

			If (($Select -gt $a) -or ($Select -lt 1)) {
				Write-Host 'Invalid selection' -ForegroundColor Red
				$ValidSelection = $FALSE
			}
			ElseIf ($Select -lt $a-2) { # Use an existing Teams gateway to replace the SfB gateway
				$PSTNGWMatch.Add($PSTNGateway, $TeamsPSTNGW[$Select-1])
				$ValidSelection = $TRUE
			}
			ElseIf ($Select -eq $a-2) { # Create a Teams gateway using the same name as the existing SfB gateway
				Try {
					New-CsOnlinePSTNGateway -Identity $PSTNGateway -ErrorAction Stop
					$PSTNGWMatch.Add($PSTNGateway, $PSTNGateway)
					$ValidSelection = $TRUE
				}
				Catch {
					Write-Host "Could not create $PSTNGateway. This is likely because the selected domain is not configured for this tenant." -ForegroundColor Red
					$ValidSelection = $FALSE
				}
			}
			ElseIf ($Select -eq $a-1) { # Create a new Teams gateway using a new name to replace the existing SfB gateway
				Try {
					$NewPSTNGateway = New-CsOnlinePSTNGateway -ErrorAction Stop
					$PSTNGWMatch.Add($PSTNGateway, $NewPSTNGateway.Identity)
					$ValidSelection = $TRUE
				}
				Catch {
					Write-Host 'Could not create new PSTN gateway. This is likely because the selected domain is not configured for this tenant, or the gateway name already exists.' -ForegroundColor Red
					$ValidSelection = $FALSE
				}
			}
			ElseIf ($Select -eq $a) { # Don't select a matching gateway
				$PSTNGWMatch.Add($PSTNGateway, $NULL)
				$ValidSelection = $TRUE
			}
		}
	} Until ($ValidSelection -eq $TRUE)
}

# Show a summary of the SfB gateways and the chosen Teams gateway match
Write-Host
Write-Host 'List of SfB PSTN Gateways (Name) and the matching Teams PSTN Gateways (Value)'
Write-Host '============================================================================='
$PSTNGWMatch
Write-Host

# Erase all existing Teams Enterprise Voice information unless the user started the script with the -KeepExisting switch
If (!$KeepExisting) {
	Write-Host -ForegroundColor Yellow 'WARNING: About to ERASE all existing Teams dialplans/voice routes/policies etc prior to copying from SfB.'
	Write-Host -ForegroundColor Yellow 'This does NOT include any existing PSTN gateways, but any translation rules will be removed.'
	Write-Host -ForegroundColor Yellow 'If you want to keep the existing configuration, restart the script and add the -KeepExisting switch.'
	$Confirm = Read-Host -Prompt 'Continue (Y/N)?'
	If ($Confirm -notmatch '^[Yy]$') {
		Write-Host -Object 'Exiting without making changes.'
		Exit
	}
	
	Write-Host -Object 'Erasing all existing dialplans/voice routes/policies etc from Teams'
	
	Write-Verbose 'Erasing all Teams dialplans'
	$null = (Get-CsTenantDialPlan -ErrorAction SilentlyContinue | Remove-CsTenantDialPlan -ErrorAction SilentlyContinue)
	Write-Verbose 'Erasing all Teams voice routes'
	$null = (Get-CsOnlineVoiceRoute -ErrorAction SilentlyContinue | Remove-CsOnlineVoiceRoute -ErrorAction SilentlyContinue)
	Write-Verbose 'Erasing all Teams voice routing policies'
	$null = (Get-CsOnlineVoiceRoutingPolicy -ErrorAction SilentlyContinue | Remove-CsOnlineVoiceRoutingPolicy -ErrorAction SilentlyContinue)
	Write-Verbose 'Erasing all Teams PSTN usages'
	$null = (Set-CsOnlinePstnUsage -Identity Global -Usage $NULL -ErrorAction SilentlyContinue)
	Write-Verbose 'Removing all Teams translation rules from existing gateways'
	$null = (Get-CsOnlinePSTNGateway -ErrorAction SilentlyContinue | Set-CsOnlinePSTNGateway -OutbundTeamsNumberTranslationRules $NULL -OutboundPstnNumberTranslationRules $NULL -ErrorAction SilentlyContinue)
	Write-Verbose 'Erasing all Teams translation rules'
	$null = (Get-CsTeamsTranslationRule -ErrorAction SilentlyContinue | Remove-CsTeamsTranslationRule -ErrorAction SilentlyContinue)
}

Write-Host -Object 'Copying dialplans from SfB on-prem to Teams'
ForEach ($Dialplan in $Dialplans) {
	Write-Verbose "Copying $($Dialplan.Identity) dialplan"
	$DPExists = (Get-CsTenantDialPlan -Identity $Dialplan.Identity -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Identity)

	$DPDetails = @{
		Identity = $Dialplan.Identity -replace "^Site\:", "" # Since site-level dialplans aren't supported in Teams, site-level dialplans will be converted to user-level dialplans
		OptimizeDeviceDialing = $Dialplan.OptimizeDeviceDialing
		Description = $Dialplan.Description
		NormalizationRules = $Dialplan.NormalizationRules
	}

	# Only include the external access prefix if one is defined. MS throws an error if you pass a null/empty ExternalAccessPrefix
	If ($Dialplan.ExternalAccessPrefix) {
		$DPDetails.Add('ExternalAccessPrefix', $Dialplan.ExternalAccessPrefix)
	}

	If ($DPExists) {
		$null = (Set-CsTenantDialPlan @DPDetails)
	}
	Else {
		$null = (New-CsTenantDialPlan @DPDetails)
	}
}

# Copy PSTN usages from SfB on-prem to MS Teams
Write-Host -Object 'Copying PSTN usages from SfB on-prem to Teams'

ForEach ($PSTNUsage in $PSTNUsages) {
	Write-Verbose "Copying $PSTNUsage PSTN usage"
	$null = (Set-CsOnlinePstnUsage -Identity Global -Usage @{Add = $PSTNUsage} -WarningAction SilentlyContinue -ErrorAction SilentlyContinue)
}

# Copy voice routes from SfB on-prem to MS Teams
Write-Host -Object 'Copying voice routes from SfB on-prem to Teams'

ForEach ($VoiceRoute in $VoiceRoutes) {
	Write-Verbose "Copying $($VoiceRoute.Identity) voice route"
	$VRExists = (Get-CsOnlineVoiceRoute -Identity $VoiceRoute.Identity -ErrorAction SilentlyContinue).Identity

	# Replace SfB gateways with the equivalent Teams gateways
	$TeamsGatewayList = @()
	$SfBGatewayList = $VoiceRoute.PstnGatewayList
	ForEach ($SfBGateway in $SfBGatewayList) {
		$TeamsGateway = $PSTNGWMatch[[regex]::Match($SfBGateway, 'PstnGateway:(.*)').Groups[1].Value] # Strip the PstnGateway from the gateway name
		If ($TeamsGateway) {
			$TeamsGatewayList += $TeamsGateway
		}
	}

	$VRDetails = @{
		Identity = $VoiceRoute.Identity
		NumberPattern = $VoiceRoute.NumberPattern
		Priority = $VoiceRoute.Priority
		OnlinePstnUsages = $VoiceRoute.PstnUsages
		Description = $VoiceRoute.Description
	}
	
	# Only include gateway list if one is defined. MS throws an error if you pass a null/empty OnlinePstnGatewayList
	If ($TeamsGatewayList) {
		$VRDetails.Add('OnlinePstnGatewayList', $TeamsGatewayList)
	}
	
	If ($VRExists) {
		$null = (Set-CsOnlineVoiceRoute @VRDetails)
	}
	Else {
		$null = (New-CsOnlineVoiceRoute @VRDetails)
	}
}

# Copy voice policies to Teams voice routing policies
Write-Host -Object 'Copying SfB voice policies to Teams voice routing policies'

ForEach ($VoicePolicy in $VoicePolicies) {
	Write-Verbose "Copying $($VoicePolicy.Identity) voice policy"
	$VPExists = (Get-CsOnlineVoiceRoutingPolicy -Identity $VoicePolicy.Identity -ErrorAction SilentlyContinue).Identity

	$VPDetails = @{
		Identity = $VoicePolicy.Identity -replace "^Site\:", "" # Since site-level voice policies aren't supported in Teams, site-level SfB voice policies will be converted to user-level Teams voice routing policies
		OnlinePstnUsages = $VoicePolicy.PstnUsages
		Description = $VoicePolicy.Description
	}
	
	If ($VPExists) {
		$null = (Set-CsOnlineVoiceRoutingPolicy @VPDetails)
	}
	Else {
		$null = (New-CsOnlineVoiceRoutingPolicy @VPDetails)
	}
}


# Create variable to hold variable names: Entry1=variable name, Entry2=text, Entry3=parameter name
$TransRuleVars = @(('OutboundCallingTransRules', 'calling', 'OutbundTeamsNumberTranslationRules'), ('OutboundCalledTransRules', 'called', 'OutboundPSTNNumberTranslationRules'))

ForEach ($TransRuleVar in $TransRuleVars) {
	# Create Teams translation rules from SfB equivalents and assign to gateways
	Write-Host -Object "Copying SfB outbound $($TransRuleVar[1]) translation rules to Teams translation rules"

	ForEach ($TeamsTransRule in (Get-Variable -Name $TransRuleVar[0] -ValueOnly)) {
		Write-Verbose "Copying $($TeamsTransRule.Identity) outbound $($TransRuleVar[1]) translation rule"
		$OCTRExists = (Get-CsTeamsTranslationRule -Identity $TeamsTransRule.Name -ErrorAction SilentlyContinue).Identity

		$OCTRDetails = @{
			Identity = $TeamsTransRule.Name
			Pattern = $TeamsTransRule.Pattern
			Translation = $TeamsTransRule.Translation
			Description = $TeamsTransRule.Description
		}
		
		$SfBGateway = [regex]::Match($TeamsTransRule.Identity, 'PstnGateway:(.*)/').Groups[1].Value
		$TeamsGateway = $PSTNGWMatch[$SfbGateway]
		
		If ($OCTRExists) {
			If (Get-CsTeamsTranslationRule -Identity $TeamsTransRule.Name | Where {$_.Pattern -eq $TeamsTransRule.Pattern -and $_.Translation -eq $TeamsTransRule.Translation}) {
				Write-Verbose "Matching existing rule: $($TeamsTransRule.Name)"
			}
			Else { # Rulename is the same, but the translation or pattern is different, so need a new rule
				Write-Verbose "$($TeamsTransRule.Name) rule has same name, but different details. Creating new rule: $($TeamsTransRule.Name)_$TeamsGateway"
				$OCTRDetails.Identity = $TeamsTransRule.Name + '_' + $TeamsGateway
				$null = (New-CsTeamsTranslationRule @OCTRDetails)
			}
		}
		Else {
			Write-Verbose "Creating translation rule called $($TeamsTransRule.Name)"
			$null = (New-CsTeamsTranslationRule @OCTRDetails)
		}
		
		$Params = @{$TransRuleVar[2] = @{Add=$OCTRDetails.Identity}}
		Write-Verbose "Assigning $($TeamsTransRule.Name) translation rule to gateway"
		Set-CsOnlinePSTNGateway -Identity $TeamsGateway @Params 
	}
}
Write-Host -Object 'Finished!'
