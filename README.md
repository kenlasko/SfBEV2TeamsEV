# SfBEV2TeamsEV
PowerShell script that copies an on-prem Skype for Business Enterprise Voice configuration to a Microsoft Teams Enterprise Voice Direct Routing deployment

## Getting Started

Download the script onto a Windows machine that has the Skype for Business on-prem PowerShell module installed along with the Office 365 PowerShell module.

### Prerequisites

Requires that you have the Skype for Business on-prem and Office 365 PowerShell modules installed, and you have at least read-access to your Skype for Business Enterprise Voice configuration. You may have to set your execution policy to unrestricted to run this script: 

Set-ExecutionPolicy Unrestricted


## Running the Script

Simply run **.\Copy-SfBEV2TeamsEV.ps1** from a PowerShell prompt on a Skype for Business server (or a computer with the SfB Management tools installed). If you are not already connected to your Teams tenant, the script will prompt for authentication. If your admin account is not a @tenantname.onmicrosoft.com account, then you should use the **-OverrideAdminDomain** switch. The script will copy all SfB dialplans with normalization rules, routes, voice policies, PSTN gateways and outbound translation rules to their Teams equivalents. It will attempt to match SfB gateways to Teams gateways, and will prompt for how to match if one can't be found automatically.

By default, the script will clean out any existing Teams EV config, including dialplans, voice routes, voice routing policies, PSTN usages and translation rules. You can override this behaviour by using the **-KeepExisting** switch.

## More Information

Check my blog post for more details about this script: https://ucken.blogspot.com/2020/02/copy-sfb-ev-to-teams-ev.html

## Authors

**Ken Lasko** 
* https://github.com/kenlasko
* https://ucdialplans.com
* https://ucken.blogspot.com
* https://twitter.com/kenlasko
* https://www.linkedin.com/in/kenlasko71/
