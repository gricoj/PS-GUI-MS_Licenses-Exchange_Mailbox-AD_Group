# PowerShell GUI for Licenses and Intune

In this enviorment we had Hybrid Exchange enviorment, users mailboxes where either on *Exchange Online* or *Exchange On-Premise*. We were actively migrating users from *Exchange On-Premise* to *Exchange Online* (Office 365)

We also had a hybrid MDM enviorment, migrating users from Citrix XenMobile to Microsoft Intune. The Hybrid Exchange enviorment meant that we had to have *Device Configuration Profiles* and *App Configuration Profiles* that would support both *Exchange Online* and *Exchange On-Premise* users. Additionally, users require a [license](https://docs.microsoft.com/en-us/intune/fundamentals/licenses) in order to access Intune.

We approached this problem by creating a security group: *MDM_OnPremExchange*, we would apply *Exchange On-Premise* configuration profiles only to users who are part of the *MDM_OnPremExchange* group. Similarly, we would apply *Exchange Online* configuration profiles to all users excluding those who are part of the *MDM_OnPremExchange* group.

The GUI's purposes:
- To show whether the user's maibox is on *Exchange Online*  or *Exchange On-Premise*
- To show the user's licenses
- To add/remove a user from the *MDM_OnPremExchange* group

### Table of Contents
[Requirements](https://github.com/gricoj/PS-License-Intune-GUI#requirements)

[Functions](https://github.com/gricoj/PS-License-Intune-GUI#functions)

[Contraints](https://github.com/gricoj/PS-License-Intune-GUI#contraints)


## Requirements
The GUI will require two PowerShell Modules: *ActiveDirectory* and *AzureAD*
The *AzureAD* module can be installed by the command:
```powershell
Install-Module AzureAD
```
The *ActiveDirectory* module can be installed by enabling the *Active Directory lightweight Directory Services* in *Windows   Features* (*Control Panel* / *Programs and Features* / *Turn  Windows Features on or off*)

## Functions

#### Get-ExchangeStatus
Purpose: To get whether a user's mailboxes is on *Exchange Online* or *Exchange On-Premise*

Arguments: The function takes the user's username

Process: We use the *Get-ADUser* cmdlet to get the user's *msExchRecipmientTypeDetails*, depending on that value we determine whether the user's mailbox is on *Exchange Online* or *Exchange On-Premise*

Output: Returns *O365* if the user's mailbox is on *Exchange Online*. Returns *On Premise* if the user's mailbox is on *Exchange On-Premise*. Returns *error* if there the username is invalid.
```powershell
function Get-ExchangeStatus {
    param([string]$Username)
    $Site = ((Get-ADUser $Username -Properties *).msExchRecipientTypeDetails)
    if($Site -eq 2147483648){ return "O365" }
    elseif($Site -eq 1){ return "On Premise" }
    else { return "Error" }
}
```
#### Get-UserLicenseDetail
Purpose: To get the licenses applied to a user

Arguments: The function takes the user's principal name (user's email address)

Process: We use the *Get-AzureADUser* cmdlet to get the licence SKUs a user has. We then compare license SKUs to the licenses the organization has using the *Get-AzureADSubscribedSku* cmdlet to get the name of the license. We store the names of the licenses the user has in an array.
Output: We return the array of license names
```powershell
function Get-UserLicenseDetail {
    param([string]$UserPrincipalName)
    $SkuIDs = (Get-AzureADUser -ObjectId $UserPrincipalName | Select -ExpandProperty AssignedLicenses).SkuId
    $LicenseName = @()
    foreach($SkuID in $SkuIDs){
        $LicenseName += (Get-AzureADSubscribedSku | Where {$_.SkuId -eq $SkuID}).SkuPartNumber
    }
    Return $LicenseName
}
```
#### In-OnPremGroup
Purpose: To get whether the user is in the *MDM_OnPremExchange* security group

Arguments: The function takes the user's username

Process: We use the *Get-ADGroupMember* cmdlet to get a list of all the users in the *MDM_OnPremExchange* group. We then check if the user is in that group

Output: Returns *$true* if the user is in the MDM_OnPremExchange* security group. Returns *$false* if the user is not in the *MDM_OnPremExchange* group
```powershell
function In-OnPremGroup {
    param([string]$Username)
    $Users = (Get-ADGroupMember -Identity MDM_OnPremExchange).SamAccountName
    $InGroup = $false
    if($Users -contains $Username){ $InGroup = $true }
    return $InGroup
}
```
## Constraints
- A user should only be removed from the *MDM_OnPremExchange* group if
    - The user's mailbox is on *Exchange Online* and the user is in the *MDM_OnPremExchange* group
- A user should only be added to the *MDM_OnPremExchange* group if
    - The user's mailbox is on *Exchange On Premise*, the user is not in the *MDM_OnPremExchange* group and the user has appropriate licenses
