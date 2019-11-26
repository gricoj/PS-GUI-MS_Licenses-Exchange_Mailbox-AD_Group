# PowerShell GUI for Licenses and Intune
In this enviorment there was a hybrid Exchange enviorment, users mailboxes where either in *Exchange Online* or *Exchange On-Premise*. This meant that we had to have *Device Configuration Profiles* and *App Configuration Profiles* that would support both types of users. 

We approached this problem by creating a security group *MDM_OnPremExchange*, we would apply *Exchange On-Premise* configuration profiles only to users who are part of the *MDM_OnPremExchange* group. Similarly, we would apply *Exchange Online* configuration profiles to all users excluding those who are part of the *MDM_OnPremExchange* group.

The GUI's purposes:
- To show whether the user's maibox is on *Exchange Online*  or *Exchange On-Premise*
- To show the user's licenses
- To add/remove a user from the *MDM_OnPremExchange* group

