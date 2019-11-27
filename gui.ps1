Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '400,250'
$Form.text                       = "Form"
$Form.BackColor                  = "#4a4a4a"
$Form.TopMost                    = $false

$Search_Box                        = New-Object system.Windows.Forms.TextBox
$Search_Box.multiline              = $false
$Search_Box.text                   = "Username"
$Search_Box.BackColor              = "#777773"
$Search_Box.width                  = 132
$Search_Box.height                 = 20
$Search_Box.location               = New-Object System.Drawing.Point(14,30)
$Search_Box.Font                   = 'Microsoft Sans Serif,10'
$Search_Box.ForeColor                = "#ffffff"

$Search_Button                         = New-Object system.Windows.Forms.Button
$Search_Button.text                    = "Search"
$Search_Button.width                   = 60
$Search_Button.height                  = 30
$Search_Button.location                = New-Object System.Drawing.Point(160,25)
$Search_Button.Font                    = 'Microsoft Sans Serif,10'
$Search_Button.ForeColor                = "#ffffff"

$ExchangeStatus_Box                        = New-Object system.Windows.Forms.TextBox
$ExchangeStatus_Box.multiline              = $false
$ExchangeStatus_Box.BackColor              = "#777773"
$ExchangeStatus_Box.width                  = 210
$ExchangeStatus_Box.height                 = 20
$ExchangeStatus_Box.location               = New-Object System.Drawing.Point(154,76)
$ExchangeStatus_Box.Font                   = 'Microsoft Sans Serif,10'
$ExchangeStatus_Box.ForeColor                = "#ffffff"

$ExchangeStatus_Label                          = New-Object system.Windows.Forms.Label
$ExchangeStatus_Label.text                     = "Exchange Status"
$ExchangeStatus_Label.AutoSize                 = $true
$ExchangeStatus_Label.width                    = 25
$ExchangeStatus_Label.height                   = 10
$ExchangeStatus_Label.location                 = New-Object System.Drawing.Point(30,76)
$ExchangeStatus_Label.Font                     = 'Microsoft Sans Serif,10'
$ExchangeStatus_Label.ForeColor                = "#ffffff"

$LicenseDetail_Box                        = New-Object system.Windows.Forms.TextBox
$LicenseDetail_Box.multiline              = $true
$LicenseDetail_Box.BackColor              = "#777773"
$LicenseDetail_Box.width                  = 210
$LicenseDetail_Box.height                 = 74
$LicenseDetail_Box.location               = New-Object System.Drawing.Point(154,105)
$LicenseDetail_Box.Font                   = 'Microsoft Sans Serif,10'
$LicenseDetail_Box.ForeColor                = "#ffffff"

$LicenseDetail_Label                          = New-Object system.Windows.Forms.Label
$LicenseDetail_Label.text                     = "License Detail"
$LicenseDetail_Label.AutoSize                 = $true
$LicenseDetail_Label.width                    = 25
$LicenseDetail_Label.height                   = 10
$LicenseDetail_Label.location                 = New-Object System.Drawing.Point(37,130)
$LicenseDetail_Label.Font                     = 'Microsoft Sans Serif,10'
$LicenseDetail_Label.ForeColor                = "#ffffff"

$InADGroup_Box                        = New-Object system.Windows.Forms.TextBox
$InADGroup_Box.multiline              = $true
$InADGroup_Box.BackColor              = "#777773"
$InADGroup_Box.width                  = 70
$InADGroup_Box.height                 = 20
$InADGroup_Box.location               = New-Object System.Drawing.Point(194,185)
$InADGroup_Box.Font                   = 'Microsoft Sans Serif,10'
$InADGroup_Box.ForeColor                = "#ffffff"

$InADGroup_Label                          = New-Object system.Windows.Forms.Label
$InADGroup_Label.text                     = "In MDM_OnPremExchange"
$InADGroup_Label.AutoSize                 = $true
$InADGroup_Label.width                    = 25
$InADGroup_Label.height                   = 10
$InADGroup_Label.location                 = New-Object System.Drawing.Point(20,185)
$InADGroup_Label.Font                     = 'Microsoft Sans Serif,10'
$InADGroup_Label.ForeColor                = "#ffffff"

$AddRemove_Button                         = New-Object system.Windows.Forms.Button
$AddRemove_Button.width                   = 60
$AddRemove_Button.height                  = 20
$AddRemove_Button.location                = New-Object System.Drawing.Point(270,185)
$AddRemove_Button.Font                    = 'Microsoft Sans Serif,8'
$AddRemove_Button.ForeColor                = "#ffffff"
$AddRemove_Button.visible                 = $false

$Form.controls.AddRange(@($Search_Box,$Search_Button,$AddRemove_Button,$ExchangeStatus_Box,$ExchangeStatus_Label,$LicenseDetail_Box,$LicenseDetail_Label,$InADGroup_Box,$InADGroup_Label))

function Get-ExchangeStatus {
    param([string]$Username)
    $Site = ((Get-ADUser $Username -Properties *).msExchRecipientTypeDetails)
    if($Site -eq 2147483648){ return "O365" }
    elseif($Site -eq 1){ return "On Premise" }
    else { return "Error" }
}

function Get-UserLicenseDetail {
    param([string]$UserPrincipalName)
    $SkuIDs = (Get-AzureADUser -ObjectId $UserPrincipalName | Select -ExpandProperty AssignedLicenses).SkuId
    $LicenseName = @()
    foreach($SkuID in $SkuIDs){
        $LicenseName += (Get-AzureADSubscribedSku | Where {$_.SkuId -eq $SkuID}).SkuPartNumber
    }
    Return $LicenseName
}

function In-OnPremGroup {
    param([string]$Username)
    $Users = (Get-ADGroupMember -Identity MDM_OnPremExchange).SamAccountName
    $InGroup = $false
    if($Users -contains $Username){ $InGroup = $true }
    return $InGroup
}

$Search_Box.Add_Click({ $Search_Box.Clear()})

$Search_Button.Add_Click({
    $AddRemove_Button.visible = $false
    $ExchangeStatus_Box.Clear()
    $LicenseDetail_Box.Clear()
    $InADGroup_Box.Clear()
    if($Search_Box.Text -eq "" -or $Search_Box.Text -eq "Username"){ return }

    $UserName = $Search_Box.Text
    $ExchangeStatus_Box.Text = Get-ExchangeStatus -Username $UserName
    
    $UPN = (Get-ADUser $UserName).UserPrincipalName
    $Licenses = (Get-UserLicenseDetail -UserPrincipalName $UPN)

    foreach($License in $Licenses){
        $LicenseDetail_Box.AppendText($License)
        $LicenseDetail_Box.AppendText("`n")
    }

    $InADGroup_Box.Text = In-OnPremGroup -Username $UserName

    if($InADGroup_Box.Text -eq $false -and $ExchangeStatus_Box.Text -eq "On Premise" -and $LicenseDetail_Box.lines -contains "EMS" ){$AddRemove_Button.text = "Add"; $AddRemove_Button.visible = $true; }
    elseif($InADGroup_Box.Text -eq $true -and $ExchangeStatus_Box.Text -eq "O365"){ $AddRemove_Button.text = "Remove"; $AddRemove_Button.visible = $true;}
        
})

$AddRemove_Button.Add_Click({
    if($AddRemove_Button.text -eq "Remove"){Remove-ADGroupMember -Identity MDM_OnPremExchange -Members $Search_Box.Text}
    elseif($AddRemove_Button.text -eq "Add"){Add-ADGroupMember -Identity MDM_OnPremExchange -Members $Search_Box.Text}
    Start-Sleep -s 3
    $InADGroup_Box.Text = In-OnPremGroup -UserPrincipalName $Search_Box.Text
    $AddRemove_Button.visible = $false
})   

$Form.ShowDialog()
