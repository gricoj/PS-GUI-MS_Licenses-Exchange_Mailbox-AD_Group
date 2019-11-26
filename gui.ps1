Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '400,250'
$Form.text                       = "Form"
$Form.BackColor                  = "#4a4a4a"
$Form.TopMost                    = $false

$TextBox1                        = New-Object system.Windows.Forms.TextBox
$TextBox1.multiline              = $false
$TextBox1.text                   = "Username"
$TextBox1.BackColor              = "#777773"
$TextBox1.width                  = 132
$TextBox1.height                 = 20
$TextBox1.location               = New-Object System.Drawing.Point(14,30)
$TextBox1.Font                   = 'Microsoft Sans Serif,10'
$TextBox1.ForeColor                = "#ffffff"

$Button1                         = New-Object system.Windows.Forms.Button
$Button1.text                    = "Search"
$Button1.width                   = 60
$Button1.height                  = 30
$Button1.location                = New-Object System.Drawing.Point(160,25)
$Button1.Font                    = 'Microsoft Sans Serif,10'
$Button1.ForeColor                = "#ffffff"

$TextBox2                        = New-Object system.Windows.Forms.TextBox
$TextBox2.multiline              = $false
$TextBox2.BackColor              = "#777773"
$TextBox2.width                  = 210
$TextBox2.height                 = 20
$TextBox2.location               = New-Object System.Drawing.Point(154,76)
$TextBox2.Font                   = 'Microsoft Sans Serif,10'
$TextBox2.ForeColor                = "#ffffff"

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "Exchange Status"
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(30,76)
$Label1.Font                     = 'Microsoft Sans Serif,10'
$Label1.ForeColor                = "#ffffff"

$TextBox3                        = New-Object system.Windows.Forms.TextBox
$TextBox3.multiline              = $true
$TextBox3.BackColor              = "#777773"
$TextBox3.width                  = 210
$TextBox3.height                 = 74
$TextBox3.location               = New-Object System.Drawing.Point(154,105)
$TextBox3.Font                   = 'Microsoft Sans Serif,10'
$TextBox3.ForeColor                = "#ffffff"

$Label2                          = New-Object system.Windows.Forms.Label
$Label2.text                     = "License Detail"
$Label2.AutoSize                 = $true
$Label2.width                    = 25
$Label2.height                   = 10
$Label2.location                 = New-Object System.Drawing.Point(37,130)
$Label2.Font                     = 'Microsoft Sans Serif,10'
$Label2.ForeColor                = "#ffffff"

$TextBox4                        = New-Object system.Windows.Forms.TextBox
$TextBox4.multiline              = $true
$TextBox4.BackColor              = "#777773"
$TextBox4.width                  = 70
$TextBox4.height                 = 20
$TextBox4.location               = New-Object System.Drawing.Point(194,185)
$TextBox4.Font                   = 'Microsoft Sans Serif,10'
$TextBox4.ForeColor                = "#ffffff"

$Label3                          = New-Object system.Windows.Forms.Label
$Label3.text                     = "In MDM_OnPremExchange"
$Label3.AutoSize                 = $true
$Label3.width                    = 25
$Label3.height                   = 10
$Label3.location                 = New-Object System.Drawing.Point(20,185)
$Label3.Font                     = 'Microsoft Sans Serif,10'
$Label3.ForeColor                = "#ffffff"

$Button2                         = New-Object system.Windows.Forms.Button
$Button2.width                   = 60
$Button2.height                  = 20
$Button2.location                = New-Object System.Drawing.Point(270,185)
$Button2.Font                    = 'Microsoft Sans Serif,8'
$Button2.ForeColor                = "#ffffff"
$Button2.visible                 = $false

$Form.controls.AddRange(@($TextBox1,$Button1,$Button2,$TextBox2,$Label1,$TextBox3,$Label2,$TextBox4,$Label3))


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
        [array]$LicenseName += (Get-AzureADSubscribedSku | Where {$_.SkuId -eq $SkuID}).SkuPartNumber
    }
    
    Return $LicenseName
}

function In-OnPremGroup {
    param([string]$UserPrincipalName)
    $Users = (Get-ADGroupMember -Identity MDM_OnPremExchange).SamAccountName
    $InGroup = $false
    if($Users -contains $UserPrincipalName){ $InGroup = $true }
    return $InGroup
}

    

$TextBox1.Add_Click({ $TextBox1.Clear()})

$Button1.Add_Click({
    $Button2.visible = $false
    $TextBox2.Clear()
    $TextBox3.Clear()
    $TextBox4.Clear()
    if($TextBox1.Text -eq "" -or $TextBox1.Text -eq "Username"){ return }

    $UserName = $TextBox1.Text
    $TextBox2.Text = Get-ExchangeStatus -Username $UserName
    
    $UPN = (Get-ADUser $UserName).UserPrincipalName
    $Licenses = (Get-UserLicenseDetail -UserPrincipalName $UPN)

    foreach($License in $Licenses){
        $TextBox3.AppendText($License)
        $TextBox3.AppendText("`n")
    }

    $TextBox4.Text = In-OnPremGroup -UserPrincipalName $UserName

    if($TextBox4.Text -eq $false -and $TextBox2.Text -eq "On Premise" -and $TextBox3.lines -contains "EMS" ){$Button2.text = "Add"; $Button2.visible = $true; }
    elseif($TextBox4.Text -eq $true -and $TextBox2.Text -eq "O365"){ $Button2.text = "Remove"; $Button2.visible = $true;}
        
})

$Button2.Add_Click({
    if($Button2.text -eq "Remove"){Remove-ADGroupMember -Identity MDM_OnPremExchange -Members $TextBox1.Text}
    elseif($Button2.text -eq "Add"){Add-ADGroupMember -Identity MDM_OnPremExchange -Members $TextBox1.Text}
    Start-Sleep -s 3
    $TextBox4.Text = In-OnPremGroup -UserPrincipalName $TextBox1.Text
    $Button2.visible = $false
})
    

$Form.ShowDialog()