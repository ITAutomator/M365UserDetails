# https://github.com/ITAutomator/M365UserDetails

#####
## To enable scrips, Run powershell 'as admin' then type
## Set-ExecutionPolicy Unrestricted
#####


#################### Transcript Open
$Transcript = [System.IO.Path]::GetTempFileName()               
Start-Transcript -path $Transcript | Out-Null
#################### Transcript Open

### Main function header - Put ITAutomator.psm1 in same folder as script
$scriptFullname = $PSCommandPath ; if (!($scriptFullname)) {$scriptFullname =$MyInvocation.InvocationName }
$scriptXML      = $scriptFullname.Substring(0, $scriptFullname.LastIndexOf('.'))+ ".xml"  ### replace .ps1 with .xml
$scriptCSV      = $scriptFullname.Substring(0, $scriptFullname.LastIndexOf('.'))+ ".csv"  ### replace .ps1 with .csv
$scriptDir      = Split-Path -Path $scriptFullname -Parent
$scriptName     = Split-Path -Path $scriptFullname -Leaf
$scriptBase     = $scriptName.Substring(0, $scriptName.LastIndexOf('.'))
if ((Test-Path("$scriptDir\ITAutomator.psm1"))) {Import-Module "$scriptDir\ITAutomator.psm1" -Force} else {write-output "Err 99: Couldn't find ITAutomator.psm1";Start-Sleep -Seconds 10;Exit(99)}
# Get-Command -module RethinkitFunction  ##Shows a list of available functions
############

####
Write-Host "-----------------------------------------------------------------------------"
Write-Host ("$scriptName        Computer:$env:computername User:$env:username PSver:"+($PSVersionTable.PSVersion.Major))
Write-Host ""
Write-Host "Bulk actions in O365"
Write-Host ""
Write-Host "CSV: $(Split-Path $scriptCSV -leaf)"
Write-Host "-----------------------------------------------------------------------------"
$no_errors = $true
$error_txt = ""
$results = @()
LoadModule "AzureAD"
$PSCred = $false
Try
{
    $conn=Connect-AzureAD
    $PSCred = $true
}
Catch{}
if (!($PSCred)) 
{ #creds not entered
    Write-Host "[No credentials entered]"
} #creds not entered
else
{ #creds entered
    Write-Host "--------------------"
    Write-Host "CONNECTED: $($conn.TenantDomain) [via $($conn.Account.Id)]"
    Write-Host "--------------------"
    $props=@()
    $props+="ObjectId"
    $props+="ObjectType"
    $props+="UserType"
    $props+="AccountEnabled"
    $props+="UserPrincipalName"
    $props+="Mail"
    $props+="DisplayName"
    $props+="GivenName"
    $props+="Surname"
    $props+="JobTitle"
    $props+="Department"
    $props+="CompanyName"
    $props+="TelephoneNumber"
    $props+="Mobile"
    #$props+="FacsimileTelephoneNumber"
    $props+="StreetAddress"
    $props+="City"
    $props+="PostalCode"
    $props+="State"
    $props+="Country"
    $props+="PhysicalDeliveryOfficeName"
    $props+="PreferredLanguage"
    #$props+="UsageLocation"
    #$props+="ShowInAddressList"
    #$props+="ConsentProvidedForMinor"
    #$props+="AgeGroup"
    #$props+="DirSyncEnabled"
    #$props+="CreationType"
    #$props+="MailNickName"
    #$props+="SipProxyAddress"
    ####
    Write-host "Exporting n properties: $($props.Count)"
	####### Retrieve Azure AD User list
    $adusers = Get-AzureADUser -All $true ##-Filter "userPrincipalName eq '$($x.Mail)'"
    Write-Host "AzureADUser Count: $($adusers.count) [All users]"
    $adusers = $adusers | Where-Object UserType -EQ Member
    Write-Host "AzureADUser Count: $($adusers.count) [UserType=Members (vs Guests)]"
    $adusers = $adusers | Where-Object AccountEnabled -eq $true
    Write-Host "AzureADUser Count: $($adusers.count) [AccountEnabled=True]"
    $adusers = $adusers | where-object { $_.AssignedLicenses.count -ne 0}
    #### clear rows
    $rows = @()
    ForEach ($aduser in $adusers)
    {
        # create a new empty row
        $row = New-Object -TypeName psobject
        # append each column needed
        ForEach ($prop in $props)
        {
            $row | Add-Member -Type NoteProperty -Name $prop -Value $aduser.($prop)
        }
        # append manager columns
        $mgr = Get-AzureADUserManager -objectId $aduser.ObjectId
        $row | Add-Member -Type NoteProperty -Name "MgrName"  -Value $mgr.DisplayName
        $row | Add-Member -Type NoteProperty -Name "MgrEmail" -Value $mgr.Mail
        ### append row
        $rows+= $row
    }
    ###
    Write-Host "AzureADUser Count: $($adusers.count) [AssignedLicenses=True]"
    Write-host "Exporting info to CSV..."
    $date = get-date -format "yyyy-MM-dd_HH-mm-ss"
    $scriptCSVdated= $scriptCSV.Replace(".csv"," $($date).csv")
    $rows | Export-Csv $scriptCSVdated -NoTypeInformation
	#######
    Get-PSSession 
    Get-PSSession | Remove-PSSession
    Write-host "File: $(split-path $scriptCSVdated -Leaf)" -ForegroundColor Yellow
    Write-Host "------------------------------------------------------------------------------------"
    $message ="Done. Press [Enter] to exit."
    Write-Host $message
    Write-Host "------------------------------------------------------------------------------------"
	#################### Transcript Save
    Stop-Transcript | Out-Null
    $date = get-date -format "yyyy-MM-dd_HH-mm-ss"
    New-Item -Path (Join-Path (Split-Path $scriptFullname -Parent) ("\Logs")) -ItemType Directory -Force | Out-Null #Make Logs folder
    $TranscriptTarget = Join-Path (Split-Path $scriptFullname -Parent) ("Logs\"+[System.IO.Path]::GetFileNameWithoutExtension($scriptFullname)+"_"+$date+"_log.txt")
    If (Test-Path $TranscriptTarget) {Remove-Item $TranscriptTarget -Force}
    Move-Item $Transcript $TranscriptTarget -Force
    #################### Transcript Save
} #creds entered
PauseTimed -quiet 3 #$message
Pause
