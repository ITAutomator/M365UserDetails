#####
## To enable scrips, Run powershell 'as admin' then type
## Set-ExecutionPolicy Unrestricted
#####

############################
##
## O365MailboxFullAccess.ps1
##
## Applies permission changes to a folder
## Note: For delegates - see delegates.ps1
##       For full permissions - see fullAccess.ps1
##
## CSV columns
## IdentityFrom,IdentityTo,AddRemove,AutoMapping
## Truvvo Assistant Team,admin@roundtableip.com,Add,FALSE
## Truvvo Client Service,admin@roundtableip.com,Add,FALSE
## 
## IdentityFrom: Mailbox giving access.  Email or Display name
## IdentityTo:   Mailbox getting access. Email or Display name
## 
## AddRemove 
## Add   :  Add full access
## Remove:  Remove full access
##
## AutoMapping
## TRUE   : Automatically adds the inbox to left side of IdentityTo's Outlook (30 mins)  This is portal default.
## FALSE  : Doesn't automap the inbox
##
############################


#################### Transcript Open
$Transcript = [System.IO.Path]::GetTempFileName()               
Start-Transcript -path $Transcript | Out-Null
#################### Transcript Open

### Main function header - Put RethinkitFunctions.psm1 in same folder as script
$scriptFullname = $PSCommandPath ; if (!($scriptFullname)) {$scriptFullname =$MyInvocation.InvocationName }
$scriptXML      = $scriptFullname.Substring(0, $scriptFullname.LastIndexOf('.'))+ ".xml"  ### replace .ps1 with .xml
$scriptCSV      = $scriptFullname.Substring(0, $scriptFullname.LastIndexOf('.'))+ ".csv"  ### replace .ps1 with .csv
$scriptDir      = Split-Path -Path $scriptFullname -Parent
$scriptName     = Split-Path -Path $scriptFullname -Leaf
$scriptBase     = $scriptName.Substring(0, $scriptName.LastIndexOf('.'))
if ((Test-Path("$scriptDir\RethinkitFunctions.psm1"))) {Import-Module "$scriptDir\RethinkitFunctions.psm1" -Force} else {write-output "Err 99: Couldn't find RethinkitFunctions.psm1";Start-Sleep -Seconds 10;Exit(99)}
# Get-Command -module RethinkitFunction  ##Shows a list of available functions
############
$O365_PasswordXML   = $scriptDir+ "\O365_Password.xml"
############
if (!(Test-Path $scriptCSV))
{
    ######### Template
    "IdentityFrom,IdentityTo,AddRemove,AutoMapping" | Add-Content $scriptCSV
    "Truvvo Assistant Team,admin@roundtableip.com,Add,FALSE" | Add-Content $scriptCSV
	"Truvvo Client Service,admin@roundtableip.com,Add,FALSE" | Add-Content $scriptCSV
    ######### Template
	$ErrOut=201; Write-Host "Err $ErrOut : Couldn't find '$(Split-Path $scriptCSV -leaf)'. Template CSV created. Edit CSV and run again.";Pause; Exit($ErrOut)
}
## ----------Fill $entries with contents of file or something
$entries=@(import-csv $scriptCSV)
$entries_cols = ($entries | Get-Member | Where-Object -Property "MemberType" -EQ "NoteProperty" | Select-Object "Name").Name
$entriescount = $entries.count
##
$props = "Mail","DisplayName","TelephoneNumber","Mobile","JobTitle","CompanyName","StreetAddress","City","State","PostalCode","Country"
####
Write-Host "-----------------------------------------------------------------------------"
Write-Host ("$scriptName        Computer:$env:computername User:$env:username PSver:"+($PSVersionTable.PSVersion.Major))
Write-Host ""
Write-Host "Bulk actions in O365"
Write-Host ""
#Write-Host "admin_username: $($Globals.admin_username)"
Write-Host "XML: $(Split-Path $O365_PasswordXML -leaf)"
Write-Host "CSV: $(Split-Path $scriptCSV -leaf) ($($entriescount) entries)"
Write-Host "Possible column names are:"
Write-Host "Mail (Required),$($props -join ",")"  -ForegroundColor Green
Write-Host "Columns found:"
Write-Host ($entries_cols -join ", ") -ForegroundColor Green
Write-Host 'Use ""        to leave column as is (no change)'
Write-Host 'Use "<clear>" to clear column of contents'
Write-Host ""
$entries | Format-Table
Write-Host "-----------------------------------------------------------------------------"

$no_errors = $true
$error_txt = ""
$results = @()

###
$module= "AzureAD"                  ; Write-Host "Loadmodule $($module)..." -NoNewline ; $lm_result=LoadModule $module ; Write-Host $lm_result
try {Get-AzureADDomain | Where-Object -Property IsDefault -EQ $true}
Catch{Connect-AzureAD}
###

    
    Write-Host "--------------------"
    Write-Host "CONNECTED: $($PSCred.UserName)"
    Write-Host "--------------------"
    $processed=0
    $message="$entriescount Entries. Continue?"
    $choices = [System.Management.Automation.Host.ChoiceDescription[]] @("&Yes","&No")
    [int]$defaultChoice = 0
    $choiceRTN = $host.ui.PromptForChoice($caption,$message, $choices,$defaultChoice)
    if ($choiceRTN -eq 1)
    { "Aborting" }
    else 
    { ## continue choices

    ### Connect to O365
    $choiceLoop=0
    $i=0
    $change_i=0    
    foreach ($x in $entries)
    {
        $i++
        write-host "-----" $i of $entriescount $x
        if ($choiceLoop -ne 1)
            {
            $message="Process entry "+$i+"?"
            $choices = [System.Management.Automation.Host.ChoiceDescription[]] @("&Yes","Yes to &All","&No","No and E&xit")
            [int]$defaultChoice = 1
            $choiceLoop = $host.ui.PromptForChoice($caption,$message, $choices,$defaultChoice)
            }
        if (($choiceLoop -eq 0) -or ($choiceLoop -eq 1))
            {
            $processed++
		    ####### Start code for object $x
						
			#######
			#IdentityFrom	IdentityTo	AddRemove	AutoMappingToIdentityTo
            #Get-MailboxPermission -Identity john@contoso.com -User "Ayla"
            #Add-MailboxPermission -Identity "Jeroen Cool" -User "Mark Steele" -AccessRights FullAccess -InheritanceType All -AutoMapping $false#
            #Remove-MailboxPermission -Identity Test1 -User Test2 -AccessRights FullAccess -InheritanceType All
            #######
            $aduser = Get-AzureADUser -Filter "userPrincipalName eq '$($x.Mail)'"
            if ($aduser)
            { # found aduser            
			    #######
                
			    ####### Display 'before' info
                Write-host "[Before]"
                ($aduser | Select-Object $entries_cols | Format-List | Out-String) -Split "`r`n" | Where({ $_ -ne "" }) | Write-Host
			    #######
                $change_made = $false
                ForEach ($prop in $entries_cols)
                {
                    If ($x.$prop -eq "")
                    { #No data
                    }
                    ElseIf ($x.$prop -eq "<clear>")
                    {
                        if ($aduser.$prop -eq "")
                        {
                            Write-Host "$($prop): $($aduser.$prop) <clear> [Already OK]"
                        }
                        else
                        {
                            Write-Host "$($prop): $($aduser.$prop) <clear>" -ForegroundColor Yellow
                            Set-AzureADUser -ObjectId $aduser.ObjectId -$($prop) ""
                            $change_made = $true
                        }
                    }
                    ElseIf ($x.$prop -eq $aduser.$prop)
                    {  #No update
                    } 
                    Else
                    {
                        Write-Host "$($prop): [$($aduser.$prop)] will be changed to [$($x.$prop)]" -ForegroundColor Yellow
                        $myargs = @{
                          ObjectId = $aduser.ObjectId
                          $prop = $x.$prop
                        }
                        Set-AzureADUser @myargs
                        $change_made = $true
                    }
                }
                if ($change_made)
                {
                    $change_i+=1
                    ####### Display 'after' info
                    Write-host "[After]"
                    $aduser = Get-AzureADUser -Filter "userPrincipalName eq '$($x.Mail)'"
                    ($aduser | Select-Object $entries_cols | Format-List | Out-String) -Split "`r`n" | Where({ $_ -ne "" }) | Write-Host
			        #######
                }
                else
                {
                    Write-host "[OK] Nothing changed"
                }
            } # found aduser
            else
            {
                Write-Warning "No such user found"
            }
            ####### End code for object $x
            }
        if ($choiceLoop -eq 2)
            {
            write-host ("Entry "+$i+" skipped.")
            }
        if ($choiceLoop -eq 3)
            {
            write-host "Aborting."
            break
            }
        }
    } ## continue choices
    Write-Host "Summary"
    Write-Host "Changes made: $($change_i)"
    Write-Host "Removing any open sessions..."
    Get-PSSession 
    Get-PSSession | Remove-PSSession
    Write-Host "------------------------------------------------------------------------------------"
    $message ="Done. " +$processed+" of "+$entriescount+" entries processed. Press [Enter] to exit."
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

PauseTimed -quiet 3 #$message
Pause