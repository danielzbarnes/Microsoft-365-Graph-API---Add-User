<#
  .SYNOPSIS
      Creates a user in Microsoft 365 using Graph API without requiring any powershell modules installed.
  .DESCRIPTION
      This script is designed to take input from a client ticket support system(DeskDirector, ZenDesk, CloudRadial, etc) to create a Microsoft 365 account.
      It's intented to be used with the standard version of Powershell ISE installed with windows by copy pasting the entire input block from the ticket into terminal.
      All actions are executed using standared Invoke-RestMethod to call Graph API without requiring powershell modules to be installed.
      
      This requires a Service Principal account be setup in Microsoft Entra ID with the Permissiones Required:
        -Directory.ReadWrite.All
        -UserAuthenticationMethod.ReadWrite.All
        
      Actions taken by this script:
        -create user account
        -assign manager
        -assign groups
        -assign authentication phone number
        -assign M365 license
  .NOTES
      By default Microsoft Entra ID Service Principals authentication tokens are valid for only one hour.
      This script automatically refreshes the token when run if the Powershel ISE remains open.
  
      Input for this script was designed to work with tickets created from DeskDirector.

      Example Input:      
          ### Employee's First Name
              John
          ### Employee's Last Name
              Doe
          ### Manager Name
              Jane Doe
          ### Mobile Phone Number
              +1 (555) 555-5555
          ### Job Title
              Technician
          ### Department
              Engineering
          ### Office
              Los Angeles
          ### M365 Teams/Groups/Distribution Lists
              Engineering Team, CNC Group, etc
#>

# User details object
$userDetails = [PSCustomObject]@{

    firstName = ''
    lastName = ''
    manager = ''
    title = ''
    division = ''
    office = ''
    department = ''
    phone = ''
    grpsList = ''
    additionalNotes = ''
}

# Object to hold details for reporting
$userCreationDetails =  [PSCustomObject]@{

    id = ''
    displayName = ''
    upn = ''
    authPhone = ''
    office = ''
}

# initialize empty groups list
$grpsList        = New-Object -TypeName System.Collections.ArrayList
$grpSuccessList  = New-Object -TypeName System.Collections.ArrayList
$grpFailList     = New-Object -TypeName System.Collections.ArrayList
$licenseConsumed = New-Object -TypeName System.Collections.ArrayList
$failedLicenses  = New-Object -TypeName System.Collections.ArrayList
$UserUri = 'https://graph.microsoft.com/v1.0/users' # base user uri

function Get-Auth{
    # Prompt user for the Service Principal password and store as a secure string

    do{
        $global:userClientSecret = Read-Host -Prompt "Enter the Authorization Secret Value for the Identity Service Principal" -AsSecureString

        if ($global:userClientSecret.Length -eq 0) {
    
            Write-Host "No input was entered!" -ForegroundColor Red
            Write-Host "Please enter the API secret key."
        }

    } until($global:userClientSecret.Length -gt 0)
}

function Get-Secret{
    # helper function to read the password from the secure string
    
    return [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($global:userClientSecret))
}

function Get-MSToken{
    # function to retreive the authentication token
    # fillin the Service Principal details here

    $clientID = 'clientID'
    $tenantName = 'somedomain.com'
    $tenantID = 'tenantID'
    $scopeURL = 'https://graph.microsoft.com/.default' 

    $body = @{
        Grant_Type    = "client_credentials"
        Scope         = $scopeURL
        client_Id     = $clientID
        Client_Secret = Get-Secret
    }

    return Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token" -Method POST -Body $body
}

function Perform-RESTcall {
    <#
        .SYNOPSIS
            Reusable function to perform all REST calls 
        .DESCRIPTION
            Handles all REST API calls, uses a switch to handle diacritic characters that are not permitted in M365 User Principal Names 
    #>
    [cmdletbinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$uri,
        [Parameter(Mandatory=$true)]
        [string]$method,
        [switch]$upn, # switch to handle diacritic letters found in latin languages(ï, ë, ç)
        $body
    )

    $headers = @{Authorization = "Bearer $($global:userSession.access_token)"; "ConsistencyLevel"="eventual"; Accept = "application/json"}
    $contentType = 'application/json'
    
    if ($upn) { # content type needs to be modified to use charset utf-8, but the charset causes issues with other API calls
        $contentType = 'application/json; charset=utf-8'
    }

    return (Invoke-RestMethod -Uri $uri -headers $headers -ContentType $contentType -Method $method -body $body)
}

function MyPause { 
    <#
      .SYNOPSIS
          Helper function to pause script execution
      .DESCRIPTION
          This function uses a short pause to make output more readable as the script is executing.
          The long pause is needed to prevent REST API calls from failing due to M365 user account being unready to be updated.
    #>
    param (
        [switch]$long
    )
    
    $duration = if($long) { 2000 } else { 500 }
    Start-Sleep -Milliseconds $duration
}

function Initial-Message{
    # Function to display an initial message to the Technician executing the script
    # Add instructions and examples here

    write-Host `n$('='*100)`n -ForegroundColor DarkGray
    Write-Host 'Display some initial message'
    write-Host `n$('='*100)`n -ForegroundColor DarkGray
}

function Display-Message {
    # script to create a header and display a message, with or without color

    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$msg,
        $color
    )
    write-Host `n$('='*100)`n -ForegroundColor DarkGray

    if ($color) { 
        Write-Host $msg -ForegroundColor $color
    } else {
        Write-Host $msg
    }
}

function Display-UserReport {
    # Final user report at the end of execution
    
    Display-Message "User Account Creation Report`n"

    Write-Host 'id:' $userCreationDetails.id
    MyPause
    Write-Host 'Display Name: ' -NoNewline
    Write-Host $userCreationDetails.displayName -ForegroundColor Green
    MyPause
    Write-Host 'UPN:' -NoNewline
    Write-Host $userCreationDetails.upn -ForegroundColor Green
    MyPause
    Write-Host 'Password: ' -NoNewline
    MyPause
    Write-Host 'Welcome1' -ForegroundColor Green
    MyPause
    Write-Host 'Office Location:' $userCreationDetails.office
    MyPause

    if ($userCreationDetails.authPhone -ne ''){
        Write-Host 'Phone Authentication added:' $userCreationDetails.authPhone
    } else {
        Write-Host 'No Phone Authentication added' -ForegroundColor Red
    }

    MyPause
    Write-Host "`nGroups Successfully Added:"
    loopPrint $grpSuccessList "Green"

    if ($grpFailList.Count -gt 0 ){
        Write-Host "`nGroups Failed to Add:"
        loopPrint $grpFailList "Red"
    }
    Write-Host "`nLicenses Consumed:"
    loopPrint $licenseConsumed "Green"

    if ($failedLicenses.Count -gt 0){
        Write-Host "`nLicenses Failed to Add:"
        loopPrint $failedLicenses "Red"
    }

    Write-Host "`nAdditional Notes:"
    Write-Host $userDetails.additionalNotes
}

function Get-Confirmation{
    # Helper function to confirm a decision
    
    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Prompt
    )

    while($true){    
        $input = Read-Host -Prompt $Prompt

        if ($input -eq 'y') {
            return $true
        } elseif ($input -eq 'n'){
            return $false
        } else {
            Write-Host "Please enter a valid selection." -ForegroundColor Red
        }
    }
}

function Concat-Name{
    <#
        .SYNOPSIS
            Helper function to handled multiple first or last names, hyphenated names and removes Apostrophe characters to create desired UPNs.
        .DESCRIPTION
            Handles users with two first names or last names or hyphenated names by removing spaces, hyphens and apostrophes.
            Spaces need to be removed for UPNs but M365 permits hyphens and apostrophes in UPNs.

            Example:
              First John Last-Doe  -> FirstJohn.LastDoe@domain.com
              Erin O'hara -> Erin.Ohara@domain.com
    #>
    param (
        [String]$name = ''    
    )
    
    if ($name.contains(' ') -or $name.Contains('-') -or $name.Contains("'")){

        [String]$delimiter =  if ($name.contains(' ')) {
                $name[$name.IndexOf(' ')]
            } elseif($name.contains('-')) {
                $name[$name.IndexOf('-')]
            } else {
                $name[$name.IndexOf("'")]
            }
        return ($name.Replace($delimiter,''))
    } else {
        return $name
    }
}

function Remove-Diacritics {
    <#
        .SYNOPSIS
            Helper function to remove diactric characters from User Principal Names
        .DESCRIPTION
            Microsoft 365 does not allow diacritic characters found in Latin and Greek names(Ã¥, Ã¤, Ã¶) in User Principal Names.
            These characters need to be normalized to standard English characters.      
    #>
    param (
        [String]$src = [String]::Empty
    )
        $normalized = $src.Normalize( [Text.NormalizationForm]::FormD )
        $sb = new-object Text.StringBuilder
        
        $normalized.ToCharArray() | % { 
            if( [Globalization.CharUnicodeInfo]::GetUnicodeCategory($_) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
                [void]$sb.Append($_)
            }
        }
    return $sb.ToString()
}

function psObjPrint {
    # Helper function to print the properties of a PSObject
    
    param (
        [PSCustomObject]$obj
    )
    
    $properties = $obj | Get-Member -MemberType NoteProperty

    foreach ($prop in $properties){

        Write-Host "$($prop.name) :: $($obj.$($prop.Name))"
        MyPause
    }
}

function loopPrint {
    # Helper function to print entries from a list, in color if needed

    param ( 
      $list,
      $color
    )
    
    foreach ($item in $list){
    
        if ($color) {
            Write-Host $item -ForegroundColor $color 
        }
        else {
            Write-Host $item
        }
        MyPause
    }
}

function ProcessRawData {

    <#
        .SYNOPSIS
            Function to process the raw input string from the onboarding ticket
        .DESCRIPTION
            Input received from onboarding tickets is stored as single string that needs to be broken up and parse for user information.
            
            Takes the example input block:
            
                ### Employee's First Name
                    John
                ### Employee's Last Name
                    Doe
                ### Manager Name
                    Jane Doe
                ### Mobile Phone Number
                    +1 (555) 555-5555
                ### Job Title
                    Technician
                ### Department
                    Engineering
                ### Office
                    Los Angeles
                ### M365 Teams/Groups/Distribution Lists
                    Engineering Team, CNC Group, etc

            And will create an Array of strings with each entry on a single line:

                ### Employee's First Name:::John
                ### Employee's Last Name:::Doe
                ### Manager Name:::Jane Doe
                ....
    #>
    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$rawData
    )

    $rawArray = $rawData -split "`r`n" -ne ''
    $headerLine = ''
    $mergedList = New-Object -TypeName System.Collections.ArrayList
    $headerIndex = 0
    $mergedIndex = 0

    for ($i = 0; $i -le $rawArray.Count-1; $i++){

        $line = $rawArray[$i].Trim()

        if($line.startswith("###")){
            
            $headerIndex = $i
            $headerLine = $line+":::" #adds a custom delimiter to make parsing everything following the ### FieldName
            $rawArray[$i] = $headerLine

        } else {

            $prevEle = $rawArray[$headerIndex]
            $mergedIndex = $mergedList.add("$prevEle$line")

            # through parsing some fields such as Groups where there are lists, a delimiter needs to be insterted to prevent groups from being squished together
            # ie, Groups: group1, group2 from becoming group1_group2
            if ($mergedIndex -ne 0 -and $mergedList[$mergedIndex-1].Contains($prevEle)){

                $mergedList[$mergedIndex-1] += ('||'+$line) 
                $mergedList.RemoveAt($mergedIndex) # remove the duplicate line
            }
        }
    }
    return $mergedList
}

function Build-UserObj {
    <#
        .SYNOPSIS
            Function to parse the User Array to build the user object
        .DESCRIPTION
            Parses through the Array returned by ProcessRawData to fillin the data for the $userDetails object.
            The field names need to be changed to work with the input from whatever the ticket creation source is
    #>
    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        $userDataList
    )

    # switch case to fillin the $userDetails parameters. Searches by field name then parses for everything following the ":::" delimiter
    switch -Wildcard ($userDataList){
        
        "*First Name*" { 
            # if name contains ' ' or '-' then remove(assumes will either be one or the other)
            $userDetails.firstName = (Concat-Name ($_.substring($_.IndexOf(':::')+3)))
        }
        "*Last Name*"  { 
            $userDetails.lastName = (Concat-Name ($_.substring($_.IndexOf(':::')+3)))
        }
        "*Job title*"  { $userDetails.title = $_.substring($_.IndexOf(':::')+3) }
        "*Manager's*"    { $userDetails.manager = [regex]::Match($_, '[^:::]+(?= <)').value } # Manager field has both User Name and <email> parses for only the name
        "*division*"   { $userDetails.division = $_.substring($_.IndexOf(':::')+3) } 
        "*Personal Phone Number*" { $userDetails.phone = $_.substring($_.IndexOf(':::')+3) }
        "*Department*" {
            if ($_.contains(">")){ # When DeskDirector uses Other with written in data it adds an additional line with a ">" char
                $userDetails.department = $_.substring($_.IndexOf('> ')+2)
            }else {
                if ($userDetails.department -eq ''){
                    $userDetails.department = $_.substring($_.IndexOf(':::')+3)
                }
            }
        }
        "*Office*" {
            if ($_.contains(">")){
                $userDetails.office = $_.substring($_.IndexOf('> ')+2)
            } else {
                if ($userDetails.office -eq ''){
                    $userDetails.office = $_.substring($_.IndexOf(':::')+3)
                }
            }
        }
        "*Distribution List Access*"{
            
            # This field is a text box that users can enter single line lists or multi-line lists
            $inputGroupList = $_ -split ":::" # this returns an array with the header[0] and the list[1]
            $probChars = '[.;:]'

            $inputGroupsStr = $inputGroupList[1] -replace $probChars, '' # strip out any misc chars that can cause problems

            # this handles the multi-line list group
            if ($_.contains(',') -or $_.contains('||')){ 

                $inputGroupsStr = $inputGroupsStr.replace(',', '||')

                foreach ($item in ($inputGroupsStr.split('||'))){
                    if ($item.Trim().length -ne 0){
                        $grpsList.Add($item.Trim()) | Out-Null
                    }
                }
            } else { # this assume 1 group per line

                $singleGroup = $_ -split ":::"
                $grpsList.add($singleGroup[1]) | Out-Null
            }
        }
        "*Additional Requests*" {
            $userDetails.additionalNotes = $_.substring($_.IndexOf(':::')+3)
        }
    }
}

function Build-GroupList{
    <#
        .SYNOPSIS
            Function to build the list of groups
        .DESCRIPTION
            This depends largely on the groups in the organizations tenant and how to decide the job criteria applied to the user account

            Example:
            if (Title -eq Engineer) $grpsList.Add("Engineering Group") | Out-Null
            if (Location -eq Los Angeles) $grpsList.Add("California Group") | Out-Null
    #>
    
    $grpsList.Add('<add tenant groups needed based on organization criteria>') | Out-Null
    
    $userDetails.GrpsList = $grpsList
}

function Assign-License {
    <# 
        .SYNOPSIS
            Function to assign M365 license to newly created user account
        .DESCRIPTION
            At time of writing, Graph API cannot return a specific license skuid in the tenant.
            ALL the licenses need to be retreived and parsed for the needed skuid
            Licenses also need to be parsed by Sku code a complete list can be found here:
            https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference

            A few common license names and codes:
              Microsoft 365 E3 - SPE_E3
              Microsoft 365 E5 Security - IDENTITY_THREAT_PROTECTION
              Microsoft 365 F3 - SPE_F1
              Microsoft 365 F5 Security Add-On - SPE_F5_SEC
              Microsoft 365 Business Basic - O365_BUSINESS_ESSENTIALS
              Enterprise Mobility + Security E3 - EMS
              teams audio confrencing - Microsoft_Teams_Audio_Conferencing_select_dial_out
    #>

    Display-Message 'Assign Licenses Operation'
    MyPause

    # this returns ALL the license SKUs in the tenant because ?$search is broken at time of writing
    $skuSearchUri = 'https://graph.microsoft.com/v1.0/subscribedSkus?$select=skuid,skuPartNumber,consumedUnits,prepaidUnits'
    $addLicenseUri = 'https://graph.microsoft.com/v1.0/users/'+$userCreationDetails.id+'/microsoft.graph.assignLicense'

    # the user friendly names are NOT included in the sku objects
    $licenseNames = @{
        'SPE_E3' = 'Microsoft 365 E3'
        'IDENTITY_THREAT_PROTECTION' = 'Microsoft 365 E5 Security'
        'Microsoft_Teams_Audio_Conferencing_select_dial_out' = 'Microsoft Teams Audio Conferencing includes dial-out to USA/CAN only'
        'SPE_F1' = 'Microsoft 365 F3'
        'SPE_F5_SEC' = 'Microsoft 365 F5 Security Add-on'
        'O365_BUSINESS_ESSENTIALS' = 'Microsoft 365 Business Basic'
        'EMS' = 'Enterprise Mobility + Security E3'
    }

    $standardLicenseIDs = @('SPE_E3')
    # use logic to assign desired license based on job position, etc

    $license = @()
    $licenseHashList = @()

    Write-Host "`nUser Type: " -NoNewline

    Write-Host '<determine user licenses needs based on org needs>'

    MyPause
    Write-Host "Licenses Required:"
    loopPrint ($licenseNames.GetEnumerator() | Where-Object { $_.key -in $license } | Select-Object -ExpandProperty Value) "Yellow"
    MyPause

    if ($license.Count -gt 0){
    
        MyPause
        # get all the licenses then parse for only the ones that are needed
        $skus = (Perform-RESTcall $skuSearchUri "Get").value  | Where-Object { $license -contains $_.skuPartNumber }
        
        Write-Host "`nChecking Licenses:"

        foreach ($sku in $skus){

            Write-Host "$($licenseNames[$sku.skuPartNumber]): " -NoNewline
            if ($sku.consumedUnits -lt $sku.prepaidUnits.enabled){
                
                # Write-Host 'Used License: '$sku.consumedUnits # don't think this matters
                Write-Host 'Available Licenses:'($sku.prepaidUnits.enabled - $sku.consumedUnits) -f Green
                MyPause

                $hashID = @{'skuId' = $sku.skuid}
        
                $licenseHashList += $hashID
                $licenseConsumed.add($licenseNames[$sku.skuPartNumber]) | Out-Null
        
            } else {

                MyPause
                Write-Host 'No licenses available!' -BackgroundColor Red
                Write-Host $licenseNames[$sku.skuPartNumber] -NoNewline -ForegroundColor Red
                Write-Host ' Used Licenses: ' -NoNewline
                Write-Host ($sku.consumedUnits) -NoNewline -ForegroundColor Red
                Write-Host "/$($sku.prepaidUnits.enabled)"
                $failedLicenses.add($licenseNames[$sku.skuPartNumber]) | Out-Null
            }
            MyPause
        }

        $licenseJson = @{
            'addLicenses' = $licenseHashList
            'removeLicenses' = @()
        } | ConvertTo-Json
        
        Perform-RESTcall $addLicenseUri "Post" $licenseJson | Out-Null

        Write-Host "`nLicenses Successfully Assigned." -ForegroundColor Green

    } else {
        Write-Warning "something went wrong with the licensing."
        Write-Host 'No license applied.' -ForegroundColor Red
    }
}

function Assign-Groups{
    <#
        .SYNOPSIS
            Assigns groups to the user account
        .DESCRIPTION
            This function only assigns the groups to the user.
            The list is built in the Build-GroupsList function and added to the $userDetails object
    #>


    Display-Message "Attempting to Assign Groups`n"

    $baseGrpUri = 'https://graph.microsoft.com/v1.0/groups'
    $userOdata = 'https://graph.microsoft.com/v1.0/directoryObjects/'+$userCreationDetails.id

    $userJson = @{ '@odata.id'= $userOdata } | ConvertTo-Json

    foreach ($grp in $userDetails.grpsList){

        MyPause
        $groupSearchUri = ''

        # Search for email address because it's likely a distribution list
        if ($grp.contains('@')){

            # groupTypes = ['unified'] for M365, all others types = []
            # MailEnable = True, SG = True or False, means either SG/DL

            MyPause
            $mailUri = $baseGrpUri+"?`$filter=mail eq '$grp'&`$select=id,groupTypes,mailEnabled,securityEnabled"
            $mailSr = (Perform-RESTcall $mailUri "Get").value # this returns an array

            if ($null -ne $mailSr[0].id){

                Write-Host 'Mail Group Found:'$grp -ForegroundColor Green

                # groupTypes = 1 for M365 group, mailEnabled = False for standard SG(not mail enabled)
                if (($mailSr[0].groupTypes.Count -gt 0) -or (-not $mailSr[0].mailEnabled)){ 

                    MyPause
                    $addGrpUri = $baseGrpUri + '/' + $mailSr[0].id + '/members/$ref'
                    Perform-RESTcall $addGrpUri "Post" $userJson | Out-Null # if successful, call returns nothing

                    $grpSuccessList.add($grp) | Out-Null

                } elseif($mailSr[0].mailEnabled -and (-not $mailSr[0].securityEnabled)) {
                    MyPause
                    Write-Host "`nDistribution Group Detected!" -ForegroundColor Red
                    Write-Host 'Unable to add: ' -NoNewline
                    Write-Host $grp -ForegroundColor Red
                    $grpFailList.Add($grp) | Out-Null
                } elseif($mailSr[0].mailEnabled -and $mailSr[0].securityEnabled) {
                    MyPause
                    Write-Host "`nMail-Enabled Security Group Detected!" -ForegroundColor Red
                    Write-Host 'Unable to add: ' -NoNewline
                    Write-Host $grp -ForegroundColor Red
                    $grpFailList.Add($grp) | Out-Null
                }
            } else {
                Write-Host 'Group not found: '$grp -f Red
                $grpFailList.Add($grp) | Out-Null
            }
        } else {

            # check for "&" character because it will break the URI in the API call
            $callGroup = if ($grp.contains('&')) { $grp.replace('&','%26')} else { $grp }

            MyPause
            $nameUri = $baseGrpUri+"?`$filter=displayName eq '$callGroup'&`$select=id,displayName,mail"
            $nameSr = Perform-RESTcall $nameUri "Get"

            # this should always be a single object, but it might throw exception if returns more than one
            if ($null -ne $nameSr.value.id){

                MyPause
                Write-Host "Group Found: $grp" -ForegroundColor Cyan

                $addGrpUri = $baseGrpUri + '/' + $nameSr.value.id + '/members/$ref'
                Perform-RESTcall $addGrpUri "Post" $userJson | Out-Null

                $grpSuccessList.add($grp) | Out-Null

            } else {
                Write-Host 'Group not found: '$grp -f Red
                $grpFailList.Add($grp) | Out-Null
            }
        }
    }
    Write-Host "`nGroups Successfully Assigned." -ForegroundColor Green
}

function Assign-AuthPhone{
    <#
        .SYNOPSIS
            Function to add the phone number for Entra ID Phone Authentication.
        .DESCRIPTION
            A substantial delay is required when updating the User Account to add the Authentication Phone number.
            This may be due to lag between when the User Account is created and when it is ready to be modified.
    #>
    
    Display-Message "Assign Phone Authentication`n"

    # DeskDirector output forces the phone format as +1 (xxx) xxx-xxxx, but standard phone format should be acceptable
    if ($userDetails.phone -ne ''){

        Write-Host "Phone Number: " -NoNewline
        Write-Host $userDetails.phone -ForegroundColor Green
        MyPause
        Write-host "waiting for verify user account exists. standby."


        # This appears to need a minimum delay of 4 seconds to prevent failing
        # More delay was added as a precaution
        MyPause -long
        MyPause -long
        MyPause -long
        $addPhoneURI = 'https://graph.microsoft.com/v1.0/users/'+$userCreationDetails.id+'/authentication/phoneMethods'
        
        $phoneJson = @{
            phoneNumber = $userDetails.phone
            phoneType = 'mobile'
        } | ConvertTo-Json

        $phoneSr = Perform-RESTcall -uri $addPhoneURI -method "Post" $phoneJson
        $userCreationDetails.authPhone = $phoneSr.phoneNumber

        Write-Host "`nSuccessfully Assigned for Authentication Number." -ForegroundColor Green
        MyPause
        
    } else {
        Write-Warning 'No phone number detected from the ticket details.'
        Write-Host "`nNo Authentication Phone Number assigned to user" -ForegroundColor Red
    }
}
function Assign-Manager{
    <#
        .SYNOPSIS
            Function to assign a manager to the User Account.
        .DESCRIPTION
            This function has also been observed having issues with assigning the Manager after User Account creation.
    #>

    Display-Message "Assign Manager Operation`n"

    MyPause

    $mgrSearchUri = $UserUri+'?$filter=displayName eq '''+$userDetails.manager+'''&$select=id,displayName'
    $mgrSR = Perform-RESTcall -uri $mgrSearchUri -method "Get"

    if ($mgrSR.value.count -ne 0){

        Write-Host "Manager found: " -NoNewline
        Write-Host $userDetails.manager -ForegroundColor Green
        MyPause
        Write-Host "Attempting to assign Manager. standby."
        MyPause

        # Pauses here may or may not be necessary if called after Assign-AuthPhone
        # uncomment pauses if function fails for unclear reasons.
        #MyPause -long
        #MyPause -long
        $addMgrUri = $UserUri+'/'+$userCreationDetails.id+'/manager/$ref'
        $mgrUri = $UserUri+'/'+($mgrSR.value.id)
        $mgrJson = @{ '@odata.id'= $mgrUri } | ConvertTo-Json
        
        Perform-RESTcall -uri $addMgrUri -method "Put" -body $mgrJson

        Write-Host 'Manager Successfully Assigned.' -ForegroundColor Green

    } else {
        Write-Warning 'Manager not found!'
        Write-Host 'Please add Manager manually in Admin Center.' -ForegroundColor Yellow
    }
}

function Perform-CreateUser {
    # Function to perform User Account creation and call all assignment fuctions

    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        $newUserHash
    )
 
    MyPause -long
    $userRequest = Perform-RESTcall -uri $UserUri -method "Post" -body (ConvertTo-Json $newUserHash) -upn

    Display-Message "User successfully created.`n" "Green"
    psObjPrint $userRequest
    MyPause -long
    $userCreationDetails.id = $userRequest.id
    $userCreationDetails.displayName = $userRequest.displayName
    $userCreationDetails.upn = $userRequest.userPrincipalName
    $userCreationDetails.office = $userRequest.officeLocation

    Assign-AuthPhone
    Assign-Manager
    Assign-Groups
    Assign-License
}

function Create-User {
    <#
        .SYNOPSIS
            Function to being creating the User Account.
        .DESCRIPTION
            Builds user object to send to Perform-CreateUser.
            Verifies UPN availability here. If detected offers alternate UPN.
    #>

    Write-Host "`nStarting Script" -ForegroundColor Green
        
    $displayName = $userDetails.firstName, $userDetails.lastName -join ' '
    $userPrincipal = (Remove-Diacritics $displayName.Replace(' ','.'))+'@somedomain.com'

    $newUserHash = @{
        'accountEnabled'= $true
        'givenName' = $userDetails.firstName
        'surname' = $userDetails.lastName
        'displayName'= $displayName
        'mailNickname'= ($userPrincipal.substring(0,$userPrincipal.IndexOf('@'))) # alias for the user required for the Graph call
        'userPrincipalName' = $userPrincipal
        'officeLocation' = $userDetails.office
        'department' = $userDetails.department
        'jobTitle' = $userDetails.title
        'usageLocation' = "US" # Location is required to apply licensing
        'passwordProfile' = @{
            'password' = 'Welcome1' # password is hardcoded here, but can be generated randomly
            'forceChangePasswordNextSignIn' = $true
        }
    }

    $searchUri = $UserUri+"?`$filter=userPrincipalName eq '"+$userPrincipal+"'&`$select=id,userPrincipalName,givenName,surname,jobTitle"

    MyPause -long
    Display-Message ('Searching for existing user: '+$userPrincipal)

    $searchRequest = Perform-RESTcall -uri $searchUri -method "Get"
    
    $userCount = $searchRequest.Value.Count

    MyPause
    Write-Host "Number of users found: $userCount" # think this was just to check how many users it would return...

    # check if the user email exists
    if ($userCount -eq 0){

        MyPause
        Write-Host "User doesn't exist."
        MyPause
        Write-Host "Attempting to create user."

        MyPause
        Perform-CreateUser $newUserHash

    } else { # if user already exists, then handle it some way, add 01 to end of UPN or something

        MyPause
        Write-Warning "This user already exists"
        Write-Host "Please Verify this is NOT the same user." -ForegroundColor Yellow
        $searchRequest.value[0] | fl

        $newUpn = $userPrincipal.Replace('@', '01@')
        $prompt = 'Create new user using UPN: '+$newUpn+'?(y/n)'

        if (Get-Confirmation -Prompt $prompt){

            Write-Host "Checking new UPN availability."

            $newSearchUri = $UserUri+"?`$filter=userPrincipalName eq '"+$newUpn+"'&`$select=id,userPrincipalName,givenName,surname,jobTitle"
            $newSearch = Perform-RESTcall -uri $newSearchUri -method "Get"

            if ($newSearch.value.count -eq 0){

                MyPause
                Write-Host "$newUpn is available. `nAttempting to create account."
                $newUserHash['userPrincipalName'] = $newUpn
                $newUserHash['mailNickname'] = $newUpn.Substring(0,$newUpn.IndexOf('@'))
                
                Perform-CreateUser $newUserHash

            } else { # user.name01 already exists handle it some other way
                Write-Host "$newUpn exists also already!" -BackgroundColor Red
                Write-Warning "Username also already exists..."
                exit                
            }
        } else {
            Write-Host 'Aborting User Creation.' -ForegroundColor Red
        }
    }
}

try {

    Initial-Message

    if ($global:userSession){ # check if the session exists

        # refresh thee token using the previously entered Service Principal password
        $global:userSession = Get-MSToken

        $rawData = Read-Host -prompt "Enter the RAW User details from DeskDirector, ZenDesk, CloudRadial, etc"

        Build-UserObj (ProcessRawData $rawData)
        Build-GroupList

        Create-User

        Display-UserReport

        Display-Message "User Creation Operation Complete." "Green"

    }
    else{

        Write-Warning "The connection needs to be initialized." 

        # Prompt for Identity Service Principal authentication
        Get-Auth
        $global:userSession = Get-MSToken # store the Auth token in a global variable

        Write-Host "Restart the script to continue."
    }
}
catch {

    Write-Warning "An Error Occurred."
    Write-Warning $_.Exception.Message
    Write-Warning $_.Exception.GetType().FullName
    Write-Warning $_.CategoryInfo
    Write-Warning $_.FullyQualifiedErrorId
}
