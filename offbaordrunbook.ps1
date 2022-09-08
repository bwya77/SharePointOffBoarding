Function Connect-GraphAPI {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory)]
        [string]$clientID,
        [Parameter(Mandatory)]
        [string]$tenantID,
        [Parameter(Mandatory)]
        [string]$clientSecret
    )
    begin {
        $ReqTokenBody = @{
            Grant_Type    = "client_credentials"
            Scope		  = "https://graph.microsoft.com/.default"
            client_Id	  = $clientID
            Client_Secret = $clientSecret
        }
    }
    process {

        $tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Method POST -Body $ReqTokenBody

    }
    end {
        return $tokenResponse
    }

}
Function Get-ListItems {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory)]
        [string]$siteID,
        [Parameter(Mandatory)]
        [string]$listID,
        [Parameter(Mandatory)]
        [string]$accessToken
    )
    begin {
        $headers = @{
            Authorization = "Bearer $accessToken"
        }
        $apiUrl = "https://graph.microsoft.com/v1.0/sites/$siteID/lists/$listID/items?expand=fields"
    }
    process {
        $listItems = Invoke-RestMethod -Uri $apiURL -Headers $headers -Method GET
    }
    end {
        return $listItems.value.fields
    }
}
function Searchfor-User
{
	param (
		[system.string]$UPN,
		[system.string]$AccessToken
	)
	Begin
	{
		$request = @{
			Method = "Get"
			Uri    = "https://graph.microsoft.com/v1.0/users/?`$filter=(userPrincipalName eq '$UPN')"
			ContentType = "application/json"
			Headers = @{ Authorization = "Bearer $AccessToken" }
		}
	}
	Process
	{
		$Data = Invoke-RestMethod @request
	}
	End
	{
        return $Data.value
	}
}
function Set-ListItemField
{
	Param (
		[Parameter(Mandatory)]
		[system.string]$AccessToken,
		[Parameter(Mandatory)]
		[System.String]$Field,
		[Parameter(Mandatory)]
		[System.Int32]$ItemNumber,
        [Parameter(Mandatory)]
        $Data,
        [Parameter(Mandatory)]
        [System.String]$SiteID,
        [Parameter(Mandatory)]
        [System.String]$ListID
	)
	Begin
	{
		If ($Field -eq "User") {
			$Body = @"
{
    "Title": "$Data"
}
"@
		} ElseIf ($Field -eq "Email") {
			$Body = @"
{
    "E_x002d_Mail": "$Data"
}
"@
		} ElseIf ($Field -eq "Notes") {
			$Body = @"
{
    "Notes": "$Data"
}
"@
		} ElseIf ($Field -eq "Status") {
			$Body = @"
{
    "Status": "$Data"
}
"@
		} ElseIf ($Field -eq "Licenses") {
			$Body = @"
{
    "Licenses": "$Data"
}
"@
        } ElseIf ($Field -eq "MailboxType") {
        $Body = @"
{
    "MailboxType": "$Data"
}
"@
	} ElseIf ($Field -eq "ForwardingAddress") {
        $Body = @"
{
    "ForwardingAddress": "$Data"
}
"@
	} ElseIf ($Field -eq "MailboxFullAccess") {
        $Body = @"
{
    "MbxFullAccess": "$Data"
}
"@
	} ElseIf ($Field -eq "Groups") {
        $Body = @"
{
    "Groups": "$Data"
}
"@
    }
}
	Process
	{
        $request = @{
			Method = "Patch"
			Uri    = "https://graph.microsoft.com/v1.0/sites/$siteID/lists/$listID/items/$itemnumber/fields"
			ContentType = "application/json"
			Headers = @{ Authorization = "Bearer $($AccessToken)" }
			Body   = $Body
		}
		$Response = Invoke-RestMethod @request
	}
	End
	{
        return $Response
	}
}
function Get-UserLicenses {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory)]
        [string]$userPrincipalName,
        [Parameter(Mandatory)]
        [string]$accessToken
    )
    begin {
        $headers = @{
            Authorization = "Bearer $accessToken"
        }
        $apiUrl = "https://graph.microsoft.com/v1.0/users/$userPrincipalName/licenseDetails"
    }
    process {
        $userLicenses = Invoke-RestMethod -Uri $apiURL -Headers $headers -Method GET
    }
    end {
        return $userLicenses.value
    }
}
function Get-MailboxSettings {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory)]
        [string]$userPrincipalName,
        [Parameter(Mandatory)]
        [string]$accessToken
    )
    begin {
        $headers = @{
            Authorization = "Bearer $accessToken"
        }
        $apiUrl = "https://graph.microsoft.com/beta/users/$userPrincipalName/mailboxSettings"
    }
    process {
        $mailboxSettings = Invoke-RestMethod -Uri $apiURL -Headers $headers -Method GET
    }
    end {
        return $mailboxSettings
    }
}
function Set-MailboxForwarding {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory)]
        [string]$accessToken,
        [Parameter(Mandatory)]
        [string]$userPrincipalName,
        [Parameter(Mandatory)]
        [string]$ForwardingAddress,
        [Parameter()]
        [string]$ForwardingName,
        [Parameter()]
        [string]$RuleName = 'Automation - Offboarding Forwarding'
    )
    begin {
        $headers = @{
            Authorization = "Bearer $($Token.access_token)"
        }

        $apiUrl = "https://graph.microsoft.com/v1.0/users/brad@thelazyadministrator.com/mailFolders/inbox/messageRules"
    }
    process {
        #Search for our user in Azure AD. If you dont care to have your user be an internal user, you can skip this part and remove it
        $FwdUser = Searchfor-User -UPN $ForwardingAddress -AccessToken $token.access_token
        #if we found our fwding user
        if ($FwdUser) {
            $ForwardingName = $FwdUser.displayName
            $ForwardingAddress = $FwdUser.mail

            $params = @{
                DisplayName = $RuleName
                Sequence = 1
                IsEnabled = $true
                Actions = @{
                    ForwardTo = @(
                        @{
                            EmailAddress = @{
                                Name = $ForwardingName
                                Address = $ForwardingAddress
                            }
                        }
                    )
                    StopProcessingRules = $true
                }
            }
            $body = $params | ConvertTo-Json -Depth 10
    
            $mailboxForwarding = Invoke-RestMethod -Uri $apiURL -Headers $headers -Method POST -Body $body -ContentType "application/json" 
        }
    }
    end {
        return $mailboxForwarding
    }
}
function Get-MailboxForwarding {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory)]
        [string]$userPrincipalName,
        [Parameter(Mandatory)]
        [string]$accessToken,
        [Parameter()]
        [string]$RuleName = 'Automation - Offboarding Forwarding'
    )
    begin {
        $headers = @{
            Authorization = "Bearer $accessToken"
        }
        $apiUrl = "https://graph.microsoft.com/v1.0/users/$userPrincipalName/mailFolders/inbox/messageRules"
    }
    process {
        $mailboxForwarding = Invoke-RestMethod -Uri $apiURL -Headers $headers -Method GET

    }
    end {
        return $mailboxForwarding.value | Where-Object {$_.DisplayName -eq $RuleName}
    }
}
function Remove-UserLicenses {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory)]
        [string]$userPrincipalName,
        [Parameter(Mandatory)]
        [string]$accessToken,
        [Parameter(Mandatory)]
        [string]$LicenseSkuID
    )
    begin {
        $headers = @{
            Authorization = "Bearer $accessToken"
        }
        $apiUrl = "https://graph.microsoft.com/v1.0/users/$userPrincipalName/assignLicense"
    }
    process {
        $body = @{
            addLicenses = @()
            removeLicenses= @($LicenseSkuID)
        } | ConvertTo-Json -Depth 10
        $removeLicense = Invoke-RestMethod -Uri $apiURL -Headers $headers -Method POST -Body $body -ContentType "application/json"
    }
    end {
        return $removeLicense
    }
}
function Get-GroupMembership {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory)]
        [string]$userPrincipalName,
        [Parameter(Mandatory)]
        [string]$accessToken
    )
    begin {
        $headers = @{
            Authorization = "Bearer $accessToken"
        }
        $apiUrl = "https://graph.microsoft.com/v1.0/users/$userPrincipalName/memberOf"
    }
    process {
        $groupMembers = Invoke-RestMethod -Uri $apiURL -Headers $headers -Method GET
    }
    end {
        return $groupMembers.value | where-object {$_.roleTemplateId -eq $null}
    }
}
function Remove-GroupMembership {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory)]
        [string]$userID,
        [Parameter(Mandatory)]
        [string]$accessToken,
        [Parameter(Mandatory)]
        [string]$GroupID
    )
    begin {
        $headers = @{
            Authorization = "Bearer $accessToken"
        }
        $apiUrl = "https://graph.microsoft.com/v1.0/groups/$GroupID/members/$userID/`$ref"
    }
    process {
        $removeGroupMember = Invoke-RestMethod -Uri $apiURL -Headers $headers -Method DELETE
    }
    end {
        return $removeGroupMember
    }
}

$clientId = Get-AutomationVariable -Name "clientID"
$tenantID = Get-AutomationVariable -Name "tenantID"
$clientSecret = Get-AutomationVariable -Name "clientSecret"

#Connect to MSGraph API
$token = Connect-GraphAPI -clientID $clientId -tenantID $tenantID -clientSecret $clientSecret

#Get all items within the SharePoint List
$items = Get-ListItems -listID '1baffb5c-d51a-4803-b534-2a83c3c867fd' -accessToken $token.access_token -siteID 'bwya77.sharepoint.com,218d5607-899f-4ec4-888a-0657c4fa2b11,af51a2a9-880d-4109-a7b1-84962fafb8a2' 
#Iterate through all users
foreach ($i in $items)
{
    #Get any and all notes that are already in the field so we do not overwrite anything 
    [array]$Notes = $i.Notes
    $licenseArray = @()
    #Search for our user
    $User = Searchfor-User -UPN $i.Title -AccessToken $token.access_token


    if ($i.status -eq "Pending")
    {
        #Only search of the user if we have not done it prior 
        if ($i.notes -notlike "*User was found in Azure AD*")
        {
            if ($User) {
                $Notes += "User was found in Azure AD`n"

                #Set the email field in the SharePoint List
                $Notes += "Email Address: $($User.mail)`n"
                Set-ListItemField -AccessToken $token.access_token -Field "Email" -ItemNumber $i.id -Data $User.Mail -listID '1baffb5c-d51a-4803-b534-2a83c3c867fd' -siteID 'bwya77.sharepoint.com,218d5607-899f-4ec4-888a-0657c4fa2b11,af51a2a9-880d-4109-a7b1-84962fafb8a2' 

                #Get all licenses for the user
                $Licenses = Get-UserLicenses -userPrincipalName $user.userPrincipalName -accessToken $token.access_token
                #Iterate through all licenses and create a clean array
                $Licenses | foreach-object {
                    $licenseArray += "$($_.skupartnumber) `n"
                }
                #Get the mailbox type for the user (the property will be userPurpose)
                $mailboxSettings = Get-MailboxSettings -userPrincipalName $user.userPrincipalName -accessToken $token.access_token
                 
                #Write the mailbox type for the user
                $Notes += "Mailbox Type: $($mailboxSettings.userPurpose)`n"
                Set-ListItemField -AccessToken $token.access_token -Field "MailboxType" -ItemNumber $i.id -Data $mailboxSettings.userPurpose -listID '1baffb5c-d51a-4803-b534-2a83c3c867fd' -siteID 'bwya77.sharepoint.com,218d5607-899f-4ec4-888a-0657c4fa2b11,af51a2a9-880d-4109-a7b1-84962fafb8a2'   
                
                #Write the licenses the user has back to the SharePoint List
                $Notes += "Licenses: $($licenseArray)`n"
                Set-ListItemField -AccessToken $token.access_token -Field "Licenses" -ItemNumber $i.id -Data $licenseArray -listID '1baffb5c-d51a-4803-b534-2a83c3c867fd' -siteID 'bwya77.sharepoint.com,218d5607-899f-4ec4-888a-0657c4fa2b11,af51a2a9-880d-4109-a7b1-84962fafb8a2' 

                #Get the groups the user is a member of
                $groupMembership = Get-GroupMembership -userPrincipalName $user.userPrincipalName -accessToken $token.access_token
                $grouparray = @()
                $groupMembership | foreach-object {
                    $Notes += "Adding the Group: $($_.displayName)`n"
                    $grouparray += "$($_.displayName) `n"
                }
                Set-ListItemField -AccessToken $token.access_token -Field "Groups" -ItemNumber $i.id -Data $grouparray -listID '1baffb5c-d51a-4803-b534-2a83c3c867fd' -siteID 'bwya77.sharepoint.com,218d5607-899f-4ec4-888a-0657c4fa2b11,af51a2a9-880d-4109-a7b1-84962fafb8a2' 
            }
            Else {
                $Notes += "User was not found in Azure AD`n"
                #Set the status to Error
                Set-ListItemField -AccessToken $token.access_token -Field "Status" -ItemNumber $i.id -Data "Error" -listID '1baffb5c-d51a-4803-b534-2a83c3c867fd' -siteID 'bwya77.sharepoint.com,218d5607-899f-4ec4-888a-0657c4fa2b11,af51a2a9-880d-4109-a7b1-84962fafb8a2' 
            }
        }
        #Only search of the forwarding user if we have not done it prior
        if ($i.notes -notlike "*Forwarding user was found in Azure AD*") {
            #See if the forwarding user is in Azure Active Directory
            $ForwardingUser = Searchfor-User -UPN $i.ForwardingAddress -AccessToken $token.access_token
            if ($ForwardingUser) {
                $Notes += "Forwarding user was found in Azure AD`n"
            }
            Else {
                $Notes += "Forwarding user was not found in Azure AD`n"
                #Set the status to Error
                Set-ListItemField -AccessToken $token.access_token -Field "Status" -ItemNumber $i.id -Data "Error" -listID '1baffb5c-d51a-4803-b534-2a83c3c867fd' -siteID 'bwya77.sharepoint.com,218d5607-899f-4ec4-888a-0657c4fa2b11,af51a2a9-880d-4109-a7b1-84962fafb8a2' 
            }
        }
        #If there were no errors, then change the status to Acknowledged
        $Notes += "Setting status to Acknowledged`n"
        Set-ListItemField -AccessToken $token.access_token -Field "Status" -ItemNumber $i.id -Data "Acknowledged" -listID '1baffb5c-d51a-4803-b534-2a83c3c867fd' -siteID 'bwya77.sharepoint.com,218d5607-899f-4ec4-888a-0657c4fa2b11,af51a2a9-880d-4109-a7b1-84962fafb8a2' 
    }
    ElseIf ($i.status -eq "Acknowledged")
    {
        #Figure out how many days and hours  until the user is to be off-boarded, if days left is less than or equal to 0 and hours is less than or equal to 0, then the user is to be off-boarded. NOTE: the default time in the timepicker is 7PM but can be changed in SharePoint
        $Timespan = New-TimeSpan -Start (Get-Date) -End $i.OffboardDate
        if (($Timespan.days -le 0) -and ($timespan.hours -le 0))
        {
            #Remove liceses from the user
            #Get all licenses for the user
            $Licenses = Get-UserLicenses -userPrincipalName $user.userPrincipalName -accessToken $token.access_token
            foreach ($license in $Licenses) {
                $Notes += "Removing $($license.skuPartNumber) license from $($i.Title)`n"
                Remove-UserLicenses -userPrincipalName $user.userPrincipalName -accessToken $token.access_token -licenseSkuID $license.skuId
            }

            #set the automatic mail forwarding rule
            $Notes += "Setting automatic mail forwarding rule to forward email to $($i.ForwardingAddress)`n"
            Set-MailboxForwarding -userPrincipalName $i.Title -accessToken $token.access_token -ForwardingAddress $i.ForwardingAddress
            $MailRuleCheck = Get-MailboxForwarding -userPrincipalName $i.Title -accessToken $token.access_token
            if ($MailRuleCheck) {
                $Notes += "Mail forwarding rule was set`n"
            }
            else {
                $Notes += "Mail forwarding rule was not set`n"
            }

            #Remove the user from the groups
            $groups = Get-GroupMembership -userPrincipalName $user.userPrincipalName -accessToken $token.access_token
            foreach ($group in $groups) {
                $Notes += "Removing $($user.DisplayName) from $($group.displayName)`n"
                Remove-GroupMembership -userID $user.id -accessToken $token.access_token -groupID $group.id
            }

        #If there were no errors, then change the status to Complete
        $Notes += "Setting status to Complete`n"
        Set-ListItemField -AccessToken $token.access_token -Field "Status" -ItemNumber $i.id -Data "Complete" -listID '1baffb5c-d51a-4803-b534-2a83c3c867fd' -siteID 'bwya77.sharepoint.com,218d5607-899f-4ec4-888a-0657c4fa2b11,af51a2a9-880d-4109-a7b1-84962fafb8a2' 

        }
    }
    Elseif ($i.status -eq "Error")
    {
        #If we could not find the user in Azure AD, attempt to self clear
        if ($i.notes -like "*User was not found*") {
            $User = Searchfor-User -UPN $i.Title -AccessToken $token.access_token
            if ($User) {
                $Notes = $Notes.Replace("User was not found in Azure AD","")
            }
        }
        #See if the error was because of the forwarding user not being in Azure AD
        if ($i.notes -like "*Forwarding user was not found*") {
            $ForwardingUser = Searchfor-User -UPN $i.ForwardingAddress -AccessToken $token.access_token
            if ($ForwardingUser) {
                $Notes = $Notes.Replace("Forwarding user was not found in Azure AD","")
            }
        }

        If ($Notes -notlike "*not*")
        {
            #If our notes contain no errors, we know all have cleared and we can set the status to Pending again
            Set-ListItemField -AccessToken $token.access_token -Field "Status" -ItemNumber $i.id -Data "Pending" -listID '1baffb5c-d51a-4803-b534-2a83c3c867fd' -siteID 'bwya77.sharepoint.com,218d5607-899f-4ec4-888a-0657c4fa2b11,af51a2a9-880d-4109-a7b1-84962fafb8a2' 
        }
    }

    #At the end: write all notes to the list
    Set-ListItemField -AccessToken $token.access_token -Field "Notes" -ItemNumber $i.id -Data $Notes -listID '1baffb5c-d51a-4803-b534-2a83c3c867fd' -siteID 'bwya77.sharepoint.com,218d5607-899f-4ec4-888a-0657c4fa2b11,af51a2a9-880d-4109-a7b1-84962fafb8a2' 

}
