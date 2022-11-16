<#
.Synopsis
   Class Module for interacting with Outlook 365 / Exchange Cmdlets
.NOTES
    Author:  Chris D. Johnson
    Requires:  PowerShell V5
    
    Copyright (c) 2020 Chris D. Johnson All rights reserved.  
    Licensed under the Apache 2.0. See LICENSE file in the project root for full license information.  
#>

using module ..\Person\
using module ..\Tag\
using module ..\Breeze\
using module ..\Logger\
using module ..\BreezeCache\

class Exchange {
    $DryRun = $false
    $GroupsList = $null
    $GroupEmailDomain = $null
    $BreezeCache = $null
    $Force = $false
    $Connected = $false


    Exchange([string] $connectionUri, [string] $user, [string] $password, 
      [string] $groupEmailDomain, [BreezeCache] $breezeCache, [boolean] $force, [string] $certThumbprint, [string] $appId, [string] $org) {
        [Logger]::Write("Connecting to Exchange...", $true)

        Connect-ExchangeOnline -CertificateThumbPrint $certThumbprint -AppID $appId -Organization $org
        $this.Connected = $true
        $this.GroupEmailDomain = $groupEmailDomain
        $this.BreezeCache = $breezeCache
        $this.Force = $force
    }

    [PSObject] GetDistributionGroups() {
        if ($null -eq $this.GroupsList) {
            [Logger]::Write("Fetching distribution groups...")
            $this.GroupsList = New-Object -TypeName "System.Collections.ArrayList"
            $groups = Get-DistributionGroup
            $this.GroupsList.AddRange($groups)
        }
        return $this.GroupsList
    }

    [PSObject] HasDistributionGroup([string] $groupName) {
        foreach ($group in $this.GetDistributionGroups()) {
            if ($group.DisplayName -eq $groupName) {
                return $true
            }
        }
        return $false;
    }

    [PSObject] GetGroup([string] $groupName) {
        return  Get-Group -Identity $groupName 
    }

    [PSObject] CreateDistributionGroup([string] $groupName, [string] $alias, [string] $notes) {
        <#
            .DESCRIPTION
            Creates a new Distribution Group, returing a DistributionGroup object that it created:
            https://docs.microsoft.com/en-us/previous-versions/office/exchange-server-api/ff327740(v=exchg.150)

            If specified, override the primary email domain
        #>
        if (-not $this.DryRun) {
            [Logger]::Write("Creating new distribution group: " + $groupName, $true)
            if([String]::IsNullOrEmpty($alias)) {
                $group = New-DistributionGroup -Name $groupName -Type "Distribution" `
                    -RequireSenderAuthenticationEnabled $false -Confirm:$false `
                    -Notes $notes                        

                #     # UnifiedGroups do NOT work with Mail Contacts
                #     #New-UnifiedGroup -Name $groupName -DisplayName $groupName -RequireSenderAuthenticationEnabled $false 
            } else {
                $group = New-DistributionGroup -Name $groupName -Type "Distribution" `
                    -RequireSenderAuthenticationEnabled $false -Confirm:$false `
                    -Notes $notes -Alias $alias
                #     # UnifiedGroups do NOT work with Mail Contacts
                #     #New-UnifiedGroup -Name $groupName -DisplayName $groupName -RequireSenderAuthenticationEnabled $false 
            }

            if($group -ne $null) {
                $domain = $this.GroupEmailDomain
                $oldPrimaryEmail = $group.PrimarySmtpAddress
                if($oldPrimaryEmail -ne $null -and -not [string]::IsNullOrEmpty($domain)) {
                    $emailName = $oldPrimaryEmail.Split("@")[0]
                    $newPrimaryEmail = $emailName + "@" + $domain
                    [Logger]::Write("Setting distribution group email: " + $newPrimaryEmail, $true)
                    $group = Set-DistributionGroup -Identity $groupName  -Confirm:$false `
                        -PrimarySmtpAddress $newPrimaryEmail
                }
            } else {
                [Logger]::Write("New-DistributionGroup returned null", $true)
            }

            if ($group -ne $null -and $this.GroupList -ne $null) {
                $this.GroupsList.Add($group)
            }
            return $group
        }
        return $null
    }

    [void] RemoveGroup([string] $groupName) {
        Remove-DistributionGroup -Identity $groupName -Confirm:$false
    }

    [PSObject] GetDistributionGroupMembers([string] $groupName) {
        <#
                .DESCRIPTION
                Returns all of the members of the specified group name in an array.  
                Type: 
                    ReducedRecipient:
                    https://docs.microsoft.com/en-us/previous-versions/office/exchange-server-api/ff346107(v=exchg.150)
                Common Parameters: 
                    Name
                    PrimarySmtpAddress
                    DisplayName
                    FirstName
                    LastName
                    StreetAddress
                    City
                    StateorProvince
                    PostalCode
                    Phone
                    MobilePhone
                    HomePhone
                    Company
                    Title
                    Notes
                    Office
            #>        
        return Get-DistributionGroupMember -Identity $groupName
    }

    static [boolean] HasDistributionGroupMember([PSObject] $groupMembers, [string] $memberEmail) {
        foreach ($groupMember in $groupMembers) {
            if ($groupMember.PrimarySmtpAddress -eq $memberEmail) {
                return $true
            }
        }
        return $false
    }

    [PSObject] GetMailContact([string] $memberEmail) {
        return Get-MailContact -Identity $memberEmail
    }

    [PSObject] GetContactFromEmail([string] $memberEmail) {
        return Get-Contact -Identity $memberEmail 
    }
    [PSObject] GetContact([string] $name) {
        return  Get-Contact -Identity $name
    }

    [PSObject] GetUserMailBox([string] $userEmail) {
        return Get-User -Identity $userEmail -RecipientTypeDetails UserMailBox
    }

    [PSObject] AddContactToDistributionGroup([string] $groupName, [string] $memberEmail) {
        <#
            .DESCRIPTION
            Adds an existing contact to the specified distribution group by contact email.
            Return Type: 
                ReducedRecipient:
                https://docs.microsoft.com/en-us/previous-versions/office/exchange-server-api/ff346107(v=exchg.150) 
        #>
        $result = $null
        [Logger]::Write("AddContactToDistributionGroup: " + $groupName + ", " + $memberEmail)
        if (-not $this.DryRun) {
            $result = Add-DistributionGroupMember  -Identity $groupName  -Confirm:$false -Member $memberEmail
        }
        [Logger]::Write("AddContactToDistributionGroup: result=" + $result)
        return $result
    }

    [void] SyncDistributionGroupFromTag([Tag] $tag) {
        if($null -eq $this.BreezeCache -or $this.Force -or $this.BreezeCache.HasTagChanged($tag)) {
            $persons = $tag.GetPersons()
            if($persons.Length -ne 0) {
                [Logger]::Write("Synchronizing tag: " + $tag.GetName() + "(" + $tag.GetId() + ") with " + $persons.Length + " persons..", $true)
                $notes = "breezetag:" + $tag.GetId()
                $this.SyncDistributionGroupToBreezePersons($tag.name, $persons, $notes)
            } else {
                [Logger]::Write("Skipping tag (no people with emails): " + $tag.GetName(), $true)
            }
            if($null -ne $this.BreezeCache) {
                $this.BreezeCache.CacheTag($tag)
            }
    } else {
            [Logger]::Write("Skipping unchanged tag: " + $tag.GetName(), $true)
        }

    }

    [void] SyncDistributionGroupToBreezePersons([string] $groupName, [Person[]] $persons, [string] $notes) {
        if (-not $this.HasDistributionGroup($groupName)) {
            [Logger]::Write("Creating new DistributionGroup: $groupName")
            $this.CreateDistributionGroup($groupName, $null, $notes)
        } else {
            [Logger]::Write("Distribution Group exists: " + $groupName, $false, [Logger]::LOGLEVEL_DEBUG)
        }

        # Verify that each person exists as a contact, but only if they have an email address
        $emailList = [System.Collections.ArrayList]::new()
        foreach ($person in $persons) {
            if (-not [string]::IsNullOrEmpty($person.GetFirstPrimaryEmail())) {
                $this.SyncContactFromBreezePerson($person)
                $emailList.Add($person.GetFirstPrimaryEmail())
            }
        }

        # Replace all members.
        # Remove any duplicates as a result of merging persons with emails.
        $emailList = $emailList | Sort-Object | Get-Unique
        [Logger]::Write("Updating DistributionGroupMembers: $emailList", $false, [Logger]::LOGLEVEL_DEBUG)
        Update-DistributionGroupMember  -Identity $groupName  -Confirm:$false -Members $emailList 
    }

    [PSObject] SyncContactFromBreezePerson([Person] $person) {
        <#
            Exchange contacts require a unique NAME and EMAIL.  

            The caller must pre-merge all Persons to make sure that they are already unique:
            1.  If multiple people share a userid, they are merged into one person.
            2.  If multiple people share the same name, then this will be an error for now.

            How to reconcile:
            Lookup the person by email and name
                If the contact is the same, then perform a simple field update.

                If the contacts are different or only one was found, then a family likely changed or someone got married.
                    Delete both contacts, and recreate a new one.

                If neither returns, just create a new contact.
        #>

        [Logger]::Write("Synchronzing Person: $person", $true)

        # Skip the check if force is disabled and the person hasn't changed
        if(-not $this.Force -and 
          ($null -eq $this.BreezeCache -or (-not $this.BreezeCache.HasPersonChanged($person)))) {
            [Logger]::Write("Skipping (unchanged).", $true)
            return $null
        }

        $name = $person.GetName()
        $displayname = $person.GetDisplayName()
        $firstname = $person.GetFirstName()
        $lastname = $person.GetLastName()
        $email = $person.GetFirstPrimaryEmail()
        $streetaddress = $person.GetStreetAddress()
        $city = $person.GetCity()
        $state = $person.GetState()
        $zip = $person.GetZip()
        $homephone = $person.GetHomePhone()
        $mobilephone = $person.GetMobilePhone()
        $workphone = $person.GetWorkPhone()
        $notes = "BREEZEID:" + $person.GetId()

        # Clean-up any discrepancies between contacts where names or email addresses don't match
        # If the contact is a User, then use them as-is
        $mailContact = $this.GetMailContact($email)
        $contact = $this.GetContact($name)

        [Logger]::Write("Found mailContact: $mailContact", $true)
        [Logger]::Write("Found contact: $contact", $true)
        
        # If we didn't find a contact or mailcontact, check if there is a UserMailbox
        # if there is, skip
        if ($null -eq $contact -and $null -eq $mailContact) {
            $user = $this.GetUserMailBox($email)
            if($null -ne $user) {
                [Logger]::Write("Skipping (Person is a User).", $true)
                return $user
            }
        }

        if ($null -eq $mailContact) {
            if ($null -ne $contact) {
                # Email changed
                $existingMailContact = $this.GetMailContact($name)
                $existingEmail = $existingMailContact.PrimarySmtpAddress
                [Logger]::Write("Deleting conflicting contact: $name, $existingEmail")
                if (-not $this.DryRun) {
                    Remove-MailContact -Identity $existingEmail -Confirm:$false
                    $contact = $null
                }
            }
        }
        else {
            if ($null -eq $contact) {
                # Name changed
                [Logger]::Write("Deleting conflicting mailcontact: $email")
                if (-not $this.DryRun) {
                    Remove-MailContact `
                            -Identity $email `
                            -Confirm:$false
                    $mailContact = $null
                }
            }
            elseif ($mailContact.Identity -ne $contact.Identity) {
                [Logger]::Write("Deleting conflicting contact: " + $email)
                if (-not $this.DryRun) {
                    Remove-MailContact `
                    -Identity $name `
                    -Confirm:$false
                }
                $contact = $null

                [Logger]::Write("Deleting conflicting mailcontact: " + $email, $true)
                if (-not $this.DryRun) {
                    Remove-MailContact `
                    -Identity $email `
                    -Confirm:$false
                }
                $mailContact = $null
            }
        }

        if ($null -eq $contact) {
            # Create the contact if it doesn't exist.
            [Logger]::Write("Creating new MailContact: $name, $displayname, $email, $firstname, $lastname")
            if (-not $this.DryRun) {
                $mailContact =  New-MailContact `
                    -Name $name `
                    -DisplayName $displayname `
                    -ExternalEmailAddress $email `
                    -FirstName $firstname `
                    -LastName $lastname

                if ($null -eq $mailContact) {
                    throw [System.ApplicationException]::new("Unable to create MailContact for person: $person")
                }
                [Logger]::Write("Setting contact info for $name")
                Set-Contact $name `
                    -StreetAddress  $streetaddress `
                    -City  $city `
                    -StateorProvince  $state `
                    -PostalCode  $zip `
                    -Phone  $workphone `
                    -MobilePhone  $mobilephone `
                    -HomePhone  $homephone `
                    -Notes  $notes `
                    -Office  $workphone                
            }
        }
        else {
            if (-not [Person]::ContactEquals($person, $contact)) {
                [Logger]::Write("Updating contact info: $email")
                Set-Contact $email `
                    -StreetAddress  $streetaddress `
                    -City  $city `
                    -StateorProvince  $state `
                    -PostalCode  $zip `
                    -Phone  $workphone `
                    -MobilePhone  $mobilephone `
                    -HomePhone  $homephone `
                    -Notes  $notes `
                    -Office  $workphone

            }
        }

        if($null -ne $this.BreezeCache) {
            $this.BreezeCache.CachePerson($person)
        }


        return $mailContact
    }

    [void] RemoveMailContact([string] $email) {
        [Logger]::Write("Deleting mailcontact: $email")
        if (-not $this.DryRun) {
            Remove-MailContact `
                -Identity $email `
                -Confirm:$false
        }
    }


    [void] Disconnect() {
        if ($this.Connected -eq $true) {
            [Logger]::Write("Disconnecting from Exchange...", $true)
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            $this.Connected = $false
        }
    }

}