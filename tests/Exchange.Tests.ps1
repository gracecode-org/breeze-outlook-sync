<#
.Synopsis
   Pester compatible module for testing the Exchange class.
.NOTES
    Author:  Chris D. Johnson
    Requires:  PowerShell V5
    
    Copyright (c) 2020 Chris D. Johnson All rights reserved.  
    Licensed under the Apache 2.0. See LICENSE file in the project root for full license information.  
#>

using module ..\Modules\config\
using module ..\Modules\Person\
using module ..\Modules\Tag\
using module ..\Modules\Exchange\
using module .\MockBreeze.psm1

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$Config = [Config]::new( $env:APPDATA + "\BreezeOutlookSync\config.json").GetConfigObject()

class Utility {
    static ValidateContact([Person] $p, [PSObject] $c) {
        $c | Should -not -BeNullOrEmpty
        $p.GetName() | Should -be $c.Name
        $p.GetName() | Should -be $c.Identity
        $p.GetFirstName() | Should -be $c.FirstName
        $p.GetLastName() | Should -be $c.LastName
        $p.GetDisplayName() | Should -be $c.DisplayName
        $p.GetStreetAddress() | Should -be $c.StreetAddress
        $p.GetCity() | Should -be $c.City
        $p.GetState() | Should -be $c.StateOrProvince
        $p.GetZip() | Should -be $c.PostalCode
        $p.GetHomePhone() | Should -be $c.Phone
        $p.GetMobilePhone() | Should -be $c.MobilePhone
        $p.GetWorkPhone() | Should -be $c.Office
    }

}

Describe "SyncContacts" {
    # Test variations
    # New/Existing Distribution List
    #  >0 valid MailContacts
    #  0 MailContacts
    # D New/Existing MailContact 
    # New/Existing Dist. List Member
    # D Existing MailContact, change email
    # D Existing MailContact, change name
    # Existing MailContact, add family member with same email.
    # Remove Dist. List (not exist in breeze)
    # Existing Dist. List, Remove member
    # Remove Contact (doesn't exist in breeze) : FUTURE

    $exchange
    $breeze

    BeforeAll {
        $exchange = [Exchange]::new($Config.Exchange.Connection.URI, `
            $Config.Exchange.Connection.username, 
            $Config.Exchange.Connection.password, $null, $null, $true)
        $breeze = [MockBreeze]::new()
    }

    AfterAll {
        $exchange.Disconnect()
    }

    Context "Contact Sync" {
        #$PSDefaultParameterValues = @{ 'It:Skip' = $true }

        $person = $null
        BeforeEach {
            $person = $breeze.GetPersonById([MockBreeze]::GetRawPersonIdFromIndex(0))
        }

        AfterEach {
            try {
                $exchange.RemoveMailContact($person.email)
            } catch {
            }
            $person = $null
        }


        It "New/Update MailContact" {
            $mailContact = $exchange.SyncContactFromBreezePerson($person)
            $mailContact | Should -not -BeNullOrEmpty
            $mailContact.PrimarySmtpAddress | Should -Be "testuser@yopmail.com"

            $mailContact = $exchange.GetMailContact("testuser@yopmail.com")
            $mailContact | Should -not -BeNullOrEmpty
            $mailContact.PrimarySmtpAddress | Should -Be "testuser@yopmail.com"

            $name = $person.GetName()
            $name | Should -be "Test User"

            $contact = $exchange.GetContact($name)
            [Utility]::ValidateContact($person, $contact)

            # Update the contact
            $mailContact = $exchange.SyncContactFromBreezePerson($person)
            $mailContact | Should -not -BeNullOrEmpty
            $contact = $exchange.GetContact($name)
            [Utility]::ValidateContact($person, $contact)


        }

        It "Update Various Info" {
            $mailContact = $exchange.SyncContactFromBreezePerson($person)
            $mailContact | Should -not -BeNullOrEmpty

            # Change various contact information
            $person.streetaddress = "1000 Rose Street"
            $person.city = "La Crosse"
            $person.state = "WI"
            $person.homephone = "(608) 555-5555"
            $person.mobile = "(608) 555-6666"
            $person.work = "(608) 555-7777"

            $mailContact = $exchange.SyncContactFromBreezePerson($person)
            $mailContact | Should -not -BeNullOrEmpty
            $contact = $exchange.GetContactFromEmail($person.GetFirstPrimaryEmail())
            [Utility]::ValidateContact($person, $contact)
        }

        It "Update the Email" {
            $mailContact = $exchange.SyncContactFromBreezePerson($person)
            $mailContact | Should -not -BeNullOrEmpty
            $contact = $exchange.GetContactFromEmail($person.GetFirstPrimaryEmail())
            [Utility]::ValidateContact($person, $contact)

            $person.email = "newemail@yopmail.com"
            $mailContact = $exchange.SyncContactFromBreezePerson($person)
            $mailContact | Should -not -BeNullOrEmpty
            $contact = $exchange.GetContactFromEmail("newemail@yopmail.com")
            $contact | Should -not -BeNullOrEmpty
            $name = $person.GetName()
            $name | Should -be "Test User"
            $contact = $exchange.GetContact($name)
            [Utility]::ValidateContact($person, $contact)
        }

        It "Update the Name" {
            $mailContact = $exchange.SyncContactFromBreezePerson($person)
            $mailContact | Should -not -BeNullOrEmpty
            $contact = $exchange.GetContactFromEmail($person.GetFirstPrimaryEmail())
            [Utility]::ValidateContact($person, $contact)

            $person.first = "New"
            $mailContact = $exchange.SyncContactFromBreezePerson($person)
            $mailContact | Should -not -BeNullOrEmpty
            $contact = $exchange.GetContactFromEmail("testuser@yopmail.com")
            $contact | Should -not -BeNullOrEmpty
            $name = $person.GetName()
            $name | Should -be "New User"
            $contact = $exchange.GetContact($name)
            [Utility]::ValidateContact($person, $contact)
        }
    }
}
Describe "SyncGroup" {
    $exchange
    $breeze

    BeforeAll {
        $exchange = [Exchange]::new($Config.Exchange.Connection.URI, `
            $Config.Exchange.Connection.username, 
            $Config.Exchange.Connection.password, $null, $null, $true)
        $breeze = [MockBreeze]::new()
    }

    AfterAll {
        $exchange.Disconnect()
    }

    Context "Distribution Group Sync" {
        AfterAll {
            Write-Host "Cleaning Up Test Users... "
            $persons = $breeze.GetPersonsFromTagId($breeze.GetTagByName("TEST").id, $true)
            foreach ($person in $persons) {
                try {
                    $exchange.RemoveMailContact($person.email)
                } catch {
                }
            }
            try {
                $exchange.RemoveGroup("TEST")
            } catch {
            }
        }

        It "Basic Dist Group" {
            $group = $exchange.CreateDistributionGroup("TEST", $null, "Test Notes")

            $group | Should -Not -BeNullOrEmpty

            # TODO: This is a pain: https://serverfault.com/questions/958967/powershell-get-notes-of-an-exchange-distribution-group
            #$group.Notes | Should -Be "Test Notes"

            $exchange.HasDistributionGroup("TEST") | Should -BeTrue

            $group = $exchange.GetGroup("TEST")
            $group | Should -Not -BeNullOrEmpty

            $exchange.RemoveGroup("TEST")

            $group = $exchange.GetGroup("TEST")
            $group | Should -BeNullOrEmpty

        }

        It "Group Membership" {
            # Create group and members
            $group = $exchange.CreateDistributionGroup("TEST", $null, $null)
            $group | Should -Not -BeNullOrEmpty

            $tagId = $breeze.GetTagByName("TEST").id
            $persons = $breeze.GetPersonsFromTagId($tagId, $true)
            $exchange.SyncDistributionGroupToBreezePersons("TEST", $persons, $null)
            [PSObject[]] $contacts = $exchange.GetDistributionGroupMembers("TEST")

            $contacts.length | Should -Be 3

            foreach ($person in $persons) {
                try {
                    $exchange.RemoveMailContact($person.email)
                } catch {
                }
            }

            # Remove one member and re-sync
            $firstPerson, $otherPersons=$persons

            $exchange.SyncDistributionGroupToBreezePersons("TEST", $otherPersons, $null)
            $contacts = $exchange.GetDistributionGroupMembers("TEST")
            $contacts.length | Should -Be 2

            # Remove another
            $exchange.SyncDistributionGroupToBreezePersons("TEST", $firstPerson, $null)
            $contacts = $exchange.GetDistributionGroupMembers("TEST")
            $contacts.length | Should -Be 1

            # Add back two members and re-sync
            $exchange.SyncDistributionGroupToBreezePersons("TEST", $persons, $null)
            $contacts = $exchange.GetDistributionGroupMembers("TEST")
            $contacts.length | Should -Be 3

            #Cleanup
            $persons = $breeze.GetPersonsFromTagId($tagId, $true)
            foreach ($person in $persons) {
                try {
                    $exchange.RemoveMailContact($person.email)
                } catch {
                }
            }

            $exchange.RemoveGroup("TEST")
        }
    }
}
