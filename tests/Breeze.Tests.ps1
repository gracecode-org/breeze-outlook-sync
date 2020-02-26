<#
.Synopsis
   Pester compatible module for testing the Breeze class.
.NOTES
    Author:  Chris D. Johnson
    Requires:  PowerShell V5
    
    Copyright (c) 2020 Chris D. Johnson All rights reserved.  
    Licensed under the Apache 2.0. See LICENSE file in the project root for full license information.  
#>
using module ..\Modules\Config\
using module ..\Modules\Person\
using module ..\Modules\Tag\
using module ..\Modules\Breeze\
using module ..\Modules\BreezeCache\
using module .\MockBreeze.psm1

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$Config = [Config]::new($env:APPDATA + "\BreezeOutlookSync\config.json").GetConfigObject()

function GetTestPersons([Breeze] $breeze, [boolean] $setIds) {
    <#
    Create a bunch of test users
        | id | First | Middle | Last | Email | Home | Mobile | Office | Notes |
        | -- | ----- | ------ | ---- | ----- | ---- | ------ | ------ | ----- | 
        | 0 | Test |   | User | testuser@yopmail.com | (507) 123-9999 | (507) 123-0000 | (507) 123-8888 | |
        | 1 | Test | A | User | testauser@yopmail.com | (507) 123-9999 | (507) 123-0000 | (507) 123-8888 | Middle name differentiator from user 1 |
        | 2 | Test1 |  | User | testuser@yopmail.com | (507) 123-9999 | (507) 123-0000 | (507) 123-8888 | Same email as user 1
        | 3 | Test2A | | User | test2user@yopmail.com | (507) 123-9999 | (507) 123-0000 | (507) 123-8888 | Same email as user 4,5,6 | 
        | 4 | Test2B | | User | test2user@yopmail.com | (507) 123-9999 | (507) 123-0000 | (507) 123-8888 | Same email as user 4,5,6 |
        | 5 | Test2C | | User | test2user@yopmail.com | (507) 123-9999 | (507) 123-0000 | (507) 123-8888 | Same email as user 4,5,6 |
    #>

    $profileFieldsJSON=$breeze.GetProfileFieldsAsJSON()

    $persons = [Person[]] $(
        [Person]::new($profileFieldsJSON, 0, "Test", "Nick", $null, "User", "testuser@yopmail.com", "(507) 123-9999", "(507) 123-0000", "(507) 123-8888", "TESTING", "100 Main St", "Rochester", "MN", "55901", "TEST")
        [Person]::new($profileFieldsJSON, 0, "Test", "Nick", "A", "User", "testauser@yopmail.com", "(507) 123-9999", "(507) 123-0000", "(507) 123-8888", "TESTING", "100 Main St", "Rochester", "MN", "55901", "TEST")
        [Person]::new($profileFieldsJSON, 0, "Test1", $null, $null, "User", "testuser@yopmail.com", "(507) 123-9999", "(507) 123-0000", "(507) 123-8888", "TESTING", "100 Main St", "Rochester", "MN", "55901", "TEST")
        [Person]::new($profileFieldsJSON, 0, "Test2A", $null, $null, "User", "test2user@yopmail.com", "(507) 123-9999", "(507) 123-0000", "(507) 123-8888", "TESTING", "100 Main St", "Rochester", "MN", "55901", "TEST")
        [Person]::new($profileFieldsJSON, 0, "Test2B", $null, $null, "User", "test2user@yopmail.com", "(507) 123-9999", "(507) 123-0000", "(507) 123-8888", "TESTING", "100 Main St", "Rochester", "MN", "55901", "TEST")
        [Person]::new($profileFieldsJSON, 0, "Test2C", $null, $null, "User", "test2user@yopmail.com", "(507) 123-9999", "(507) 123-0000", "(507) 123-8888", "TESTING", "100 Main St", "Rochester", "MN", "55901", "TEST")
    )

    if($setIds) {
        $count=0
        foreach ($person in $persons) {
            $person.id = $count
            $count++
        }
    }
    return $persons
}



Describe "Person" {
    $TEST_PROFILE_FIELDS=[IO.File]::ReadAllText([MockBreeze]::PROFILE_FIELDS_FILE)

    It "GetProfileFieldId" {
        $p1 = [Person]::new($TEST_PROFILE_FIELDS, 0, "Test", "Nick", "Mid", "User", `
        "testuser@yopmail.com", "(507) 123-9999", "(507) 123-0000", "(507) 123-8888", "TESTING", `
        "100 Main St", "Rochester", "MN", "55901", "TEST")
        $p1.GetProfileFieldId("Main", "Comments") | Should -Be 2144404811
    }

    It "Constructors" {

        $p1 = [Person]::new($TEST_PROFILE_FIELDS, 12345678, "Test", "Nick", "Mid", "User", `
            "testuser@yopmail.com", "(507) 123-9999", "(507) 123-0000", "(507) 123-8888", "TESTING", `
            "100 Main St", "Rochester", "MN", "55901", @("TEST3", "TEST", "TEST2"))
        
        # Test the JSON Format from the people?json_filter API
        $p2PSObject = [PSCustomObject]@{
            id="12345678"
            first_name="Test"
            nick_name="Nick"
            force_first_name="Test"
            middle_name="Mid"
            last_name="User"
            details = [PSCustomObject]@{  # Nested details (from some APIs)
                details = [PSCustomObject]@{
                    email_primary="testuser@yopmail.com"
                    home="(507) 123-9999"
                    mobile="(507) 123-0000"
                    work="(507) 123-8888"
                    "2144404811"="TESTING"  # Comments
                    street_address="100 Main St"
                    city="Rochester"
                    state="MN"
                    zip="55901"
                    tagNames=@("TEST,TEST2,TEST3")
                }
            }
        }
        $p2 = [Person]::new($TEST_PROFILE_FIELDS, $p2PSObject)
        $p1.ContentEquals($p1) | Should -BeTrue
        $p1.ContentEquals($p2) | Should -BeTrue
        $p2.ContentEquals($p1) | Should -BeTrue

        # Test the JSON Format from the people/<id> API
        $p3PSObject = [PSCustomObject]@{
            id="12345678"
            force_first_name="Test"
            first_name="Nick"
            nick_name="Nick"
            middle_name="Mid"
            last_name="User"
            details = [PSCustomObject]@{
                1669889288 = [PSCustomObject]@(
                    [PSCustomObject]@{
                        address = "testuser@yopmail.com"
                        is_primary = "1"
                        allow_bulk = "1"
                        is_private = "0"
                        field_type = "email_primary"
                    }
                )
                2045627654 = [PSCustomObject]@(
                    [PSCustomObject]@{
                        field_type = "phone"
                        phone_number = "(507) 123-9999"
                        phone_type = "home"
                        do_not_text = "0"
                        is_private = "0"
                    },
                    [PSCustomObject]@{
                        field_type = "phone"
                        phone_number = "(507) 123-0000"
                        phone_type = "mobile"
                        do_not_text = "0"
                        is_private = "0"
                    },
                    [PSCustomObject]@{
                        field_type = "phone"
                        phone_number = "(507) 123-8888"
                        phone_type = "work"
                        do_not_text = "0"
                        is_private = "0"
                    }
                ) 
                2144404811="TESTING"  # Comments
                1429519026= [PSCustomObject]@(
                    [PSCustomObject]@{
                        field_type= "address_primary"
                        street_address= "100 Main St"
                        city= "Rochester"
                        state= "MN"
                        zip= "55901"
                        longitude= "-92.500123"
                        latitude= "44.07567"
                        is_primary= "1"
                        is_private= "0"
                    }
                )
                tagNames=@("TEST,TEST2,TEST3")
            }
        }
        $p3 = [Person]::new($TEST_PROFILE_FIELDS, $p3PSObject)        
        $p1.ContentEquals($p3) | Should -BeTrue
        $p3.ContentEquals($p1) | Should -BeTrue
    }

    It "HashCode" {
        $p1a = [Person]::new($TEST_PROFILE_FIELDS, 12345678, "Test", "Nick", "Mid", "User", `
            "testuser@yopmail.com", "(507) 123-9999", "(507) 123-0000", "(507) 123-8888", "TESTING", `
            "100 Main St", "Rochester", "MN", "55901", @("TEST3", "TEST", "TEST2"))

        $p1b = [Person]::new($TEST_PROFILE_FIELDS, 12345678, "Test", "Nick", "Mid", "User", `
        "testuser@yopmail.com", "(507) 123-9999", "(507) 123-0000", "(507) 123-8888", "TESTING", `
        "100 Main St", "Rochester", "MN", "55901", @("TEST3", "TEST", "TEST2"))

        $p1a.HashCode() | Should -Be $p1b.HashCode()

        $p2 = [Person]::new($TEST_PROFILE_FIELDS, 111, "Test", "Nick", "Mid", "User", `
            "testuser@yopmail.com", "(507) 123-9999", "(507) 123-0000", "(507) 123-8888", "TESTING", `
            "100 Main St", "Rochester", "MN", "55901", @("TEST3", "TEST", "TEST2"))

        $p1a.HashCode() | Should  -Not -Be $p2.HashCode()

        $p2 = [Person]::new($TEST_PROFILE_FIELDS, 12345678, "XX", "Nick", "Mid", "User", `
        "testuser@yopmail.com", "(507) 123-9999", "(507) 123-0000", "(507) 123-8888", "TESTING", `
        "100 Main St", "Rochester", "MN", "55901", @("TEST3", "TEST", "TEST2"))

        $p1a.HashCode() | Should  -Not -Be $p2.HashCode()

        $p2 = [Person]::new($TEST_PROFILE_FIELDS, 12345678, "Test", "xx", "Mid", "User", `
        "testuser@yopmail.com", "(507) 123-9999", "(507) 123-0000", "(507) 123-8888", "TESTING", `
        "100 Main St", "Rochester", "MN", "55901", @("TEST3", "TEST", "TEST2"))

        $p1a.HashCode() | Should  -Not -Be $p2.HashCode()

        $p2 = [Person]::new($TEST_PROFILE_FIELDS, 12345678, "Test", "Nick", "xx", "User", `
        "testuser@yopmail.com", "(507) 123-9999", "(507) 123-0000", "(507) 123-8888", "TESTING", `
        "100 Main St", "Rochester", "MN", "55901", @("TEST3", "TEST", "TEST2"))

        $p1a.HashCode() | Should  -Not -Be $p2.HashCode()

        $p2 = [Person]::new($TEST_PROFILE_FIELDS, 12345678, "Test", "Nick", "Mid", "xx", `
        "testuser@yopmail.com", "(507) 123-9999", "(507) 123-0000", "(507) 123-8888", "TESTING", `
        "100 Main St", "Rochester", "MN", "55901", @("TEST3", "TEST", "TEST2"))

        $p1a.HashCode() | Should  -Not -Be $p2.HashCode()

        $p2 = [Person]::new($TEST_PROFILE_FIELDS, 12345678, "Test", "Nick", "Mid", "User", `
        "x@yopmail.com", "(507) 123-9999", "(507) 123-0000", "(507) 123-8888", "TESTING", `
        "100 Main St", "Rochester", "MN", "55901", @("TEST3", "TEST", "TEST2"))

        $p1a.HashCode() | Should  -Not -Be $p2.HashCode()

        $p2 = [Person]::new($TEST_PROFILE_FIELDS, 12345678, "Test", "Nick", "Mid", "User", `
        "testuser@yopmail.com", "(111) 123-9999", "(507) 123-0000", "(507) 123-8888", "TESTING", `
        "100 Main St", "Rochester", "MN", "55901", @("TEST3", "TEST", "TEST2"))

        $p1a.HashCode() | Should  -Not -Be $p2.HashCode()

        $p2 = [Person]::new($TEST_PROFILE_FIELDS, 12345678, "Test", "Nick", "Mid", "User", `
        "testuser@yopmail.com", "(507) 123-9999", "(111) 123-0000", "(507) 123-8888", "TESTING", `
        "100 Main St", "Rochester", "MN", "55901", @("TEST3", "TEST", "TEST2"))

        $p1a.HashCode() | Should  -Not -Be $p2.HashCode()

        $p2 = [Person]::new($TEST_PROFILE_FIELDS, 12345678, "Test", "Nick", "Mid", "User", `
        "testuser@yopmail.com", "(507) 123-9999", "(507) 123-0000", "(111) 123-8888", "TESTING", `
        "100 Main St", "Rochester", "MN", "55901", @("TEST3", "TEST", "TEST2"))

        $p1a.HashCode() | Should  -Not -Be $p2.HashCode()

        $p2 = [Person]::new($TEST_PROFILE_FIELDS, 12345678, "Test", "Nick", "Mid", "User", `
        "testuser@yopmail.com", "(507) 123-9999", "(507) 123-0000", "(507) 123-8888", "XXX", `
        "100 Main StXX", "Rochester", "MN", "55901", @("TEST3", "TEST", "TEST2"))

        $p1a.HashCode() | Should  -Not -Be $p2.HashCode()

        $p2 = [Person]::new($TEST_PROFILE_FIELDS, 12345678, "Test", "Nick", "Mid", "User", `
        "testuser@yopmail.com", "(507) 123-9999", "(507) 123-0000", "(507) 123-8888", "TESTING", `
        "100 Main St", "XXX", "MN", "55901", @("TEST3", "TEST", "TEST2"))

        $p1a.HashCode() | Should  -Not -Be $p2.HashCode()

        $p2 = [Person]::new($TEST_PROFILE_FIELDS, 12345678, "Test", "Nick", "Mid", "User", `
        "testuser@yopmail.com", "(507) 123-9999", "(507) 123-0000", "(507) 123-8888", "TESTING", `
        "100 Main St", "Rochester", "XX", "55901", @("TEST3", "TEST", "TEST2"))

        $p1a.HashCode() | Should  -Not -Be $p2.HashCode()

        $p2 = [Person]::new($TEST_PROFILE_FIELDS, 12345678, "Test", "Nick", "Mid", "User", `
        "testuser@yopmail.com", "(507) 123-9999", "(507) 123-0000", "(507) 123-8888", "TESTING", `
        "100 Main St", "Rochester", "MN", "11111", @("TEST3", "TEST", "TEST2"))

        $p1a.HashCode() | Should  -Not -Be $p2.HashCode()

        # $p2 = [Person]::new($TEST_PROFILE_FIELDS, 12345678, "Test", "Nick", "Mid", "User", `
        # "testuser@yopmail.com", "(507) 123-9999", "(507) 123-0000", "(507) 123-8888", "TESTING", `
        # "100 Main St", "Rochester", "MN", "55901", @("XXX", "TEST", "TEST2"))

        # $p1a.HashCode() | Should  -Not -Be $p2.HashCode()

        # $p2 = [Person]::new($TEST_PROFILE_FIELDS, 12345678, "Test", "Nick", "Mid", "User", `
        # "testuser@yopmail.com", "(507) 123-9999", "(507) 123-0000", "(507) 123-8888", "TESTING", `
        # "100 Main St", "Rochester", "MN", "55901", @("TEST3", "TEST"))

        # $p1a.HashCode() | Should  -Not -Be $p2.HashCode()
    }

    It "GetDedupedPersons" {
        $breeze = [Mockbreeze]::new()
        $persons = GetTestPersons $breeze  $true

        [Person[]] $dedupedPersons = [Person]::GetDedupedPersons($persons)
        $dedupedPersons.length | Should -Be 3
        $dedupedPersons[0].id | Should -Be 0
        $dedupedPersons[1].id | Should -Be 1
        $dedupedPersons[2].id | Should -Be 3
    }

    It "GetMergedPersonsByEmail" {
        # Note:  These id's are from the mock data
        $breeze = [Mockbreeze]::new()
        $person = $breeze.GetMergedPersonsByEmail("testuser@yopmail.com")
        $personId = [MockBreeze]::GetRawPersonIdFromIndex(0)
        $person.id | Should -Be $personId
        $person.GetFirstName() | Should -Be "Test and Test1"
        
        $person = $breeze.GetMergedPersonsByEmail("testauser@yopmail.com")
        $personId = [MockBreeze]::GetRawPersonIdFromIndex(1)
        $person.id | Should -Be $personId
        $person.GetFirstName() | Should -Be "Test"

        $person = $breeze.GetMergedPersonsByEmail("test2user@yopmail.com")
        $personId = [MockBreeze]::GetRawPersonIdFromIndex(3)
        $person.id | Should -Be $personId
        $person.GetFirstName() | Should -Be "Test2A, Test2B and Test2C"
    }

    It "FirstPrimaryEmail" {
        $p1= [Person]::new($TEST_PROFILE_FIELDS, 11111, "Test", "Nick", "Mid", "User", `
        "testuser@yopmail.com", "(507) 123-9999", "(507) 123-0000", "(507) 123-8888", "TESTING", `
        "100 Main St", "Rochester", "MN", "55901", @("TEST1", "TEST2"))
        $p1.GetFirstPrimaryEmail() | Should -Be "testuser@yopmail.com"

        $p1= [Person]::new($TEST_PROFILE_FIELDS, 11111, "Test", "Nick", "Mid", "User", `
        "testuser@yopmail.com,a@test.com", "(507) 123-9999", "(507) 123-0000", "(507) 123-8888", "TESTING", `
        "100 Main St", "Rochester", "MN", "55901", @("TEST1", "TEST2"))
        $p1.GetFirstPrimaryEmail() | Should -Be "testuser@yopmail.com"

        $p1= [Person]::new($TEST_PROFILE_FIELDS, 11111, "Test", "Nick", "Mid", "User", `
        "", "(507) 123-9999", "(507) 123-0000", "(507) 123-8888", "TESTING", `
        "100 Main St", "Rochester", "MN", "55901", @("TEST1", "TEST2"))
        $p1.GetFirstPrimaryEmail() | Should -Be ""

        $p1= [Person]::new($TEST_PROFILE_FIELDS, 11111, "Test", "Nick", "Mid", "User", `
        "testuser@yopmail.com;a@test.com", "(507) 123-9999", "(507) 123-0000", "(507) 123-8888", "TESTING", `
        "100 Main St", "Rochester", "MN", "55901", @("TEST1", "TEST2"))
        $p1.GetFirstPrimaryEmail() | Should -Be "testuser@yopmail.com"

        
        $p1= [Person]::new($TEST_PROFILE_FIELDS, 11111, "Test", "Nick", "Mid", "User", `
        "  testuser@yopmail.com  ;a@test.com", "(507) 123-9999", "(507) 123-0000", "(507) 123-8888", "TESTING", `
        "100 Main St", "Rochester", "MN", "55901", @("TEST1", "TEST2"))
        $p1.GetFirstPrimaryEmail() | Should -Be "testuser@yopmail.com"
    }
}

Describe "BreezeCache" {
    $TEST_PROFILE_FIELDS=[IO.File]::ReadAllText([MockBreeze]::PROFILE_FIELDS_FILE)
    It "BreezeCachePersons" {
        $breezeCache = [BreezeCache]::new($here + "\test_personcache")
        $breezeCache.Clear()

        
        $p1 = [Person]::new($TEST_PROFILE_FIELDS, 12345678, "Test", "Nick", "Mid", "User", `
            "testuser@yopmail.com", "(507) 123-9999", "(507) 123-0000", "(507) 123-8888", "TESTING", `
            "100 Main St", "Rochester", "MN", "55901", @("TEST3", "TEST", "TEST2"))
        
        $breezeCache.HasPersonChanged($p1) | Should -BeTrue
        $breezeCache.CachePerson($p1)
        $breezeCache.HasPersonChanged($p1) | Should -BeFalse

        $p1a = [Person]::new($TEST_PROFILE_FIELDS, 12345678, "Testx", "Nick", "Mid", "User", `
        "testuser@yopmail.com", "(507) 123-9999", "(507) 123-0000", "(507) 123-8888", "TESTING", `
        "100 Main St", "Rochester", "MN", "55901", @("TEST3", "TEST", "TEST2"))

        $breezeCache.HasPersonChanged($p1a) | Should -BeTrue

    }

    It "BreezeCacheTags" {
        $breezeCache = [BreezeCache]::new($here + "\test_tagcache")
        $breezeCache.Clear()

        # Create our baseline cache
        $t1 = [Tag]::new(1, "TEST1")
        $t2 = [Tag]::new(2, "TEST2")

        $p1 = [Person]::new($TEST_PROFILE_FIELDS, 11111, "Test", "Nick", "Mid", "User", `
            "testuser@yopmail.com", "(507) 123-9999", "(507) 123-0000", "(507) 123-8888", "TESTING", `
            "100 Main St", "Rochester", "MN", "55901", @("TEST1"))
        $t1.AddPerson($p1).Length | Should -Be 1

        $p2 = [Person]::new($TEST_PROFILE_FIELDS, 22222, "Test", "Nick", "Mid", "User", `
            "testuser@yopmail.com", "(507) 123-9999", "(507) 123-0000", "(507) 123-8888", "TESTING", `
            "100 Main St", "Rochester", "MN", "55901", @("TEST2"))
        $t2.AddPerson($p2).Length | Should -Be 1


        $breezeCache.HasTagChanged($t1) | Should -BeTrue
        $breezeCache.HasPersonChanged($p1) | Should -BeTrue

        $breezeCache.CacheTag($t1)
        $breezeCache.HasTagChanged($t1) | Should -BeFalse
        $breezeCache.HasPersonChanged($p1) | Should -BeFalse

        $breezeCache.CacheTag($t2)
        $breezeCache.HasTagChanged($t2) | Should -BeFalse
        $breezeCache.HasPersonChanged($p2) | Should -BeFalse


        # Move to new tag
        $p1a = [Person]::new($TEST_PROFILE_FIELDS, 11111, "Test", "Nick", "Mid", "User", `
        "testuser@yopmail.com", "(507) 123-9999", "(507) 123-0000", "(507) 123-8888", "TESTING", `
        "100 Main St", "Rochester", "MN", "55901", @("TEST2"))
        $p2a = [Person]::new($TEST_PROFILE_FIELDS, 11111, "Test", "Nick", "Mid", "User", `
        "testuser@yopmail.com", "(507) 123-9999", "(507) 123-0000", "(507) 123-8888", "TESTING", `
        "100 Main St", "Rochester", "MN", "55901", @("TEST2"))
        $t1a = [Tag]::new(1, "TEST1")
        $t2a = [Tag]::new(2, "TEST2", @($p1a, $p2a))
        $breezeCache.HasTagChanged($t1a) | Should -BeTrue
        $breezeCache.HasPersonChanged($p1a) | Should -BeFalse
        $breezeCache.HasTagChanged($t2a) | Should -BeTrue
        $breezeCache.HasPersonChanged($p2a) | Should -BeFalse

        # Change a person in one of the tags
        $p1a = [Person]::new($TEST_PROFILE_FIELDS, 11111, "Testx", "Nick", "Mid", "User", `
        "testuser@yopmail.com", "(507) 123-9999", "(507) 123-0000", "(507) 123-8888", "TESTING", `
        "100 Main St", "Rochester", "MN", "55901", @("TEST1"))
        $t1a = [Tag]::new(1, "TEST1")
        $breezeCache.HasTagChanged($t1a) | Should -BeTrue
        $breezeCache.HasPersonChanged($p1a) | Should -BeTrue
        $breezeCache.HasTagChanged($t2) | Should -BeFalse
        $breezeCache.HasPersonChanged($p2) | Should -BeFalse

        # Add to new tag
        $p1a = [Person]::new($TEST_PROFILE_FIELDS, 11111, "Test", "Nick", "Mid", "User", `
        "testuser@yopmail.com", "(507) 123-9999", "(507) 123-0000", "(507) 123-8888", "TESTING", `
        "100 Main St", "Rochester", "MN", "55901", @("TEST1", "TEST2"))
        $t1a = [Tag]::new(1, "TEST1", @($p1a))
        $t2a = [Tag]::new(2, "TEST2", @($p1a, $p2))
        $breezeCache.HasTagChanged($t1a) | Should -BeFalse
        $breezeCache.HasPersonChanged($p1a) | Should -BeFalse
        $breezeCache.HasTagChanged($t2a) | Should -BeTrue
        $breezeCache.HasPersonChanged($p2) | Should -BeFalse

        # Add to new tag and modified person
        $p1a = [Person]::new($TEST_PROFILE_FIELDS, 11111, "Testx", "Nick", "Mid", "User", `
        "testuser@yopmail.com", "(507) 123-9999", "(507) 123-0000", "(507) 123-8888", "TESTING", `
        "100 Main St", "Rochester", "MN", "55901", @("TEST1", "TEST2"))
        $t1a = [Tag]::new(1, "TEST1", @($p1a))
        $t2a = [Tag]::new(2, "TEST2", @($p1a, $p2))
        $breezeCache.HasTagChanged($t1a) | Should -BeTrue
        $breezeCache.HasPersonChanged($p1a) | Should -BeTrue
        $breezeCache.HasTagChanged($t2a) | Should -BeTrue
        $breezeCache.HasPersonChanged($p2) | Should -BeFalse
    }
}

Describe "Breeze" {
    $breeze
    $testPersons
    $testTags
    
    BeforeAll {
        $breeze = [Breeze]::new($Config.Breeze.Connection.URI, $Config.Breeze.Connection.APIKey)



        Write-Host "Creating Test Users"
        try {
            $testTags = [Tag[]] $(
                [Tag]::new("TEST")
            )
            $breeze.AddTags($testTags)

            $testPersons = GetTestPersons $breeze  $false
            $breeze.SyncPersons($testPersons)

            # Uncomment to re-create test files to use in unit tests.
            #  Write-Host "Re-creating JSON payloads from breeze"
            # [IO.File]::WriteAllText("$here\Breeze.Tests.ProfileFields.json", $breeze.GetProfileFieldsAsJSON())
            # $count=0
            # foreach($person in $testPersons) {
            #     [IO.File]::WriteAllText("$here\Breeze.Tests.Person$count.json", $breeze.GetPersonByIdAsJSON($person.id))
            #     $count++
            # }
            # [IO.File]::WriteAllText("$here\Breeze.Tests.PersonsFromTagTEST.json", $breeze.GetPersonsFromTagIdAsJSON($testTags[0].id))

        } catch  {
            Write-Error "Caught exception: " +  $PSItem.Exception|format-list -force
            throw $PSItem.Exception
        }
    }

    AfterAll {
        Write-Host "Deleting Test Users"
        $breeze.DeletePersons($testPersons)
        $breeze.DeleteTags($testTags)
    }

    Context "Marshal Validation" {
        It "People Marshalled Correctly" {
            foreach($person in $testPersons) {
                $fetchedPerson = $breeze.GetPersonById($person.id)
                $person.ContentEquals($fetchedPerson) | Should -BeTrue
            }
        }

        It "Tag and People Tests" {
            $breeze | Should -Not -BeNullOrEmpty
            $t = $breeze.GetTagByName("TEST")
            $t.id | Should -Be $testTags[0].id
            [Person[]] $p = $breeze.GetPersonsFromTagId($t.id, $false)
            $p.Length | Should -Be $testPersons.Length
    
            [Person[]] $p = $breeze.GetPersonsFromTagId($t.id, $true)
            $p.Length | Should -Be 3
    
        }
    }
}
