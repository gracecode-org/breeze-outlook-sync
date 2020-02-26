<#
.Synopsis
   Class Module representing a Breeze Person.
.NOTES
    Author:  Chris D. Johnson
    Requires:  PowerShell V5
    
    Copyright (c) 2020 Chris D. Johnson All rights reserved.  
    Licensed under the Apache 2.0. See LICENSE file in the project root for full license information.  
#>
using assembly System.Web
using module ..\Logger\

class Person {
    # Phone number that accept a dot, a space, a dash, a forward slash, 
    # between the numbers. Will Accept a 1 or 0 in front. Area Code not necessary
    static [string] $PHONEREGEX = "((\(\d{3}\)?)|(\d{3}))([\s-./]?)(\d{3})([\s-./]?)(\d{4})"

    [int] $id = 0
    [ValidateNotNullOrEmpty()][string] $first
    [ValidateNotNullOrEmpty()][string] $last
    [string] $nickname
    [string] $middle
    [string] $email
    [string] $homephone
    [string] $mobile
    [string] $work
    [string] $streetaddress
    [string] $city
    [string] $state
    [string] $zip
    [string] $comments  # Custom
    [string[]] $tagNames

    hidden [ValidateNotNullOrEmpty()] [PSCustomObject] $ProfileFields
   
    # TODO Add nickname
    Person([string] $breezeProfileFields, [int] $id, [string] $first, [string] $nickname, [string] $middle, [string] $last, `
            [string] $email, [string] $homephone, [string] $mobile, [string] $work, [string] $comments, `
            [string] $streetaddress, [string] $city, [string] $state, [string] $zip, [string[]] $tagNames) {
        $this.id = $id
        $this.first = $first
        $this.nickname = $nickname
        $this.middle = $middle
        $this.last = $last
        $this.email = $email
        $this.homephone = $homephone
        $this.mobile = $mobile
        $this.work = $work
        $this.comments = $comments
        $this.streetaddress = $streetaddress
        $this.city = $city
        $this.state = $state
        $this.zip = $zip
        $this.tagNames = $tagNames
        $this.ProfileFields = $breezeProfileFields | ConvertFrom-Json
    }

    Person([string] $breezeProfileFields, [string] $breezeJSON) {
        $personPSObject = ConvertFrom-Json($breezeJSON)
        $this.Init($breezeProfileFields, $personPSObject)

    }

    Person([string] $breezeProfileFields, [PSObject] $personPSObject) {
        $this.Init($breezeProfileFields, $personPSObject)
    }

    hidden static [string] Trim([string] $s) {
        if ($s -eq $null) { return $s }
        return $s.Trim()
    }

    hidden [void] Init([string] $breezeProfileFields, [PSObject]  $personPSObject) {
        
        $this.ProfileFields = ConvertFrom-Json($breezeProfileFields)

        # Breeze has two formats for describing a person.  
        #  From the people?json_filter API
        #  From the people/<id> APi
        
        # Some APIs return the the person in a details object, so unwrap it.
        if ($personPSObject.details.details -ne $null) {
            # The Filter API format
            $this.id = $personPSObject.id
            $this.first = $personPSObject.force_first_name
            $this.nickname = $personPSObject.nick_name
            $this.last = $personPSObject.last_name
            $this.middle = $personPSObject.middle_name
            $this.email = $personPSObject.details.details.email_primary
            $this.streetaddress = $personPSObject.details.details.street_address
            $this.city = [Person]::Trim($personPSObject.details.details.city)
            $this.state = $personPSObject.details.details.state
            $this.zip = $personPSObject.details.details.zip
            if ($personPSObject.details.details.home -match [Person]::PHONEREGEX) {
                $this.homephone = $personPSObject.details.details.home
            }
            if ($personPSObject.details.details.work -match [Person]::PHONEREGEX){
                $this.work = $personPSObject.details.details.work
            }
            if ($personPSObject.details.details.mobile -match [Person]::PHONEREGEX){
                $this.mobile = $personPSObject.details.details.mobile
            }
            $fieldId = $this.GetProfileFieldId("Main", "Comments")
            $this.comments = $personPSObject.details.details."$fieldId"            
        }
        else {
            # The id API format
            $this.id = $personPSObject.id
            $this.first = $personPSObject.force_first_name
            $this.nickname = $personPSObject.nick_name
            $this.last = $personPSObject.last_name
            $this.middle = $personPSObject.middle_name
            $fieldId = $this.GetProfileFieldId("Contact", "Email")
            [PSCustomObject[]] $emails = $personPSObject.details."$fieldId"
            foreach ($email in $emails) {
                if ($email.field_type -eq "email_primary") {
                    $this.email = $email.address
                }
            }            
            $fieldId = $this.GetProfileFieldId("Contact", "Address")
            [PSCustomObject[]] $addresses = $personPSObject.details."$fieldId"
            foreach ($address in $addresses) {
                if ($address.field_type -eq "address_primary") {
                    $this.streetaddress = $address.street_address
                    $this.city = [Person]::Trim($address.city)
                    $this.state = $address.state
                    $this.zip = $address.zip
                }
            }            
            $fieldId = $this.GetProfileFieldId("Contact", "Phone")
            [PSCustomObject[]] $phones = $personPSObject.details."$fieldId"
            foreach ($phone in $phones) {
                if($phone.phone_number -match [Person]::PHONEREGEX) {
                    if ($phone.phone_type -eq "home") {
                        $this.homephone = $phone.phone_number
                    }
                    elseif ($phone.phone_type -eq "mobile") {
                        $this.mobile = $phone.phone_number
                    }
                    elseif ($phone.phone_type -eq "work") {
                        $this.work = $phone.phone_number
                    }
                }
            }
            $fieldId = $this.GetProfileFieldId("Main", "Comments")
            $this.comments = $personPSObject.details."$fieldId"
        }
    }

    [string] GetId() {
        return $this.id
    }

    [string] GetName() {
        if ([string]::IsNullOrEmpty($this.middle)) {
            return $this.first + " " + $this.last
        }
        else {
            return $this.first + " " + $this.middle + " " + $this.last
        }
    }

    [string] GetFirstName() {
        return $this.first
    }

    [string] GetNickName() {
        return $this.nickname
    }

    [string] GetLastName() {
        return $this.last
    }

    [string] GetMiddleName() {
        return $this.middle
    }

    [string] GetDisplayName() {
        if ([string]::IsNullOrEmpty($this.nickname)) {
            return $this.first + " " + $this.last
        }
        else {
            return $this.nickname + " " + $this.last
        }        
    }

    [string] GetFirstPrimaryEmail() {
        # Breeze actually stores multiple emails in the field separated by commas.
        # Only grab the first one.
        if ([string]::IsNullOrEmpty($this.email)) {
            return $this.email
        }

        $e = $this.email.Split(",")[0]
        $e = $e.Split(";")[0]
        return $e.Trim()
    }

    [string] GetStreetAddress() {
        return $this.streetaddress
    }

    [string] GetCity() {
        return $this.city
    }

    [string] GetState() {
        return $this.state
    }

    [string] GetZip() {
        return $this.zip
    }

    [string] GetHomePhone() {
        return $this.homephone
    }

    [string] GetMobilePhone() {
        return $this.mobile
    }

    [string] GetWorkPhone() {
        return $this.work
    }

    [string] GetComments() {
        return $this.comments
    }

    [string] ToString() {
        return "{0}: {1}, {2} {3} {4}" -f $this.id, $this.last, $this.first, $this.middle, $this.email
    }

    [string[]] GetTagNames() {
        return $this.tagNames
    }

    [string] GetFieldsJSON() {
        # TODO: SYNC OTHER FIELDS
        # Docs for update: https://app.breezechms.com/api#update_person
        $fields = [PSCustomObject] @(
            [PSCustomObject]@{
                field_id   = $this.GetProfileFieldId("Contact", "Address")
                field_type = "address"
                response   = $true
                details    = [PSCustomObject]@{
                    street_address = [System.Web.HttpUtility]::UrlEncode($this.streetAddress)
                    city           = [System.Web.HttpUtility]::UrlEncode($this.city)
                    state          = [System.Web.HttpUtility]::UrlEncode($this.state)
                    zip            = $this.zip
                }
            }
            [PSCustomObject]@{
                field_id   = $this.GetProfileFieldId("Contact", "Email")
                field_type = "email"
                response   = $true
                details    = [PSCustomObject]@{
                    address = [System.Web.HttpUtility]::UrlEncode($this.email)
                }
            }
            [PSCustomObject]@{
                field_id   = $this.GetProfileFieldId("Contact", "Phone")
                field_type = "phone"
                response   = $true
                details    = [PSCustomObject]@{
                    phone_home = $this.homephone
                }
            }
            [PSCustomObject]@{
                field_id   = $this.GetProfileFieldId("Contact", "Phone")
                field_type = "phone"
                response   = $true
                details    = [PSCustomObject]@{
                    phone_work = $this.work
                }
            }
            [PSCustomObject]@{
                field_id   = $this.GetProfileFieldId("Contact", "Phone")
                field_type = "phone"
                response   = $true
                details    = [PSCustomObject]@{
                    phone_mobile = $this.mobile
                }
            }
            [PSCustomObject]@{
                field_id   = $this.GetProfileFieldId("Main", "Comments")
                field_type = "single_line"
                response   = [System.Web.HttpUtility]::UrlEncode($this.comments)
            }
        )
        return $fields | ConvertTo-Json
    }

    static [Person[]] ToPersons([string] $breezeProfileFields, [string] $breezeJsonPersons) {
        [PSObject] $personPSObjects = $breezeJsonPersons | ConvertFrom-Json
        $personList = New-Object 'System.Collections.Generic.List[Person]'
        foreach ($personPSObject in $personPSObjects) {
            $personList.Add([Person]::new($breezeProfileFields, $personPSObject))
        }
        return $personList.ToArray()
    }

    
    static [Person[]] GetDedupedPersons([Person[]] $persons) {
        <#
            Remove any Persons from the list with the same email.  Preserve the first, discard the rest.
        #>
        $hs = New-Object 'System.Collections.Generic.HashSet[string]'
        $newPersonsList = New-Object 'System.Collections.Generic.List[PSObject]'
        foreach ($person in $persons) {
            $pemail = $person.GetFirstPrimaryEmail()
            if ((-not [String]::IsNullOrEmpty($pemail) -and `
                    (-not $hs.contains($pemail)))) {
                $newPersonsList.Add($person)
                $hs.Add($pemail)
            }
        }
        return $newPersonsList.ToArray()
    }

    static [int] GetProfileFieldId([PSCustomObject] $profileFields, [string] $sectionName, [string] $fieldName) {
       
        [boolean] $foundSection = $false
        foreach ($section in $profileFields) {
            if ($section.name -eq $sectionName) {
                $foundSection = $true
                foreach ($field in $section.fields) {
                    if ($field.name -eq $fieldName) {
                        return $field.field_id
                    }
                }
            }
        }

        $msg = "Section: $sectionName does not exist."
        if ($foundSection) {
            $msg = "Field: $fieldName does not exist in section: $sectionName"
        }

        [Logger]::Write($msg)
        [Logger]::Write($profileFields)

        throw [System.ArgumentException] $msg
    }

    [int] GetProfileFieldId([string] $sectionName, [string] $fieldName) {
        return [Person]::GetProfileFieldId($this.ProfileFields, $sectionName, $fieldName)
    }

    static [boolean] ContactEquals([Person] $p, [PSObject] $c) {
        # Compare a Person to an Exchange Contact
        if ( $c -eq $null ) { return $false }
        if ( $p.GetName() -ne $c.Name ) { return $false }
        if ( $p.GetName() -ne $c.Identity ) { return $false }
        if ( $p.GetFirstName() -ne $c.FirstName ) { return $false }
        if ( $p.GetLastName() -ne $c.LastName ) { return $false }
        if ( $p.GetDisplayName() -ne $c.DisplayName ) { return $false }
        if ( $p.GetStreetAddress() -ne $c.StreetAddress ) { return $false }
        if ( $p.GetCity() -ne $c.City ) { return $false }
        if ( $p.GetState() -ne $c.StateOrProvince ) { return $false }
        if ( $p.GetZip() -ne $c.PostalCode ) { return $false }
        if ( $p.GetHomePhone() -ne $c.Phone ) { return $false }
        if ( $p.GetMobilePhone() -ne $c.MobilePhone ) { return $false }
        if ( $p.GetWorkPhone() -ne $c.Office ) { return $false }
        return $true
    }

    [boolean] ContentEquals([Person] $p) {
        $fieldsMatch = $this.id -eq $p.id -and `
            $this.first -eq $p.first -and `
            $this.nickname -eq $p.nickname -and `
            $this.last -eq $p.last -and `
            $this.middle -eq $p.middle -and `
            $this.email -eq $p.email -and `
            $this.homephone -eq $p.homephone -and `
            $this.mobile -eq $p.mobile -and `
            $this.work -eq $p.work -and `
            $this.city -eq $p.city -and `
            $this.state -eq $p.state -and `
            $this.zip -eq $p.zip -and `
            $this.comments -eq $p.comments

        if (-not $fieldsMatch) {
            return $false
        }
        return $true;
    }
    
    [int] HashCode() {
        [int] $hashCode = $this.id;
        $hashCode = [HashCodeUtility]::UAdd($hashCode, [HashCodeUtility]::GetDeterministicHashCode($this.first))
        $hashCode = [HashCodeUtility]::UAdd($hashCode, [HashCodeUtility]::GetDeterministicHashCode($this.last))
        $hashCode = if ($this.nickname -ne $null) { [HashCodeUtility]::UAdd($hashCode, [HashCodeUtility]::GetDeterministicHashCode($this.nickname)) }
        $hashCode = if ($this.middle -ne $null) { [HashCodeUtility]::UAdd($hashCode, [HashCodeUtility]::GetDeterministicHashCode($this.middle)) }
        $hashCode = if ($this.email -ne $null) { [HashCodeUtility]::UAdd($hashCode, [HashCodeUtility]::GetDeterministicHashCode($this.email)) }
        $hashCode = if ($this.homephone -ne $null) { [HashCodeUtility]::UAdd($hashCode, [HashCodeUtility]::GetDeterministicHashCode($this.homephone)) }
        $hashCode = if ($this.mobile -ne $null) { [HashCodeUtility]::UAdd($hashCode, [HashCodeUtility]::GetDeterministicHashCode($this.mobile)) }
        $hashCode = if ($this.work -ne $null) { [HashCodeUtility]::UAdd($hashCode, [HashCodeUtility]::GetDeterministicHashCode($this.work)) } 
        $hashCode = if ($this.streetaddress -ne $null) { [HashCodeUtility]::UAdd($hashCode, [HashCodeUtility]::GetDeterministicHashCode($this.streetaddress)) }
        $hashCode = if ($this.city -ne $null) { [HashCodeUtility]::UAdd($hashCode, [HashCodeUtility]::GetDeterministicHashCode($this.city)) }
        $hashCode = if ($this.state -ne $null) { [HashCodeUtility]::UAdd($hashCode, [HashCodeUtility]::GetDeterministicHashCode($this.state)) }
        $hashCode = if ($this.zip -ne $null) { [HashCodeUtility]::UAdd($hashCode, [HashCodeUtility]::GetDeterministicHashCode($this.zip)) }
        $hashCode = if ($this.comments -ne $null) { [HashCodeUtility]::UAdd($hashCode, [HashCodeUtility]::GetDeterministicHashCode($this.comments)) }
        return $hashCode
    }
}
