<#
.Synopsis
   Class Module for interacting with the Breeze CHMS APIs.
.NOTES
    Author:  Chris D. Johnson
    Requires:  PowerShell V5
    
    Copyright (c) 2020 Chris D. Johnson All rights reserved.  
    Licensed under the Apache 2.0. See LICENSE file in the project root for full license information.  
#>

#requires -version 5

using module ..\Logger\
using module ..\Person\
using module ..\Tag\
using module ..\Config\

$here = Split-Path -Parent $MyInvocation.MyCommand.Path

class Breeze {
    <#
        Limitations:
            - Assumes all tag names are unique
            - Names can't currently have URL values in them, like & or ?
    #>
    $BreezeApiURL
    $BreezeApiKey
    $Headers 

    $TagList = $null
    $ProfileFieldsJSON = $null
    $ProfileFields = $null

    $ENDPOINTS

    Breeze() {
        
    }

    Breeze([string] $apiUri, [string] $apiKey) {
        $this.BreezeApiKey = $apiKey
        $this.BreezeApiURL = $apiUri
        $this.ENDPOINTS = [PSCustomObject] @{
            PEOPLE        = $this.BreezeApiURL + '/api/people';
            EVENTS        = $this.BreezeApiURL + '/api/events';
            PROFILEFIELDS = $this.BreezeApiURL + '/api/profile';
            CONTRIBUTIONS = $this.BreezeApiURL + '/api/giving';
            FUNDS         = $this.BreezeApiURL + '/api/funds';
            PLEDGES       = $this.BreezeApiURL + '/api/pledges';
            TAGS          = $this.BreezeApiURL + '/api/tags';
        }

        $this.Headers = @{
            'Accept'       = 'application/json';
            'Content-Type' = 'application/json';
            'Api-Key'      = $this.BreezeApiKey;
        }
    }

    [Person] GetPersonById([int] $personId) {
        $personJSON = $this.GetPersonByIdAsJSON($personId)
        return [Person]::new($this.GetProfileFieldsAsJSON(), $personJSON)
    }

    [string] GetPersonByIdAsJSON([int] $personId) {
        $endpoint = $this.ENDPOINTS.PEOPLE + "/$personId"
        $response = Invoke-WebRequest -Uri $endpoint  -Method Get -Headers $this.Headers
        return $response.Content
    }
    
    [Person[]] GetPersonsFromTagId([int]$tagId, [boolean] $dedupe) {
        $peopleJsonList = $this.GetPersonsFromTagIdAsJSON($tagId)
        if ([string]::IsNullOrEmpty($peopleJsonList)) { throw [System.ArgumentException]::new('Unable to retrieve people for tag: $tagId') }
        [Person[]] $persons = [Person]::ToPersons($this.GetProfileFieldsAsJSON(), $peopleJsonList)
        if ($dedupe) { 
            # Trim any duplicate emails and persons without emails
            $persons = [Person]::GetDedupedPersons($persons)
            $mergedPersons = [System.Collections.ArrayList]::new()
            foreach ($person in $persons) {
                # Merge any duplicate persons into a single one.
                $mergedPersons.Add($this.GetMergedPersonsByEmail($person.GetFirstPrimaryEmail()))
            }
            $persons = $mergedPersons.ToArray()
        }
        return $persons
    }

    [string] GetPersonsFromTagIdAsJSON([int]$tagId) {
        $endpoint = $this.ENDPOINTS.PEOPLE + '?details=1&filter_json={"tag_contains":"y_' + [System.Web.HttpUtility]::UrlEncode($tagId) + '"}'
        $response = Invoke-WebRequest -Uri $endpoint  -Method Get -Headers $this.Headers
        return $response.Content
    }

    [Person] SyncPerson([Person] $person) {
        <#
            .DESCRIPTION
            Creates/Updates the Person in Breeze.
            If the person.id is not set or 0, it is assumed that the person is new.
            The person.id will be updated when created.

            Throws an exception if it fails.
        #>

        if ([string]::IsNullOrEmpty($person)) { throw [System.ArgumentException]::new('person cannot be null or empty') }

        [PSObject] $breezePerson = $null
        if ($person.id -eq 0) {
            $endpoint = $this.ENDPOINTS.PEOPLE + "/add?first=" + [System.Web.HttpUtility]::UrlEncode($person.GetFirstName()) + `
                "&last=" + [System.Web.HttpUtility]::UrlEncode($person.GetLastName()) + `
                "&middle=" + [System.Web.HttpUtility]::UrlEncode($person.GetMiddleName()) + `
                "&nick=" + [System.Web.HttpUtility]::UrlEncode($person.GetNickName())
            $breezePerson = Invoke-WebRequest -Uri $endpoint  -Method Post -Headers $this.Headers | ConvertFrom-Json
            if ($breezePerson -eq $null) { throw [System.ApplicationException]::new('Unable to create person: $person') }
            $person.id = $breezePerson.id
        }


        # Docs for update: https://app.breezechms.com/api#update_person
        $fieldsJson = $person.GetFieldsJSON()

        # Update the additional Fields
        $personId = $person.id
        $endpoint = $this.ENDPOINTS.PEOPLE + "/update?person_id=$personId&fields_json=$fieldsJson"  
        $breezePerson = Invoke-WebRequest -Uri $endpoint  -Method Post -Headers $this.Headers | ConvertFrom-Json
        if ($breezePerson -eq $null) { throw [System.ApplicationException]::new('Unable to update person: $person') }


        # Assign/re-assign the appropriate tags.
        foreach ($tagName in $person.getTagNames()) {
            [Tag] $tag = $this.GetTagByName($tagName)
            [int] $tagId = $tag.id
            $endpoint = $this.ENDPOINTS.TAGS + "/assign?person_id=$personId&tag_id=$tagId"  
            $result = Invoke-WebRequest -Uri $endpoint  -Method Put -Headers $this.Headers | ConvertFrom-Json
        }

        return $person
    }

    [Person[]] SyncPersons([Person[]] $persons) {
        foreach ($person in $persons) {
            $this.SyncPerson($person)
        }
        return $persons
    }

    [void] DeletePerson([int] $personId) {
        if ($personId -eq 0) { throw [System.ArgumentException]::new('person id must be non-zero') }
        $endpoint = $this.ENDPOINTS.PEOPLE + "/delete?person_id=$personId"  
        Invoke-WebRequest -Uri $endpoint  -Method Delete -Headers $this.Headers | ConvertFrom-Json
    }

    [void] DeletePerson([Person] $person) {
        if ([string]::IsNullOrEmpty($person)) { throw [System.ArgumentException]::new('person cannot be null or empty') }
        if ($person.id -eq 0) { throw [System.ArgumentException]::new('person.id must be non-zero') }
        $this.DeletePerson($person.id)
    }

    [void] DeletePersons([Person[]] $persons) {
        foreach ($person in $persons) {
            $this.DeletePerson($person)
        }
    }
    
    hidden [PSObject] GetTagsPSObject() {
        $endpoint = $this.ENDPOINTS.TAGS + '/list_tags'

        if ($this.TagList -eq $null) {
            $this.TagList = $this.GetTagsAsJSON() | ConvertFrom-Json
        }

        return $this.TagList
    }

    hidden [string] GetTagsAsJSON() {
        $endpoint = $this.ENDPOINTS.TAGS + '/list_tags'
        $response = Invoke-WebRequest -Uri $endpoint  -Method Get -Headers $this.Headers
        return $response.Content
    }

    
    [Tag[]] GetTags() {
        # This assumes that all tag names are unique.
        [PSObject] $tagsPSObject = $this.GetTagsPSObject()

        $tags = [System.Collections.ArrayList]::new()
        foreach ($tagPSObject in $tagsPSObject) {
            $tags.Add([Tag]::new($tagPSObject))
        }
        return $tags
    }

    [Tag] GetTagByName([string] $tagName) {
        # This assumes that all tag names are unique.
        [PSObject] $tagsPSObject = $this.GetTagsPSObject()

        foreach ($tagPSObject in $tagsPSObject) {
            if ($tagPSObject.name -eq $tagName) {
                return [Tag]::new($tagPSObject)
            }
        }
        return $null
    }

    [Tag] AddTag([Tag] $tag) {
        <#
            .DESCRIPTION
            Create a new tag in the root folder.
        #>
        if ([string]::IsNullOrEmpty($tag)) { throw [System.ArgumentException]::new('tag cannot be null or empty') }
        $tagName = $tag.name
        $endpoint = $this.ENDPOINTS.TAGS + "/add_tag?name=$tagName"

        # Invalidate the taglist cache
        $this.TagList = $null
        $tagPSObject = Invoke-WebRequest -Uri $endpoint  -Method Get -Headers $this.Headers | ConvertFrom-Json
        $tag.id = $tagPSObject.id
        return $tag
    }

    [Tag[]] AddTags([Tag[]] $tags) {
        foreach ($tag in $tags) {
            $this.AddTag($tag)
        }
        return $tags
    }

    [void] DeleteTag([int] $tagId) {
        if ($tagId -eq 0) { throw [System.ArgumentException]::new('tag id must be non-zero') }
        $endpoint = $this.ENDPOINTS.TAGS + "/delete_tag?tag_id=$tagId"  
        $result = Invoke-WebRequest -Uri $endpoint  -Method Delete -Headers $this.Headers | ConvertFrom-Json
    }

    [void] DeleteTag([Tag] $tag) {
        if ([string]::IsNullOrEmpty($tag)) { throw [System.ArgumentException]::new('tag cannot be null or empty') }
        if ($tag.id -eq 0) { throw [System.ArgumentException]::new('person.id must be non-zero') }
        $this.DeleteTag($tag.id)
    }

    [Tag[]] DeleteTags([Tag[]] $tags) {
        foreach ($tag in $tags) {
            $this.DeleteTag($tag)
        }
        return $tags
    }

    [string] GetProfileFieldsAsJSON() {
        $endpoint = $this.ENDPOINTS.PROFILEFIELDS

        if ($this.ProfileFieldsJSON -eq $null) {
            $response = Invoke-WebRequest -Uri $endpoint  -Method Get -Headers $this.Headers
            $this.ProfileFieldsJSON = $response.Content
        }
        return $this.ProfileFieldsJSON
    }

    [PSObject] GetProfileFields() {
        if ($this.ProfileFields -eq $null) {
            $this.ProfileFields = ConvertFrom-Json($this.GetProfileFieldsAsJSON())
        }
        return $this.ProfileFields
    }


    [Person[]] GetPersonsByEmail([string] $email) {
        $emailFieldId = [Person]::GetProfileFieldId($this.GetProfileFields(), "Contact", "Email")
        $endpoint = $this.ENDPOINTS.PEOPLE + '?details=1&filter_json={"' + $emailFieldId + '":"' + $email + '"}'
        $response = Invoke-WebRequest -Uri $endpoint  -Method Get -Headers $this.Headers
        [Person[]] $persons = [Person]::ToPersons($this.GetProfileFieldsAsJSON(), $response.Content)
        return $persons
    }

    [Person] GetMergedPersonsByEmail([string] $email) {
        <#
            .DESCRIPTION
            Merge all people with the same email into one Person entity as follows:
            - Use the first person by sorted by id number as the template person.
            - Remove the Middle name and nickname
            - Modify the first name to have commas and AND:
                - Jane and Bob
                - Jane, Bob and Chris
            - Keep the first last name.  Not ideal, since families have multiple last names.
        #>
        [Person[]] $persons = $this.GetPersonsByEmail($email)

        if ($persons.length -eq 1) {
            return $persons[0]
        }
        else {
            $persons = ($persons | Sort-Object -property "id")
            $newPerson = $persons[0];
            $newPerson.middle = $null
            $newPerson.nickname = $null
            $newName = ""
            $count = 0
            foreach ($person in $persons) {
                $firstName = $person.GetFirstName()
                if ($count -eq 0) {
                    $newName = $firstName
                }
                elseif ($count + 1 -eq $persons.length) {
                    $newName += " and $firstName"
                }
                else {
                    $newName += ", $firstName"
                }
                $count++
            }
            $newPerson.first = $newName
            return $newPerson
        }
    }

    static [boolean] HasEmail([Person[]] $persons) {
        <#
            .DESCRIPTION
            Return true of any of the people in the list has an email address.
        #>
        foreach ($person in $persons) {
            $email = $person.GetFirstPrimaryEmail()
            if ($email -ne "") {
                return $true                
            }
        }
        return $false
    }
}
