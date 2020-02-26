<#
.Synopsis
   Class Module representing a Mock Breeze class
.NOTES
    Author:  Chris D. Johnson
    Requires:  PowerShell V5
    
    Copyright (c) 2020 Chris D. Johnson All rights reserved.  
    Licensed under the Apache 2.0. See LICENSE file in the project root for full license information.  
#>

using module ..\Modules\Person\
using module ..\Modules\Tag\
using module ..\Modules\Breeze\

$here = Split-Path -Parent $MyInvocation.MyCommand.Path

class MockBreeze : Breeze {
    static $PROFILE_FIELDS_FILE = "$script:here\Breeze.Tests.ProfileFields.json"
    static $PERSONS_TAG_TEST_FILE = "$script:here\Breeze.Tests.PersonsFromTagTEST.json"
    static $PERSON0_FILE = "$script:here\Breeze.Tests.Person0.json"
    static $PERSON1_FILE = "$script:here\Breeze.Tests.Person1.json"
    static $PERSON2_FILE = "$script:here\Breeze.Tests.Person2.json"
    static $PERSON3_FILE = "$script:here\Breeze.Tests.Person3.json"
    static $PERSON4_FILE = "$script:here\Breeze.Tests.Person4.json"
    static $PERSON5_FILE = "$script:here\Breeze.Tests.Person5.json"
    static $TAGS_FILE = "$script:here\Breeze.Tests.Tags.json"
    static $TAGPERSONS_FILE = "$script:here\Breeze.Tests.TagPersons.json"

    static [string] GetRawPersonJSONFromIndex([int] $index) {
        $var ="PERSON$index" + "_FILE"
        return [IO.File]::ReadAllText([MockBreeze]::"$var")
    }
    static [int] GetRawPersonIdFromIndex([int] $index) {
        $personJSON = [MockBreeze]::GetRawPersonJSONFromIndex($index)
        return ($personJSON | ConvertFrom-Json).id
    }

    [string] GetProfileFieldsAsJSON() {
        return [IO.File]::ReadAllText([MockBreeze]::PROFILE_FIELDS_FILE)
    }


    [string] GetPersonByIdAsJSON([int] $personId) {
        For ($i=0; $i -le 5; $i++) {
            $var ="PERSON$i" + "_FILE"
            $personJSON = [IO.File]::ReadAllText([MockBreeze]::"$var")
            $personObject = $personJSON | ConvertFrom-Json
            if ($personObject.id -eq $personId) {
                return $personJSON
            }
        }
        return $null
    }

    [Person[]] GetPersonsByEmail([string] $email) {
        
        $personsJSON = [IO.File]::ReadAllText([MockBreeze]::PERSONS_TAG_TEST_FILE)
        [Person[]] $persons = [Person]::ToPersons($this.GetProfileFieldsAsJSON(), $personsJSON)
        
        $newPersonsList = New-Object 'System.Collections.Generic.List[Person]'
        foreach ($person in $persons) {
            if ($person.GetFirstPrimaryEmail() -eq $email) {
                $newPersonsList.Add($person)
            }
        }
        return $newPersonsList.ToArray()
    }

    hidden [string] GetTagsAsJSON() {
        $tagsJSON = [IO.File]::ReadAllText([MockBreeze]::TAGS_FILE)
        return $tagsJSON
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

    [string] GetPersonsFromTagIdAsJSON([int]$tagId) {
        # This is totally hard-coded to ONE tag
        $personsJSON = [IO.File]::ReadAllText([MockBreeze]::PERSONS_TAG_TEST_FILE)
        return $personsJSON
    }
}
