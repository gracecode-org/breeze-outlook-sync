<#
.Synopsis
   Class Module representing a Breeze Tag.
.NOTES
    Author:  Chris D. Johnson
    Requires:  PowerShell V5
    
    Copyright (c) 2020 Chris D. Johnson All rights reserved.  
    Licensed under the Apache 2.0. See LICENSE file in the project root for full license information.  
#>
using module ..\Person\

class Tag {
    <#
        # Tags have folders.  Add support later.

        Tag JSON Object:
        {
            "id": "1615875",
            "name": "Council Members",
            "created_on": "2019-02-02 21:56:55",
            "folder_id": "708569"
        }
    #>
    [ValidateNotNullOrEmpty()][int] $id
    [ValidateNotNullOrEmpty()][string] $name
    [Person[]] $persons;

    Tag([string] $name) {
        $this.name = $name
    }

    Tag([int] $id, [string] $name, [Person[]] $persons) {
        $this.id = $id
        $this.name = $name
        $this.persons = $persons
    }

    Tag([int] $id, [string] $name) {
        $this.id = $id
        $this.name = $name
    }


    Tag([PSObject] $tagPSObject) {
        $this.init($tagPSObject)
        $this.persons = [Person[]]::new(0)
    }

    Tag([PSObject] $tagPSObject, [Person[]] $persons) {
        $this.init($tagPSObject)
        $this.persons = $persons
    }
    [void] Init([PSObject] $tagPSObject) {
        $this.id = $tagPSObject.id
        $this.name = $tagPSObject.name
    }

    [int] GetId() {
        return $this.id
    }

    [string] GetName() {
        return $this.name
    }

    [void] SetPersons([Person[]] $persons) {
        $this.persons = $persons
    }

    [Person[]] GetPersons() {
        return $this.persons
    }

    [Person[]] AddPerson([Person] $person) {
       $this.persons += $person
       return $this.persons
    }

    [Person[]] RemovePerson([int] $personId) {
        $newPersons = [System.Collections.ArrayList]::new()
        for($i=0;$i -lt $this.persons.Length;$i++) {
            if($this.persons[$i].GetId() -ne $personId) {
                $newPersons.Add($this.persons[$i])
            }
        }
        $this.persons = $newPersons.ToArray()
        return $this.persons
    }
}