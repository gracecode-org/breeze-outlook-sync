<#
.Synopsis
   Class Module representing a cache for Breeze Tags and Persons
.NOTES
    Author:  Chris D. Johnson
    Requires:  PowerShell V5
    
    Copyright (c) 2020 Chris D. Johnson All rights reserved.  
    Licensed under the Apache 2.0. See LICENSE file in the project root for full license information.  
#>


using module ..\Person\
using module ..\Tag\


class BreezeCache {
    $CacheDir

    BreezeCache() {
        
        $this.CacheDir = $env:APPDATA + "\BreezeOutlookSync\cache\"
        $this.CreateCacheDir()
    }

    BreezeCache([string] $cacheDir) {
        $this.CacheDir = $cacheDir
        $this.CreateCacheDir()
    }


    # Cache Structure
    # /cache/tag_<id> (name, personids)
    # /cache/person_<id> (hashcode)

    Clear() {
        Remove-Item $this.CacheDir -Recurse -Force
        New-Item -ItemType Directory -Path $this.CacheDir
    }

    hidden CreateCacheDir() {
        if ((Test-Path -PathType Container $this.CacheDir) -eq $false) {
            New-Item -ItemType Directory -Path $this.CacheDir
        }
    }


    #Return true if the tag or persons in the tag have changed
    [boolean] HasTagChanged([Tag] $tag) {
        $TagRefFileName = $this.CacheDir + "\tag_" + $tag.id + ".json"
        $TagRef = Get-Content -Path $TagRefFileName | Out-String | ConvertFrom-Json

        if ($TagRef -eq $Null) {
            return $true
        }

        if($TagRef.Name -ne $tag.name) {
            return $true
        }

        if($TagRef.PersonIds.length -ne $tag.GetPersons().length) {
            return $true
        }

        foreach($person in $tag.GetPersons()) {
            if($this.HasPersonChanged($person)) {
                return $true
            }
        }
        return $false
    }

    # Cache the tag and persons referenced by it.
    [void] CacheTag([Tag] $tag) {
        $TagRefFileName = $this.CacheDir + "\tag_" + $tag.id + ".json"
        $persons = $tag.GetPersons()
        
        [int[]] $personIds = [int[]]::new($persons.length)
        for($i=0;$i -lt $persons.length;$i++) {
            $personIds[$i] = $persons[$i].GetId()
            $this.CachePerson($persons[$i])
        }

        $TagRef = [PSCustomObject]@{
            Name = $tag.name
            PersonIds = $personIds
        }

        $TagRefJson = $TagRef | ConvertTo-Json
        Set-Content -Path $TagRefFileName -Value $TagRefJson

    }

    # Check ther person against our local cache.
    [Boolean] HasPersonChanged([Person] $person) {
        $PersonRefFileName = $this.CacheDir + "\person_" + $person.id + ".json"
        $PersonRef = Get-Content -Path $PersonRefFileName | Out-String | ConvertFrom-Json

        if($PersonRef -eq $Null) {
            return $true
        }

        $hash = $person.HashCode()

        return ($PersonRef.Hash -ne $hash)
    }

    [void] CachePerson([Person] $person) {
        $PersonRefFileName = $this.CacheDir + "\person_" + $person.id + ".json"

        $PersonRef = [PSCustomObject]@{
            Hash = $person.HashCode()
            #For diagnostic purposes
            DisplayName = $person.GetDisplayName()
            FirstPrimaryEmail = $person.GetFirstPrimaryEmail()
        }

        $PersonRefJson = $PersonRef | ConvertTo-Json
        Set-Content -Path $PersonRefFileName -Value $PersonRefJson
    }

}