<#
.Synopsis
   Breeze to Outlook Synchronizer
.DESCRIPTION
   This script performs a one-way synchronization of all Tags in Breeze
   to Exchange Distribution Lists and Contacts.

   Try to simulate a Tag -> Person relationship 

   Uses a cache directory to save costly calls to Outlook/Exchange.
   
   Retrieve all Breeze Tags with Persons.
   For each Tag:
     For each person in the tag:
       If the person has changed sync
       If the number of persons in a tag has changed OR any persons have changed, sync the entire group.

.EXAMPLE
   powershell -File SyncContacts.ps1 
.EXAMPLE
   powershell -File SyncContacts.ps1 -configfile .\config.json -force
.NOTES
    Author:  Chris D. Johnson
    Requires:  PowerShell V5

    Copyright (c) 2020 Chris D. Johnson All rights reserved.  
    Licensed under the Apache 2.0. See LICENSE file in the project root for full license information.  
#>

using module .\Modules\Person\
using module .\Modules\Breeze\
using module .\Modules\BreezeCache\
using module .\Modules\Tag\
using module .\Modules\Exchange\
using module .\Modules\Config\
using module .\Modules\Logger\

param (
    [string]$configfile = $env:APPDATA + "\BreezeOutlookSync\config.json",
    [switch]$help = $false,
    [switch]$force = $false,
    [switch]$init = $false,
    [switch]$test = $false
 )

 $here = Split-Path -Parent $MyInvocation.MyCommand.Path

function ShowHelp {
    "To get started, run the following command in a PowerShell terminal:"
    ".\SyncContacts.ps1 -init"


}

if($help) {
    ShowHelp
    exit 0
}

if($init) {
    "Creating configuration file template... "
    $dataPath = $env:APPDATA + "\BreezeOutlookSync"


    if (-not (Test-Path -PathType Container $dataPath)) {
        New-Item -ItemType Directory -Path $dataPath
    }
    Copy-Item -Path ($here + "\" + "examples\config.json") -Destination $configfile

    "Template configuration file created: " + $configfile
    "Next steps:"
    " 1. Edit the file"
    " 2. Test the file by running: SyncContacts.ps1 -test"
    exit 0
}

if (-not (Test-Path -Path $configfile -PathType Leaf)) {
    "Config file is missing or inaccessible: " + $ConfigFile
    return 1
}


"Config file: " + $configfile
$Config = [Config]::new($configfile)

[Logger]::init($Config.GetLogPath(), $Config.cfg.LogLevel, $Config.cfg.MaxLogSize, $Config.cfg.MaxLogFiles)
"Logging to:  " + ([Logger]::Logger.LogFile)
"MaxLogSize:  " + ([Logger]::Logger.MaxLogSize)
"MaxLogFiles: " + ([Logger]::Logger.MaxLogFiles)

if($force) {
    [Logger]::Write("Using force option to ignore local cache.", $true)
}

$TagsToSync = $null
if($null -ne $Config.cfg.Breeze.SyncSettings.TagsToSync) {
    $TagsToSync = [System.Collections.Generic.HashSet[string]] $Config.cfg.Breeze.SyncSettings.TagsToSync
    if($TagsToSync.Count -eq 0) {
        $TagsToSync = $null
    }
} 

$TagsToSkip = $null
if($null -ne $Config.cfg.Breeze.SyncSettings.TagsToSkip) {
    $TagsToSkip = [System.Collections.Generic.HashSet[string]] $Config.cfg.Breeze.SyncSettings.TagsToSkip
    if($TagsToSkip.Count -eq 0) {
        $TagsToSkip = $null
    }
} 


if ($null -eq $TagsToSync) {
    "Synchronizing all tags."
}

if ($null -eq $TagsToSkip) {
    "Skipping tags: " + $TagsToSkip
}

# Initialize our cache
$BreezeCache = [BreezeCache]::new($Config.GetCachePath())
"Caching to: " + $Config.GetCachePath()

# Initialize our connections
if($test) {
    "Trying to establish a connection to Breeze..."
}

$BreezeAPI = [Breeze]::new(
    $Config.cfg.Breeze.Connection.URI, 
    $Config.cfg.Breeze.Connection.APIKey)
        
if($test) {
    $fields = $BreezeAPI.GetProfileFields()
    "  Connected!"
    "Trying to establish a connection to Exchange..."
}

$Exchange = [Exchange]::new(
    $Config.cfg.Exchange.Connection.URI, 
    $Config.cfg.Exchange.Connection.username, 
    $Config.cfg.Exchange.Connection.password, 
    $Config.cfg.Exchange.groupEmailDomain,
    $BreezeCache,
    $force)

if($test) {
    "  Connected!"
}

try {
    if(-not $test) {
        [Logger]::Write("Starting Sync")

        "Getting tags from Breeze..."
        [Tag[]] $Tags = $BreezeAPI.GetTags()
        [Logger]::Write("Retrieved " + $Tags.Length + " tags.", $true)


        foreach ($tag in $Tags) {
            if($null -eq $TagsToSkip -or $TagsToSkip.Contains($tag.GetName())) {
                [Logger]::Write("Skipping tag (in skip list): " + $tag.GetName(), $true)
            } elseif($null -eq $TagsToSync -or $TagsToSync.Contains($tag.GetName())) {
                [Logger]::Write("Fetching tag: " + $tag.GetName(), $true)
                [Person[]] $persons = $BreezeAPI.GetPersonsFromTagId($tag.GetId(), $true)
                $Tag.SetPersons($persons)
                try {
                    $Exchange.SyncDistributionGrupFromTag($tag)
                } catch {
                    [Logger]::Write("Caught an exception... going to the next tag...", $true)
                    [Logger]::Write($_.Exception)
                    [Logger]::Write($_.ScriptStackTrace)
                }
            }    
        }
    }
} finally {
    $Exchange.Disconnect()
    if(-not $test) {
        [Logger]::Write("Sync Complete", $true)
    }
}