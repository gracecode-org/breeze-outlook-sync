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
    [switch]$test = $false,
    [switch]$loginonly = $false
 )

 $here = Split-Path -Parent $MyInvocation.MyCommand.Path

function ShowHelp {
    "To get started, run the following command in a PowerShell terminal:"
    ".\SyncContacts.ps1 -init"


}

function Pause() {
    Start-Sleep -s 3.5
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
} else {
    "Synchronizing tags: " + $TagsToSync
}

"Skipping tags: " + $TagsToSkip

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

# There are three authentication options supported in order:
# Certificate File and Password (a pfx file), used primarily for linux 
# Certificate Thumbprint, used for Windows.
# Username and Password (This is no longer valid and has been removed.)

$Exchange = [Exchange]::new(
    $Config.cfg.Exchange.Connection.URI, 
    $Config.cfg.Exchange.groupEmailDomain,
    $BreezeCache,
    $force
    )

if($Config.cfg.Exchange.Connection.psobject.properties.name -contains 'certThumbprint') {
    $Exchange.ConnectWithCertificateThumbprint($Config.cfg.Exchange.Connection.appId, $Config.cfg.Exchange.Connection.organization, 
        $Config.cfg.Exchange.Connection.certThumbprint)
} else {
    $Exchange.ConnectWithCertificateFile($Config.cfg.Exchange.Connection.appId, $Config.cfg.Exchange.Connection.organization, 
        $Config.cfg.Exchange.Connection.certFilePath, $Config.cfg.Exchange.Connection.certFilePassword)
}

if($test) {
    "  Connected!"
}

if($loginonly) {
    [Logger]::Write("Logging in and keeping session open.", $true)
}
else {
    $maxTagPersons = $Config.cfg.Breeze.SyncSettings.MaxTagPersons
    try {
        if(-not $test) {
            [Logger]::Write("Starting Sync")

            "Getting tags from Breeze..."
            [Tag[]] $Tags = $BreezeAPI.GetTags()
            [Logger]::Write("Retrieved " + $Tags.Length + " tags.", $true)


            foreach ($tag in $Tags) {
                if($null -ne $TagsToSkip -and $TagsToSkip.Contains($tag.GetName())) {
                    [Logger]::Write("Skipping tag (in skip list): " + $tag.GetName(), $true)
                } elseif($null -eq $TagsToSync -or $TagsToSync.Contains($tag.GetName())) {
                    [Logger]::Write("Fetching tag: " + $tag.GetName(), $true)
                    try {
                        # Retreive all the Persons from the tag, ignoring those who are invalid (no emails, no first/last name...)
                        [Person[]] $persons = $BreezeAPI.GetPersonsFromTagId($tag.GetId(), $true, $true)
                        
                        # if the number of persons exceeds the configured setting, then abort
                        if($null -eq $maxTagPersons -or ($persons.Length -le  $maxTagPersons)) {
                            $Tag.SetPersons($persons)
                            $Exchange.SyncDistributionGroupFromTag($tag)
                        } else {
                            [Logger]::Write("Skipping tag: " + $tag.GetName() + ".  MaxTagPersons exceeded: " + $persons.Length, $true)
                        }
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
}