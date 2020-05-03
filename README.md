
- [Overview](#overview)
- [How Tags and People are synchronized](#how-tags-and-people-are-synchronized)
  - [What fields are synchronized?](#what-fields-are-synchronized)
- [Installation](#installation)
- [Getting Started](#getting-started)
- [Requirements and Details](#requirements-and-details)
- [config.json](#configjson)
- [References](#references)
- [Troubleshooting](#troubleshooting)
  - [WebCmdletIEDomNotSupportedException when testing the connection](#webcmdletiedomnotsupportedexception-when-testing-the-connection)
- [Support](#support)

# Overview
This utility performs a one-way syncrhonization with Breeze CHMS Tags and Persons with Microsoft Office 365 / Exchange Distribution Groups and Mail Contacts.  A Breeze Person becomes an Office/Exchange Contact and a Tag becomes a Distribution group.  

This allows organizations to use Breeze as the authoritative source of staff and congregation members, and still use Office 365 for for primary email.

This is written as a Microsoft Powershell script because Office 365 only provides Distribution Group management using the Exchange Administration console.

# How Tags and People are synchronized
Office 365 and Breeze have a different relationship model:
- Office 365 doesn't have folders.  Every distribution list group must be uniquely named.
  - When a tag with a duplicate name is found in another folder, a warning is logged and the tag is NOT synchronized until it's fixed in Breeze
- Office 365 doesn't allow two people to share an email address.
  - When two people are found with the same email address, they are merged into one as follows:
    1. Use the first person by sorted by id number as the template person.
    2. Remove the Middle name and nickname
    3. Modify the first name to have commas and AND:
       - Jane and Bob
       - Jane, Bob and Chris
    4. Keep the first last name.  Not ideal, since families have multiple last names.
- Office 365 doesn't allow a person to have more than one email.
  - Only the first email address in breeze will be used.
- Tags are synchronized with Distribution Lists by Display Name
  - If a Distribution List exists with the same name, only it's membership is updated.
    - Once a DL is created, notes can be updated and the email address can be changed and it won't be updated.
- Distribution Groups become orphaned if a tag is deleted or renamed.
  - Periodically look through Outlook/Exchange to remove any unneeded DL's

## What fields are synchronized?
- First name
- Nickname
- Middle name
- Last name
- Primary email 
- Street address
- City
- State
- Zip
- Home phone
- Work phone
- Mobile phone

# Installation
1.  Download the latest release from the [releases page](https://github.com/gracecode-org/breeze-outlook-sync/releases).
2.  Extract to your Program Files directory (e.g. `C:\Program Files\BreezeOutlookSync`)
3.  Open a PowerShell window and run the following commands:
    1.  `ps> cd "C:\Program Files\BreezeOutlookSync`
    2.  `ps> Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser`

        See https:/go.microsoft.com/fwlink/?LinkID=135170 for more information on Script Execution Policies.
    3.  `ps> .\SyncContacts.ps1 -init`

The following output is displayed:
```
Creating configuration file template...
Template configuration file created: C:\Users\Chris\AppData\Roaming\BreezeOutlookSync\config.json
Next steps:
1. Edit the file
2. Test the file by running: SyncContacts.ps1 -test
```  

# Getting Started
1.  Edit the file created in the installation step.  See the `config.json` section for details on how to edit this file.
2.  Test your connection, in your Powershell window:
    ```
    .\SyncContacts.ps1 -Test
    Config file: C:\Users\Chris\AppData\Roaming\BreezeOutlookSync\config.json
    Logging to:  C:\Users\Chris\AppData\Roaming\BreezeOutlookSync\logs\sync.log
    MaxLogSize:  20971520
    MaxLogFiles: 25
    Caching to: C:\Users\Chris\AppData\Roaming\BreezeOutlookSync\cache
    Trying to establish a connection to Breeze...
      Connected!
    Trying to establish a connection to Exchange...
    Connecting to Exchange...
      Connected!
    Disconnecting from Exchange...
    ```
3.  In Breeze, create a "Test Tag1" tag and add a "Test Person" person to the tag.
4.  Verify your `config.json` file is set to to ONLY synchronize the `Test Tag1` and `Test Tag2` tag
5.  Run the following command in your PowerShell window, which synchronize the test tag to Exchange without modifying your current distribution list:
    ```
    .\SyncContacts.ps1
    ```
6.  Once the sync is complete, look for the `Test Tag1` Distribution Group in Office 365 and verify that the Person in the tag has been created.  You can delete the Person and Distribution group when done.
7.  Once you are comfortable that the sync is working properly, you can add more tags to the 

Example output:
```
Config file: C:\Users\Chris\AppData\Roaming\BreezeOutlookSync\config.json
Logging to:  C:\Users\Chris\AppData\Roaming\BreezeOutlookSync\logs\sync.log
MaxLogSize:  20971520
MaxLogFiles: 25
Caching to: C:\Users\Chris\AppData\Roaming\BreezeOutlookSync\cache
Connecting to Exchange...
WARNING: Using New-PSSession with Basic Authentication is going to be deprecated soon, checkout https://aka.ms/exops-docs for using Exchange Online V2 Module which uses Modern Authentication.
Getting tags from Breeze...
Retrieved 112 tags.
Fetching tag: Chris Test
Synchronizing tag: Test Tag1(2482084) with 1 persons..
Creating new distribution group: Test Tag1
New! Office 365 Groups are the next generation of distribution lists.
Groups give teams shared tools for collaborating using email, files, a calendar, and more.
You can start right away using the New-UnifiedGroup cmdlet.
Setting distribution group email: TestTag1@mydomain.com
Synchronzing Person: 21660770: Test1, Chris  christest@yopmail.com
Disconnecting from Exchange...
Sync Complete
```

# Requirements and Details
This utility has been tested with:
- Windows 10 and 
- Windows Server 2016
- PowerShell 5.1.18362.628 or later
- Tests require [Pester](https://github.com/pester/Pester)

# config.json


# References

PowerShell
- [REST API Invocation](https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/invoke-restmethod?view=powershell-6)
- [Exchange Server PowerShell](https://docs.microsoft.com/en-us/powershell/exchange/exchange-server/exchange-management-shell?view=exchange-ps)
- [Import-CSV](https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/import-csv?view=powershell-6)
- [New-MailContact](https://docs.microsoft.com/en-us/powershell/module/exchange/users-and-groups/new-mailcontact?view=exchange-ps)
- https://powershellexplained.com/2017-04-07-all-dotnet-exception-list/
- [Pester Test Framework](https://github.com/pester/Pester)

Breeze:
- [API Docs](https://app.breezechms.com/api)


# Troubleshooting

## WebCmdletIEDomNotSupportedException when testing the connection
If you get this error:
```
Invoke-WebRequest : The response content cannot be parsed because the Internet Explorer engine is not available, or
Internet Explorer's first-launch configuration is not complete. Specify the UseBasicParsing parameter and try again.
At C:\Users\cdj06\Documents\breeze-outlook-sync-git\Modules\Breeze\Breeze.psm1:251 char:25
+ ... $response = Invoke-WebRequest -Uri $endpoint  -Method Get -Headers $t ...
+                 ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : NotImplemented: (:) [Invoke-WebRequest], NotSupportedException
    + FullyQualifiedErrorId : WebCmdletIEDomNotSupportedException,Microsoft.PowerShell.Commands.InvokeWebRequestComman
   d
```

1.  Windows Start / Run
2.  `iexplore`
3.  Complete the initial security question dialog.

# Support
This software is Apache 2.0 license and the source is therefore freely available.

To ask questions and open issues and download the latest release, see:
https://github.com/gracecode-org/breeze-outlook-sync/issues


