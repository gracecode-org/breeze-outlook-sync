
- [Overview](#overview)
- [How Tags and People are synchronized](#how-tags-and-people-are-synchronized)
  - [What fields are synchronized?](#what-fields-are-synchronized)
- [Installation](#installation)
- [Requirements and Details](#requirements-and-details)
- [References](#references)

# Overview
This utility performs a one-way syncrhonization with Breeze CHMS Tags and Persons with Microsoft Office 365 / Exchange Distribution Groups and Mail Contacts.  

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
1.  Download the latest release from the releases page.
2.  Extract to your Program Files directory (e.g. `C:\Program Files\BreezeOutlookSync`)
3.  Open a PowerShell window

# Requirements and Details
This utility has been tested with:
- Windows 10 and 
- Windows Server 2016
- PowerShell 5.1.18362.628
- Tests require [Pester](https://github.com/pester/Pester)


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