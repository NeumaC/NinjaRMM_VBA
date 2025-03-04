# VBA Outlook Ticket Archiving

This VBA project automates the management and archiving of tickets in Microsoft Outlook.

## Overview
- **ProcessEmail**: Moves new incoming emails based on a ticket number in the subject to appropriate ticket folders.
- **RunEmailRule**: Entry point for processing new incoming emails.
- **RunArchiveRule**: Entry point for archiving existing ticket folders once a ticket has been closed via a status email.
- **ArchiveTasksFolder**: Core routine for archiving. Looks for emails with "Status: … → Closed" in existing ticket folders and moves the corresponding ticket folder to the archive.

## Functions
- **ExtractTicketNumber**: Extracts the ticket number (e.g., `#1234`) from the email subject.
- **GetOrCreateFolder**: Searches for or creates an Outlook subfolder.
- **FindTicketFolderInArchiveRecursively**: Recursively searches the archive structure for a folder whose name matches the requested ticket number.
- **IsStatusClosed**: Uses a regular expression to check if the ticket body indicates "Closed."

## How It Works
1. **RunEmailRule** is executed:
   - Reads emails in the Inbox and calls `ProcessEmail` for each.
   - `ProcessEmail` checks if the email is from an allowed sender, extracts the ticket number, and moves the email to the correct ticket folder. If the folder already exists in the archive, it is moved back to the working folder.

2. **RunArchiveRule** is executed:
   - Looks in the already-created ticket folders for closed tickets (identified by an email with `Status: … → Closed`). If found, that folder is moved into an archive structure based on the received date (`yyyy/mm - mmmm`).

## Global Constants
- `FOLDER_ARCHIV`: Name of the archive folder.
- `FOLDER_TICKETS`: Name of the folder for active tickets.
- `ALLOWED_SENDER`: Allowed sender email address for ticket-related emails.
- `REGEX_SUBJECT_PATTERN`: Regex pattern used to detect the ticket number in the subject.
- `REGEX_TICKET_REPLACE`: Regex pattern for optional subject cleanup.
- `REGEX_TICKETNUMBER_ONLY`: Regex pattern to extract only the ticket number from the subject.

## Requirements
- Outlook 2019 (or higher) VBA editor.
- **Microsoft VBScript Regular Expressions 5.5** reference enabled (in Tools > References).
- Sufficient permissions to access the relevant Outlook folders.

## Installation
1. Copy all the VBA code into an Outlook VBA module.
2. Enable `Microsoft VBScript Regular Expressions 5.5` in **VBA References**.
3. Adjust the constants (`FOLDER_ARCHIV`, `FOLDER_TICKETS`, `ALLOWED_SENDER`, etc.) as needed for your environment.

## Usage
- **RunEmailRule**: Called manually or via a rule to automatically sort newly received emails.
- **RunArchiveRule**: Also can be called manually or by a timer (e.g., daily) to archive old or closed tickets.

## Notes
- The code searches email bodies for `Status: … → Closed`. Adjust this pattern for your ticket system or desired status text.
- The archive is organized by year/month based on the received date (`ReceivedTime`) of the "close" email.
- For extremely large mailboxes, you may wish to consider a professional archiving solution or an Outlook add-in.

