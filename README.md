# VBA Outlook Ticket Management with NinjaRMM Integration

This VBA project automates the handling and archiving of tickets in Microsoft Outlook.  
It also integrates with the NinjaRMM API to determine if a ticket is closed based on its status from the remote system, rather than parsing status text from emails.

## Overview

- **ProcessEmail**: Moves newly received emails (with ticket numbers in the subject) to the correct ticket folders.
- **RunEmailRule**: Entry point for processing new incoming emails (calls `ProcessEmail`).
- **RunArchiveRule**: Entry point for archiving existing ticket folders once a ticket has been confirmed *closed* by querying the NinjaRMM API (via `IsTicketClosedByApi`).
- **ArchiveTasksFolder**: Core routine for archiving ticket folders. It checks the NinjaRMM API to see if a ticket has `status.statusId = 6000` (i.e. closed). If so, it moves the ticket folder into a date-based archive structure.

## Key Functions

### In `Ticketautomatismus.bas`
- **ProcessEmail**  
  - Checks if the email is from an allowed sender, extracts the ticket number from the subject, and moves the email to the corresponding ticket folder.  
  - If the folder is found in the archive, it’s moved back to the active area.

- **ArchiveTasksFolder**  
  - Iterates through existing ticket folders to see if the ticket is closed by calling `NinjaAPICall.IsTicketClosedByApi(ticketId)`.  
  - If closed, obtains the closing timestamp with `GetTicketClosedDateByApi`, and moves the folder into the archive structure organized by `<year>/<month>`.

- **ExtractTicketIdFromFolderName**  
  - A helper that parses a folder name (e.g. `#3621 (Subject Matter)`) to extract the numeric ticket ID.

- **MoveFolderToArchive**  
  - Moves the existing ticket folder into the archive subfolders: `Archiv -> <Year> -> <Month>`.  
  - If no closed date is found, it uses the current date.

### In `NinjaAPICall.bas`
- **IsTicketClosedByApi**  
  - Performs an HTTP GET request to the NinjaRMM endpoint (`/ticketing/ticket/<ticketId>`) to retrieve the ticket status.  
  - Considers `statusId = 6000` as “closed.”

- **GetTicketClosedDateByApi**  
  - Retrieves the log history (`/ticketing/ticket/<ticketId>/log-entry?type=SAVE`) and looks for an automation entry with `id = 1000` (e.g. “Close Tickets Trigger”).  
  - Extracts the Unix timestamp from the `createTime` field and converts it to a VBA `Date`.

- **Login_Click**, **GenerateReport_Click**, **ClearCache_Click**  
  - Example macros for user interaction with the Ninja API (managing authentication flows, generating sample reports, clearing cached tokens, etc.).

## New Features (v1.1.0 - in development)

- **ProgressForm**: Displays progress updates for lengthy operations such as archiving many tickets.

## How It Works

1. **RunEmailRule**  
   - Loops through emails in the Outlook Inbox and calls `ProcessEmail` for each.  
   - `ProcessEmail` checks if the email is from the allowed sender, locates or creates a ticket folder, and moves the email into that folder.

2. **RunArchiveRule**  
   - Iterates through existing ticket folders (under the “Tickets” folder).  
   - For each folder, it calls `IsTicketClosedByApi` (from `NinjaAPICall.bas`). If the ticket is closed, `GetTicketClosedDateByApi` obtains the final closing timestamp from the log history.  
   - The folder is then placed into the archive folder structure (by year/month) based on that timestamp.

## Global Constants (in Ticketautomatismus.bas)

- `FOLDER_ARCHIV`: Name of the root archive folder (e.g. `"Archiv"`).
- `FOLDER_TICKETS`: Name of the folder for active tickets (e.g. `"Tickets"`).
- `ALLOWED_SENDER`: Allowed sender email address for incoming ticket emails.
- `REGEX_SUBJECT_PATTERN`: Regex pattern used to detect the ticket number in the subject (e.g. `^\[.*\]\s*\(#(\d+)\)`).
- `REGEX_TICKET_REPLACE`: Regex pattern for optional subject cleanup.

## Requirements

1. **Microsoft Outlook 2019 (or higher)** with VBA support.  
2. **Microsoft VBScript Regular Expressions 5.5** reference enabled (in the VBA editor under Tools > References).  
3. Access to NinjaRMM API credentials (client ID, etc.) if using the integrated `IsTicketClosedByApi`.

## Installation

1. Open Outlook’s VBA editor (e.g., press `Alt+F11`).
2. Import or copy the `.bas` and `.cls` modules into the Outlook VBA Project.
3. Enable the **Microsoft VBScript Regular Expressions 5.5** reference.
4. In the code, update constants (`FOLDER_ARCHIV`, `FOLDER_TICKETS`, `ALLOWED_SENDER`) to match your environment.
5. Provide your NinjaRMM API client details if needed (see `NinjaAPICall.bas`).
6. (Optional) **Ribbon Buttons**:  
   - Open Outlook > File > Options > Customize Ribbon.  
   - Add macros (e.g. `RunEmailRule`, `RunArchiveRule`) as new buttons.

## Usage

- **RunEmailRule**: Manually invoked or attached to an Outlook rule to process new incoming emails automatically.
- **RunArchiveRule**: Manually invoked or scheduled (e.g., daily) to move closed tickets into the archive subfolders.

## Notes

- The project now relies on **API-based** detection for closed tickets (`statusId = 6000`) rather than searching email text for “Closed.”  
- For archiving, the date used is retrieved from the log entry with `automation.id = 1000` (the “Close Tickets Trigger”).  
- The subfolder naming convention for the archive is `<Year>/<Month>` (e.g. `"2023/04-April"`).
- For large mailboxes or advanced workflows, consider a more robust add-in or third-party solution for professional archiving or help-desk integration.
