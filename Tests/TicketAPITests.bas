Attribute VB_Name = "TicketAPITests"
Option Explicit

' ------------------------------------------------------------------------
'   Test macro: create a new ticket from the currently selected Outlook email
' ------------------------------------------------------------------------
Public Sub TestCreateTicketFromEmail()
    Dim sel As Outlook.Selection
    Dim ids As Collection
    Dim item As Variant
    Dim report As String

    Set sel = Application.ActiveExplorer.Selection
    If sel Is Nothing Or sel.Count = 0 Then
        MsgBox "Please select one or more emails for the test first.", vbExclamation
        Exit Sub
    End If

    ' TODO: replace 1 with your real ticketFormId
    Set ids = NinjaAPICall.CreateTicketsFromSelection(1)

    For Each item In ids
        If CLng(item) <> -1 Then
            report = report & "Ticket ID #" & item & " created" & vbCrLf
        Else
            report = report & "Ticket creation failed" & vbCrLf
        End If
    Next item

    MsgBox report, vbInformation
End Sub
