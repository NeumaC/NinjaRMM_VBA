Attribute VB_Name = "Ticketautomatismus"
Option Explicit

' ************************************************************************
'   ACHTUNG:
'   Um dieses Makro nutzen zu k?nnen, muss in den VBA-Extras/Verweise
'   der Verweis auf "Microsoft VBScript Regular Expressions 5.5" aktiviert sein.
' ************************************************************************

' -----------------------------
'   Globale Konstanten
' -----------------------------
Public Const FOLDER_ARCHIV As String = "Archiv"
Public Const FOLDER_TICKETS As String = "Tickets"

Public Const ALLOWED_SENDER As String = "it-support | Gföllner"

Public Const REGEX_SUBJECT_PATTERN As String = "^\[gfoellner-at\]\s*\(#(\d+)\)(.*)"
Public Const REGEX_TICKETNUMBER_ONLY As String = "^\[gfoellner-at\]\s*\(#(\d+)\)"
Public Const REGEX_TICKET_REPLACE As String = "TICKET \(Gföllner - ([A-Z]+)\)( / )"
Public Const REGEX_STATUS_CLOSED As String = "Status:\s?.*?\s?Geschlossen"

Public Const USE_API As Boolean = True


' ------------------------------------------------------------------------
'  Gemeinsame Hilfsfunktionen
' ------------------------------------------------------------------------

Private Function GetOrCreateFolder(ByVal parentFolder As Outlook.Folder, _
                                   ByVal folderName As String) As Outlook.Folder
    ' Hilfsfunktion: Sucht im parentFolder nach einem Unterordner "folderName"
    ' und erstellt diesen bei Bedarf.

    Dim subFolder As Outlook.Folder
    On Error Resume Next
    Set subFolder = parentFolder.Folders(folderName)
    On Error GoTo 0

    If subFolder Is Nothing Then
        Set subFolder = parentFolder.Folders.Add(folderName)
    End If

    Set GetOrCreateFolder = subFolder
End Function

' Neu: Hilfsfunktion zum rekursiven Durchsuchen des Archivordners
Private Function FindTicketFolderInArchiveRecursively(ByVal parentFolder As Outlook.Folder, _
                                                      ByVal ticketNumber As String) As Outlook.Folder
    ' Durchsucht parentFolder und s?mtliche Unterordner rekursiv nach einem Ordner,
    ' dessen Name mit ticketNumber beginnt.

    Dim child As Outlook.Folder
    For Each child In parentFolder.Folders
        ' Prüfung, ob Ordnername das gesuchte TicketNumber-Präfix hat
        If Left(child.name, Len(ticketNumber)) = ticketNumber Then
            Set FindTicketFolderInArchiveRecursively = child
            Exit Function
        Else
            ' Rekursiver Abstieg
            Dim found As Outlook.Folder
            Set found = FindTicketFolderInArchiveRecursively(child, ticketNumber)
            If Not found Is Nothing Then
                Set FindTicketFolderInArchiveRecursively = found
                Exit Function
            End If
        End If
    Next child
End Function


' Holt per API den aktuellen Ticketbetreff und benennt den Ordner falls nötig um
Private Sub UpdateTicketFolderNameFromApi(ByVal ticketFolder As Outlook.Folder, ByVal ticketId As Long)
    On Error GoTo ErrHandler

    Dim apiSubject As String
    apiSubject = NinjaAPICall.GetTicketSubjectByApi(ticketId)

    ' Sicherheitscheck: Nur umbenennen, wenn ein valider Betreff vorliegt
    If Len(Trim$(apiSubject)) > 0 Then
        Dim expectedName As String
        expectedName = "#" & CStr(ticketId) & " (" & apiSubject & ")"

        If ticketFolder.name <> expectedName Then
            ticketFolder.name = expectedName
        End If
    End If

ExitHere:
    Exit Sub
ErrHandler:
    Debug.Print "Fehler in UpdateTicketFolderNameFromApi: ", Err.Description
    Resume ExitHere
End Sub

' ------------------------------------------------------------------------
'   Verarbeitung einer eingehenden Ticket-Email (neuer Vorgang)
' ------------------------------------------------------------------------

Public Sub ProcessEmail(mailItem As Outlook.mailItem, tasksFolder As Outlook.Folder)
    Dim regex As RegExp
    Dim ticketRegex As RegExp
    Dim matchObj As match
    Dim ticketMatch As MatchCollection
    Dim folderName As String
    Dim ticketNumber As String
    Dim remainingText As String
    Dim newFolder As Outlook.Folder
    Dim existingFolder As Outlook.Folder
    Dim FolderExists As Boolean
    Dim archiveFolder As Outlook.Folder

    ' Initialisiere das RegExp-Objekt zur Erkennung des Betreffs
    Set regex = New RegExp
    regex.Pattern = REGEX_SUBJECT_PATTERN
    regex.IgnoreCase = True
    regex.Global = False

    ' Regex zum Ersetzen im Rest-Text (TICKET (Gföllner - XX) / ...)
    Set ticketRegex = New RegExp
    ticketRegex.Pattern = REGEX_TICKET_REPLACE
    ticketRegex.IgnoreCase = True
    ticketRegex.Global = False

    ' Nur E-Mails verarbeiten, die vom erlaubten Absender kommen
    If mailItem.SenderName = ALLOWED_SENDER Then
        Dim subjectText As String
        subjectText = mailItem.Subject

        If regex.test(subjectText) Then
            ' Hole das Match-Objekt
            Set matchObj = regex.Execute(subjectText)(0)

            ' Ticketnummer => #1234
            ticketNumber = "#" & matchObj.SubMatches(0)
            ' Restlicher Betreff
            remainingText = matchObj.SubMatches(1)

            ' Ersetze optional "TICKET (Gföllner - XX) /" durch "XX"
            ' wenn es vorhanden ist.
            If ticketRegex.test(remainingText) Then
                Set ticketMatch = ticketRegex.Execute(remainingText)
                If ticketMatch.Count > 0 Then
                    remainingText = ticketMatch(0).SubMatches(0) & ticketMatch(0).SubMatches(1) & _
                                    Trim(Replace(remainingText, ticketMatch(0), ""))
                End If
            End If

            ' Ordnername zusammensetzen
            folderName = ticketNumber & " (" & Trim(remainingText) & ")"

            ' Prüfen, ob ein Ordner mit derselben Ticketnummer bereits im tasksFolder existiert
            FolderExists = False
            For Each existingFolder In tasksFolder.Folders
                If Left(existingFolder.name, Len(ticketNumber)) = ticketNumber Then
                    Set newFolder = existingFolder
                    FolderExists = True
                    Exit For
                End If
            Next

            ' Falls kein Ordner existiert, prüfe, ob er im Archiv (rekursiv) zu finden ist
            If Not FolderExists Then
                On Error Resume Next
                Set archiveFolder = tasksFolder.Folders(FOLDER_ARCHIV)
                On Error GoTo 0

                If Not archiveFolder Is Nothing Then
                    ' Neu: Rekursive Suche
                    Set newFolder = FindTicketFolderInArchiveRecursively(archiveFolder, ticketNumber)
                    If Not newFolder Is Nothing Then
                        newFolder.MoveTo tasksFolder
                        FolderExists = True
                    End If
                End If
            End If

            ' Wenn immer noch kein Ordner vorhanden, neu erstellen
            If Not FolderExists Then
                Set newFolder = GetOrCreateFolder(tasksFolder, folderName)
            End If

            ' E-Mail verschieben
            If Not newFolder Is Nothing Then
                mailItem.Move newFolder
                
                ' Ordnernamen anhand aktueller Ticketbeschreibung abgleichen
                Dim tid As Long
                tid = CLng(Mid$(ticketNumber, 2))
                UpdateTicketFolderNameFromApi newFolder, tid
            End If
        End If
    End If
End Sub

' ------------------------------------------------------------------------
'   Archivierung geschlossener Tickets
' ------------------------------------------------------------------------

Private Sub ArchiveTasksFolderAPI(ByVal tasksFolder As Outlook.Folder)
    Dim i As Long
    Dim ticketFolder As Outlook.Folder
    
    ' Gehen Sie durch alle (Ticket-)Unterordner in tasksFolder
    For i = tasksFolder.Folders.Count To 1 Step -1
        Set ticketFolder = tasksFolder.Folders(i)
        
        ' Optional: Ordner ignorieren, falls es der Archivordner selbst ist
        If LCase(ticketFolder.name) = LCase(FOLDER_ARCHIV) Then
            GoTo ContinueNext
        End If
        
        ' Holen wir uns die Ticketnummer aus dem Ordnernamen,
        ' z.B. Ordner heißt: "#3621 (Betreff/Resttext)"
        Dim ticketId As Long
        ticketId = ExtractTicketIdFromFolderName(ticketFolder.name)
        
        If ticketId > 0 Then
            ' Ordnerbezeichnung mit aktuellem Betreff abgleichen
            UpdateTicketFolderNameFromApi ticketFolder, ticketId
            
            ' Jetzt: Status via API ermitteln
            If NinjaAPICall.IsTicketClosedByApi(ticketId) Then
                ' Falls wir "geschlossen" haben, ermitteln wir zusätzlich das Abschlussdatum
                Dim closedDate As Date
                closedDate = GetTicketClosedDateByApi(ticketId)
                
                ' => Ticket ist geschlossen => verschieben ins Archiv
                Call MoveFolderToArchive(ticketFolder, tasksFolder, closedDate)
            End If
        End If

ContinueNext:
    Next i
End Sub

Private Sub ArchiveTasksFolder(ByVal tasksFolder As Outlook.Folder)
    ' Diese Prozedur durchsucht alle Ticket-Unterordner im übergebenen tasksFolder
    ' nach E-Mails, in deren Body die Statusänderung "Status: * ? Geschlossen" vorkommt.
    ' Wird eine solche E-Mail gefunden, wird der komplette Ticketordner
    ' in den Archivordner verschoben.
    '
    ' Anstatt das aktuelle Datum zu nehmen, verwenden wir das Empfangsdatum
    ' (ReceivedTime) der E-Mail mit dem abschließenden Status.
    '
    ' Archivstruktur:
    '   Archiv -> <Jahr> -> <Monat> -> [Ticket-Folder]

    Dim ns As Outlook.NameSpace
    Dim archivRoot As Outlook.Folder

    Set ns = Application.GetNamespace("MAPI")

    ' 1) Archiv-Ordner innerhalb des Tasks-Ordners suchen oder erstellen
    On Error Resume Next
    Set archivRoot = GetOrCreateFolder(tasksFolder, FOLDER_ARCHIV)

    Dim i As Long
    Dim ticketFolder As Outlook.Folder

    ' 2) Durch alle Unterordner in tasksFolder laufen (Ticketordner) - rückwärts
    For i = tasksFolder.Folders.Count To 1 Step -1
        Set ticketFolder = tasksFolder.Folders.item(i)

        ' "Archiv"-Ordner selbst überspringen
        If ticketFolder.name <> archivRoot.name Then
            Dim j As Long
            Dim mailItem As Outlook.mailItem
            Dim foundStatusChange As Boolean
            foundStatusChange = False

            ' Durchsuche alle Items im Ticketordner
            For j = ticketFolder.items.Count To 1 Step -1
                If ticketFolder.items(j).Class = olMail Then
                    Set mailItem = ticketFolder.items(j)

                    ' Prüfen, ob im Body "Status: ... ? Geschlossen" gefunden wird
                    If IsStatusClosed(mailItem.Body) Then
                        foundStatusChange = True

                        ' Wir holen uns das Empfangsdatum der "geschlossenen" E-Mail
                        Dim closedMailDate As Date
                        closedMailDate = mailItem.ReceivedTime  ' Empfangenes Datum

                        ' Falls gewünscht, die Ticketnummer zur Doku holen
                        Dim extractedTicketNumber As String
                        extractedTicketNumber = ExtractTicketNumber(mailItem.Subject)
                        Debug.Print "Gefundene Ticketnummer: " & extractedTicketNumber & _
                                    "; Datum: " & closedMailDate

                        ' 3) Zielordner anlegen (Archiv -> Jahr -> Monat), basierend auf closedMailDate
                        Dim yearFolder As Outlook.Folder
                        Dim monthFolder As Outlook.Folder
                        Dim yearString As String
                        Dim monthString As String

                        yearString = Format(closedMailDate, "yyyy")
                        monthString = Format(closedMailDate, "mm - mmmm")  ' "02 - Februar"

                        Set yearFolder = GetOrCreateFolder(archivRoot, yearString)
                        Set monthFolder = GetOrCreateFolder(yearFolder, monthString)

                        ' 4) Verschieben des Ticketordners in den erstellten Archivpfad
                        ticketFolder.MoveTo monthFolder

                        Exit For  ' Schleife abbrechen
                    End If
                End If
            Next j

            If foundStatusChange Then
                GoTo NextFolder
            End If
        End If
NextFolder:
    Next i
End Sub

Private Function ExtractTicketNumber(ByVal subjectText As String) As String
    ' Diese Funktion extrahiert die Ticketnummer (z.B. "#1234")
    ' aus einem Betreff, der folgendes Muster hat:
    '   [gfoellner-at] (#1234)
    '   optional gefolgt von weiterem Text.

    Dim subjectRegex As RegExp
    Set subjectRegex = New RegExp

    subjectRegex.Pattern = REGEX_TICKETNUMBER_ONLY
    subjectRegex.IgnoreCase = True
    subjectRegex.Global = False

    If subjectRegex.test(subjectText) Then
        ExtractTicketNumber = "#" & subjectRegex.Execute(subjectText)(0).SubMatches(0)
    Else
        ExtractTicketNumber = ""
    End If
End Function

Private Function IsStatusClosed(ByVal bodyText As String) As Boolean
    ' Testet per RegEx auf "Status: .* ? Geschlossen".
    Dim re As RegExp
    Set re = New RegExp

    re.Pattern = REGEX_STATUS_CLOSED
    re.IgnoreCase = True
    re.Global = False

    IsStatusClosed = re.test(bodyText)
End Function

' Hilfsfunktion, um aus Ordnernamen "#3621 (blabla)" die 3621 zu ziehen
' Sie können auch Ihren vorhandenen Regex-Ansatz wiederverwenden.
Private Function ExtractTicketIdFromFolderName(ByVal folderName As String) As Long
    On Error Resume Next
    ' Beispiel: suchen nach Muster: "#1234"
    Dim parts() As String
    parts = Split(folderName, " ")
    ' parts(0) wäre z.B. "#3621"
    If Left(parts(0), 1) = "#" Then
        ExtractTicketIdFromFolderName = CLng(Mid$(parts(0), 2))
    End If
End Function


' Beispielfunktion, um Ordner ins Archiv zu verschieben
' und zwar in die Struktur: Archiv -> <Jahr> -> <Monat> -> [Ticket-Folder]
Private Sub MoveFolderToArchive(ByVal ticketFolder As Outlook.Folder, ByVal tasksFolder As Outlook.Folder, Optional ByVal closedDate As Date = 0)
    Dim archiveFolder As Outlook.Folder
    Set archiveFolder = GetOrCreateFolder(tasksFolder, FOLDER_ARCHIV)

    ' Falls kein close-Datum ermittelt wurde, nehmen wir das heutige Datum
    If closedDate = 0 Then closedDate = Date

    Dim yearFolder As Outlook.Folder
    Dim monthFolder As Outlook.Folder

    Dim strYear As String
    Dim strMonth As String

    strYear = CStr(Year(closedDate))
    strMonth = Format(closedDate, "mm") & "-" & Format(closedDate, "mmmm")

    Set yearFolder = GetOrCreateFolder(archiveFolder, strYear)
    Set monthFolder = GetOrCreateFolder(yearFolder, strMonth)

    ticketFolder.MoveTo monthFolder
    Debug.Print "Ordner '" & ticketFolder.name & "' wurde archiviert in " & strYear & "\\" & strMonth
End Sub


' ------------------------------------------------------------------------
'   Startprozeduren
' ------------------------------------------------------------------------

Public Sub RunEmailRule()
    ' Diese Prozedur wird aufgerufen, um neu eingehende E-Mails
    ' zu verarbeiten und sie in die entsprechenden Ticketordner zu verschieben.

    Dim inbox As Outlook.Folder
    Dim tasksFolder As Outlook.Folder
    Dim items As Outlook.items
    Dim mailItem As Outlook.mailItem
    Dim ns As Outlook.NameSpace
    Dim i As Long
    Dim item As Object
    
    ' Posteingang holen
    Set ns = Application.GetNamespace("MAPI")
    Set inbox = ns.GetDefaultFolder(olFolderInbox)

    ' "Tickets"-Ordner unterhalb des Posteingangs
    On Error Resume Next
    Set tasksFolder = inbox.Folders(FOLDER_TICKETS)
    On Error GoTo 0
    If tasksFolder Is Nothing Then
        MsgBox "Der '" & FOLDER_TICKETS & "'-Ordner existiert nicht unter dem Posteingang.", vbExclamation
        Exit Sub
    End If

    ' Hole alle Elemente im Posteingang
    Set items = inbox.items
    
    ' Durchlaufe die E-Mails im Posteingang rückwärts
    For i = items.Count To 1 Step -1
        Set item = items(i)
        If TypeOf item Is Outlook.mailItem Then
            Set mailItem = item
            ' E-Mail mit ProcessEmail verarbeiten
            ProcessEmail mailItem, tasksFolder
        End If
    Next i
End Sub

Public Sub RunArchiveRule()
    ' Diese Prozedur wird aufgerufen, um geschlossene Tickets
    ' automatisch in die Archivstruktur zu verschieben.

    Dim inbox As Outlook.Folder
    Dim tasksFolder As Outlook.Folder
    Dim ns As Outlook.NameSpace

    ' Namespace und Posteingang
    Set ns = Application.GetNamespace("MAPI")
    Set inbox = ns.GetDefaultFolder(olFolderInbox)

    ' "Tickets"-Ordner unterhalb des Posteingangs
    On Error Resume Next
    Set tasksFolder = inbox.Folders(FOLDER_TICKETS)
    On Error GoTo 0

    If tasksFolder Is Nothing Then
        MsgBox "Der '" & FOLDER_TICKETS & "'-Ordner existiert nicht unter dem Posteingang.", vbExclamation
        Exit Sub
    End If

    ' Archivierung ausführen
    If USE_API Then
        ArchiveTasksFolderAPI tasksFolder
    Else
        ArchiveTasksFolder tasksFolder
    End If

    MsgBox "Archivierung abgeschlossen.", vbInformation
End Sub


