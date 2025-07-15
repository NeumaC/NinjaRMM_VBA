Attribute VB_Name = "NinjaAPICall"
' NinjaAPICall v1.0.0
' @author NeumaC
' https://github.com/NeumaC/NinjaRMM_VBA
'
' NinjaRMM Ticket Monitoring API calls
' Docs: https://app.ninjarmm.com/apidocs/?links.active=core#/ticketing/getTicketById

Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

' Provide the Ninja client ID through these constants.
' Leave these constants empty to be prompted for the values during runtime.
Private Const cNINJACLIENTID As String = ""

' WebClient instance used for making API calls to Ninja.
Private pNinjaClient As WebClient

' Ninja client ID and client secret values used for authentication.
Private pNinjaClientId As String

' --------------------------------------------- '
' Private Properties and Methods
' --------------------------------------------- '

''
' Retrieves the Ninja API client ID.
' If the client ID is not provided through the 'cNINJACLIENTID' constant, the user is prompted to enter the client ID.
'
' @property NinjaClientId
' @type {String}
' @return {String} The Ninja API client ID.
''
Private Property Get NinjaClientId() As String
    If pNinjaClientId = "" Then
        If cNINJACLIENTID <> "" Then
            pNinjaClientId = cNINJACLIENTID
        Else
            Dim InpBxResponse As String
            InpBxResponse = InputBox("Please Enter Ninja API Client ID", "NinjaRMM Ticket Connector - Microsoft Outlook")
            If InpBxResponse <> "" Then
                pNinjaClientId = InpBxResponse
            Else
                Err.Raise 11041 + vbObjectError, "NinjaAPICall.ClientIdInputBox", "User did not provide Ninja API Client ID"
            End If
        End If
    End If
    
    NinjaClientId = pNinjaClientId
End Property

''
' Initializes and returns a WebClient instance configured for making API calls to Ninja.
'
' @property NinjaClient
' @type {WebClient}
' @return {WebClient} The configured WebClient instance.
'
' The WebClient instance is set up with the following configurations:
' - Base URL set to 'https://gfoellner.ninjarmm.eu/v2/'
' - Authenticator set to an instance of the 'NinjaAuthenticator' class, which handles Ninja's OAuth2 authentication flow.
' - The 'offline_access' scope is requested during the authentication process.
'
' The WebClient instance is cached and reused between requests.
''
Private Property Get NinjaClient() As WebClient
    If pNinjaClient Is Nothing Then
        ' Create a new WebClient instance with the base URL
        Set pNinjaClient = New WebClient
        pNinjaClient.BaseUrl = "https://gfoellner.rmmservice.eu/v2/"
        
        ' Set up the 'NinjaAuthenticator' instance for OAuth2 authentication
        Dim Auth As NinjaAuthenticator
        Set Auth = New NinjaAuthenticator
        Auth.Setup CStr(NinjaClientId)
        
        ' Request the 'offline_access' and 'accounting.reports.read' scopes
        Auth.AddScope "offline_access"
        Auth.AddScope "monitoring"
        
        ' Set the 'NinjaAuthenticator' instance as the authenticator for the WebClient
        Set pNinjaClient.Authenticator = Auth
    End If
    
    Set NinjaClient = pNinjaClient
End Property

''
' Sets the WebClient instance used for making API calls to Ninja.
'
' @property NinjaClient
' @type {WebClient}
' @param {WebClient} Client - The WebClient instance to set.
''
Private Property Set NinjaClient(client As WebClient)
    Set pNinjaClient = client
End Property

' --------------------------------------------- '
' Execution
' --------------------------------------------- '

''
' Calls the login procedures for the user interface button.
'
' @method Login_Click
'
' This function performs the following steps:
' 1. Enables logging.
' 2. Retrieves the pre-set authenticator object from the NinjaClient.
' 3. Logs out and clears the cache for the current session.
' 4. Initiates the login process.
' 5. Returns the authenticator reference to the NinjaClient.
' 6. Handles any errors that occur during the process and logs them.
''
Public Sub Login_Click()
    On Error GoTo ApiCall_Cleanup
    ' Enable logging
    WebHelpers.EnableLogging = True
    
    ' Retrieve the pre-set authenticator object
    Dim Auth As NinjaAuthenticator
    Set Auth = NinjaClient.Authenticator
    Set NinjaClient.Authenticator = Nothing
    
    ' Logout and clear cache for the current session
    Auth.Logout
    
    ' Login
    Auth.Login
    
    ' Return the authenticator reference to the NinjaClient
    Set NinjaClient.Authenticator = Auth
    ' Clear the local reference to the authenticator
    Set Auth = Nothing
    
ApiCall_Cleanup:
    ' Error handling block
    If Err.Number <> 0 Then
        ' Clean up if an error happened
        pNinjaClientId = ""
        Set NinjaClient = Nothing
        ' Construct the error description message
        Dim auth_ErrorDescription As String
        
        auth_ErrorDescription = "An error occurred during the login process." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
        
        ' Log the error
        WebHelpers.LogError auth_ErrorDescription, "NinjaAPICall.Login_Click", 11041 + vbObjectError
        ' Notify the user of the error
        MsgBox "ERROR:" & vbNewLine & vbNewLine & auth_ErrorDescription, vbCritical + vbOKOnly, "NinjaRMM Ticket Connector - Microsoft Outlook"
    End If
End Sub

''
' Clears all saved tokens for the user interface button.
'
' @method ClearCache_Click
'
' This function performs the following steps:
' 1. Enables logging.
' 2. Confirms the user's action to clear the cache.
' 3. If the user confirms, retrieves the pre-set authenticator object.
' 4. Clears all cache (tenants and tokens) and logs out of the current session.
' 5. Returns the authenticator reference to the NinjaClient.
' 6. Handles any errors that occur during the process and logs them.
''
Public Sub ClearCache_Click()
    On Error GoTo ApiCall_Cleanup
    ' Enable logging
    WebHelpers.EnableLogging = True
    
    ' Confirm user action
    Dim msgBoxResponse As VbMsgBoxResult
    msgBoxResponse = MsgBox("This action will clear saved tokens (access). You will be required to log in for the next request." & _
        vbNewLine & vbNewLine & "Proceed to clears cache?", vbExclamation + vbYesNo, "NinjaRMM Ticket Connector - Microsoft Outlook")
    
    Select Case msgBoxResponse
        Case vbYes
            ' Retrieve the pre-set authenticator object
            Dim Auth As NinjaAuthenticator
            Set Auth = NinjaClient.Authenticator
            ' Clear the reference to the authenticator in the NinjaClient
            Set NinjaClient.Authenticator = Nothing
            
            ' Clear all cache (tokens)
            Auth.ClearAllCache isClearToken:=True
            
            ' Clear current session tokens cache by logging out
            Auth.Logout
            
            ' Return the authenticator reference to the NinjaClient
            Set NinjaClient.Authenticator = Auth
            ' Clear the local reference to the authenticator
            Set Auth = Nothing
            
        Case vbNo
            ' Exit the subroutine if the user cancels the action
            Exit Sub
    End Select

ApiCall_Cleanup:
    ' Error handling block
    If Err.Number <> 0 Then
        ' Clean up if an error occurred
        pNinjaClientId = ""
        Set NinjaClient = Nothing
        ' Construct the error description message
        Dim auth_ErrorDescription As String
        
        auth_ErrorDescription = "An error occurred while clearing cache." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
    
        ' Log the error
        WebHelpers.LogError auth_ErrorDescription, "NinjaAPICall.ClearCache_Click", 11041 + vbObjectError
        ' Notify the user of the error
        MsgBox "ERROR:" & vbNewLine & vbNewLine & auth_ErrorDescription, vbCritical + vbOKOnly, "NinjaRMM Ticket Connector - Microsoft Outlook"
    End If
End Sub

' ------------------------------------------------------------------------
'  API-basierte Statusabfrage
' ------------------------------------------------------------------------

Public Function IsTicketClosedByApi(ByVal ticketId As Long) As Boolean
    On Error GoTo ErrHandler
    
    ' Beispielhaft: Wir verwenden den bereits vorhandenen WebClient aus NinjaAPICall,
    '               der z.B. als "NinjaClient" bereitsteht.
    '               Passen Sie das ggf. an Ihre Struktur an!

    Dim client As WebClient
    Set client = NinjaClient ' aus NinjaAPICall.bas oder ähnlich

    ' Neues Request-Objekt erstellen
    Dim req As New WebRequest
    ' Ressource zusammensetzen (Beispiel-Endpunkt):
    req.Resource = "ticketing/ticket/" & CStr(ticketId)
    ' GET-Methode
    req.Method = WebMethod.HttpGet
    req.ResponseFormat = WebFormat.Json  ' Wir erwarten JSON-Antwort
    
    Dim resp As WebResponse
    Set resp = client.Execute(req)
    
    If resp.StatusCode = 200 Then
        ' JSON-Antwort parsen, resp.Data ist i.d.R. ein Dictionary
        ' Wir erwarten:
        '  {
        '    "id":3621,
        '    "status": {
        '       "name":"CLOSED",
        '       "statusId":6000
        '    },
        '    ...
        '  }

        ' Aus resp.Data("status") das Sub-Dictionary holen
        Dim statusDict As Dictionary
        Set statusDict = resp.Data("status")
        
        ' statusId prüfen
        Dim sid As Long
        sid = CLng(statusDict("statusId"))
        
        If sid = 6000 Then
            IsTicketClosedByApi = True
        Else
            IsTicketClosedByApi = False
        End If
    Else
        ' Wenn z.B. Fehler 404 oder anderes - hier je nach Bedarf behandeln
        Debug.Print "API-Aufruf war nicht erfolgreich, Status: " & resp.StatusCode
        IsTicketClosedByApi = False
    End If
    
    Exit Function
    
ErrHandler:
    Debug.Print "Fehler in IsTicketClosedByApi:", Err.Description
    IsTicketClosedByApi = False
End Function

' Ermittelt aus der Log-Historie den Zeitstempel, wann das Ticket durch Automation ID=1000 geschlossen wurde
' Liefert 0, falls kein Eintrag gefunden.
Public Function GetTicketClosedDateByApi(ticketId As Long) As Date
    On Error GoTo ErrHandler

    Dim client As WebClient
    Set client = NinjaClient ' ODER anpassen, wo Ihr WebClient herkommt

    Dim req As New WebRequest
    req.Resource = "ticketing/ticket/" & CStr(ticketId) & "/log-entry?type=SAVE"
    req.Method = WebMethod.HttpGet
    req.ResponseFormat = WebFormat.Json

    Dim resp As WebResponse
    Set resp = client.Execute(req)

    If resp.StatusCode = 200 Then
        ' Wir erwarten ein Array von Log-Einträgen
        Dim arrLogs As Collection
        Set arrLogs = resp.Data ' i.d.R. ein Collection-Objekt

        Dim i As Long
        For i = 1 To arrLogs.Count
            Dim logItem As Dictionary
            Set logItem = arrLogs.item(i)

            ' Prüfen, ob automation vorhanden
            If logItem.Exists("automation") Then
                Dim autom As Dictionary
                Set autom = logItem("automation")

                ' Falls automation.id = 1000 => das ist unser finaler close-Eintrag
                If autom.Exists("id") Then
                    If CLng(autom("id")) = 1000 Then
                        ' Dann Zeitstempel aus createTime übernehmen
                        ' Wert direkt mit Val() umwandeln – unabhängig vom Dezimaltrennzeichen
                        Dim dblTime As Double
                        dblTime = Val(CStr(logItem("createTime")))

                        ' Gemäß Beispiel: 1744206453.867411000 => sek seit 1.1.1970
                        ' -> In VBA-Datum umrechnen:
                        '   1 Tag = 86400 sek
                        Dim epoch As Date
                        epoch = #1/1/1970#

                        GetTicketClosedDateByApi = epoch + (dblTime / 86400)
                        Exit Function
                    End If
                End If
            End If
        Next i
    End If

    ' Falls nicht gefunden oder kein Erfolg:
    GetTicketClosedDateByApi = 0

    Exit Function

ErrHandler:
    Debug.Print "Fehler in GetTicketClosedDateByApi: ", Err.Number, Err.Description
    GetTicketClosedDateByApi = 0
End Function


' Liefert den aktuellen Betreff eines Tickets
Public Function GetTicketSubjectByApi(ticketId As Long) As String
    On Error GoTo ErrHandler

    Dim client As WebClient
    Set client = NinjaClient

    Dim req As New WebRequest
    req.Resource = "ticketing/ticket/" & CStr(ticketId)
    req.Method = WebMethod.HttpGet
    req.ResponseFormat = WebFormat.Json

    Dim resp As WebResponse
    Set resp = client.Execute(req)

    If resp.StatusCode = 200 Then
        Dim jsonText As String
        jsonText = WebHelpers.BytesToUtf8String(resp.body)
        
        Dim ticketDict As Dictionary
        Set ticketDict = WebHelpers.ParseJson(jsonText)

        If ticketDict.Exists("subject") Then
            GetTicketSubjectByApi = CStr(ticketDict("subject"))
        End If
    Else
        GetTicketSubjectByApi = ""
    End If

    Exit Function

ErrHandler:
    Debug.Print "Fehler in GetTicketSubjectByApi:", Err.Description
    GetTicketSubjectByApi = ""
End Function

' ------------------------------------------------------------------------
'   Helper: Find Ninja user UID by email
' ------------------------------------------------------------------------
Private Function GetUserUidByEmail(ByVal userEmail As String, ByVal clientId As Long) As String
    On Error GoTo ErrHandler

    Dim req As New WebRequest
    req.Resource = "ticketing/app-user-contact"
    req.Method = WebMethod.HttpGet
    req.ResponseFormat = WebFormat.Json

    req.AddQuerystringParam "clientId", CStr(clientId)
    req.AddQuerystringParam "searchCriteria", userEmail
    req.AddQuerystringParam "pageSize", "50"

    Dim resp As WebResponse
    Set resp = NinjaClient.Execute(req)

    If resp.StatusCode = 200 Then
        Dim users As Collection
        Set users = resp.Data
        Dim i As Long
        For i = 1 To users.Count
            Dim userRec As Dictionary
            Set userRec = users.item(i)
            If LCase$(userRec("email")) = LCase$(userEmail) Then
                GetUserUidByEmail = CStr(userRec("uid"))
                Exit Function
            End If
        Next i
    End If

    GetUserUidByEmail = ""
    Exit Function

ErrHandler:
    Debug.Print "Fehler in GetUserUidByEmail:", Err.Description
    GetUserUidByEmail = ""
End Function

' ------------------------------------------------------------------------
'   Helper: convert attachment to Base64 string
' ------------------------------------------------------------------------
Private Function AttachmentToBase64(att As Outlook.Attachment) As String
    On Error GoTo ErrHandler

    Dim tmpPath As String
    tmpPath = Environ$("TEMP") & "\" & att.fileName
    att.SaveAsFile tmpPath

    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 'adTypeBinary
    stream.Open
    stream.LoadFromFile tmpPath
    Dim bytes() As Byte
    bytes = stream.Read
    stream.Close
    Kill tmpPath

    Dim xml As Object
    Dim node As Object
    Set xml = CreateObject("MSXML2.DOMDocument")
    Set node = xml.createElement("b64")
    node.DataType = "bin.base64"
    node.nodeTypedValue = bytes
    AttachmentToBase64 = node.Text

    Exit Function

ErrHandler:
    Debug.Print "Fehler in AttachmentToBase64:", Err.Description
    AttachmentToBase64 = ""
End Function

' ------------------------------------------------------------------------
'   Helper: upload attachments via comment endpoint
' ------------------------------------------------------------------------
Private Sub UploadEmailAttachmentsToTicket(ByVal ticketId As Long, mailItem As Outlook.mailItem)
    On Error GoTo ErrHandler

    If mailItem.Attachments.Count = 0 Then Exit Sub

    Dim files() As Variant
    ReDim files(0 To mailItem.Attachments.Count - 1)
    Dim i As Long
    For i = 1 To mailItem.Attachments.Count
        files(i - 1) = AttachmentToBase64(mailItem.Attachments(i))
    Next i

    Dim comment As New Dictionary
    comment.Add "public", True
    comment.Add "body", ""
    comment.Add "files", files

    Dim req As New WebRequest
    req.Resource = "ticketing/ticket/" & CStr(ticketId) & "/comment"
    req.Method = WebMethod.HttpPost
    req.RequestFormat = WebFormat.Json
    req.ResponseFormat = WebFormat.Json
    Set req.body = comment

    Dim resp As WebResponse
    Set resp = NinjaClient.Execute(req)
    ' ignore response

    Exit Sub

ErrHandler:
    Debug.Print "Fehler in UploadEmailAttachmentsToTicket:", Err.Description
End Sub

' ------------------------------------------------------------------------
'   Create a new ticket from an Outlook mail item
' ------------------------------------------------------------------------
Public Function CreateTicketFromEmail(mailItem As Outlook.mailItem, _
                                      ByVal clientId As Long, _
                                      ByVal ticketFormId As Long) As Long
    On Error GoTo ErrHandler

    Dim requester As String
    requester = GetUserUidByEmail(mailItem.SenderEmailAddress, clientId)
    If requester = "" Then Err.Raise vbObjectError + 513, , "Requester not found"

    Dim body As New Dictionary
    body.Add "clientId", clientId
    body.Add "ticketFormId", ticketFormId
    body.Add "subject", mailItem.Subject

    Dim desc As New Dictionary
    desc.Add "public", True
    desc.Add "body", mailItem.body
    desc.Add "htmlBody", mailItem.HTMLBody
    body.Add "description", desc
    body.Add "requesterUid", requester

    Dim req As New WebRequest
    req.Resource = "ticketing/ticket"
    req.Method = WebMethod.HttpPost
    req.RequestFormat = WebFormat.Json
    req.ResponseFormat = WebFormat.Json
    Set req.body = body

    Dim resp As WebResponse
    Set resp = NinjaClient.Execute(req)

    If resp.StatusCode >= 200 And resp.StatusCode < 300 Then
        Dim ticketInfo As Dictionary
        Set ticketInfo = resp.Data
        CreateTicketFromEmail = CLng(ticketInfo("id"))
        UploadEmailAttachmentsToTicket CreateTicketFromEmail, mailItem
    Else
        CreateTicketFromEmail = -1
    End If

    Exit Function

ErrHandler:
    Debug.Print "Fehler in CreateTicketFromEmail:", Err.Description
    CreateTicketFromEmail = -1
End Function

