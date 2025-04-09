VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TenantStuff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' Saves obtained Xero tenants' information into an Excel sheet.
'
' @method StoreTenants
' @param {WebResponse} GetTenantResponse - The response received from the GetTenant procedure.
'
' This function performs the following steps:
' 1. Checks if the sheet for storing tenant information exists. If not, creates a new sheet and hides it.
' 2. Checks if the table (ListObject) for storing tenant information exists. If not, creates a new table with headers based on the response keys.
' 3. Adds a new row to the table for each tenant and fills it with the tenant data.
' 4. Limits the number of rows in the table to 50, deleting any excess rows from the bottom.
' 5. Logs the success of the operation and handles any errors that occur.
''
Private Sub StoreTenants(GetTenantResponse As WebResponse)
    On Error GoTo StoreTenants_Cleanup
    
    Dim wb As Workbook
    Dim sh As Worksheet
    Dim TenantTable As ListObject
    Dim NewRow As ListRow
    
    Set wb = ThisWorkbook
    
    ' Check if the sheet for storing tenant information exists
    If Not WebHelpers.WorksheetExists(auth_SheetTenantsStorage, wb) Then
        ' Create a new sheet and set its name
        Set sh = wb.Sheets.Add(Before:=wb.Sheets(1))
        sh.name = auth_SheetTenantsStorage
        
        ' Hide the sheet from the user
        sh.Visible = xlSheetVeryHidden
    Else
        ' Sheet exists, set the sheet variable
        Set sh = wb.Sheets(auth_SheetTenantsStorage)
    End If
    
    ' Check if the ListObject (table) for storing tenant information exists
    If Not WebHelpers.ListObjectExists(auth_TableTenantsStorage, sh) Then
        ' Create the table headers for storing tenant information
        ' Expected 'Data(1)' example from get tenants response:
        '   "id": "e1eede29-f875-4a5d-8470-17f6a29a88b1",
        '   "authEventId": "d99ecdfe-391d-43d2-b834-17636ba90e8d",
        '   "tenantId": "70784a63-d24b-46a9-a4db-0e70a274b056",
        '   "tenantType": "ORGANISATION",
        '   "tenantName": "Maple Florist",
        '   "createdDateUtc": "2019-07-09T23:40:30.1833130",
        '   "updatedDateUtc": "2020-05-15T01:35:13.8491980"
        '
        ' Based on official docs (10 May 2024):
        ' https://developer.xero.com/documentation/guides/oauth2/auth-flow/#5-check-the-tenants-youre-authorized-to-access
        '
        ' Column to record GetTenant response datetime
        sh.Cells(1, 1).Value = "Get_Datetime"
        ' Columns based on response Data Keys
        sh.Range(sh.Cells(1, 2), sh.Cells(1, GetTenantResponse.Data(1).Count + 1)) = GetTenantResponse.Data(1).Keys
        
        ' Set up the new table as an Excel table (ListObject)
        Set TenantTable = sh.ListObjects.Add(SourceType:=xlSrcRange, _
                                            Source:=sh.Range(sh.Cells(1, 1), sh.Cells(1, GetTenantResponse.Data(1).Count + 1)), _
                                            XlListObjectHasHeaders:=xlYes)
        TenantTable.name = auth_TableTenantsStorage
    Else
        ' Table exists, set the table variable
        Set TenantTable = sh.ListObjects(auth_TableTenantsStorage)
    End If
    
    ' Add a new row to the table for each tenant and fill it with tenant data
    Dim d As Dictionary
    For Each d In GetTenantResponse.Data
        Set NewRow = TenantTable.ListRows.Add(Position:=1, AlwaysInsert:=True)
        With NewRow.Range
            .Cells(1, 1).Value = Now
            ' DEBUG: Must call "Parent" first to avoid unexpected relative reference
            .Parent.Range(.Cells(1, 2), .Cells(1, d.Count + 1)) = d.items
        End With
    Next d
                         
    ' Log success of storing tenant information
    WebHelpers.LogDebug "Stored current tenants info success, sheet: " & auth_SheetTenantsStorage, "XeroAuthenticator.StoreTenants"
    
    ' Limit the number of rows in the TenantTable to 50, deleting any excess rows from the bottom
    LimitTableRows TenantTable, 50
    
StoreTenants_Cleanup:
    ' Cleanup object variables
    Set wb = Nothing
    Set sh = Nothing
    Set TenantTable = Nothing
    Set NewRow = Nothing
    
    ' Error handling block
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        
        ' Construct the error description message
        auth_ErrorDescription = "An error occurred while storing tenants info into Excel sheet." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
    
        ' Log the error
        WebHelpers.LogError auth_ErrorDescription, "XeroAuthenticator.StoreTenants", 11041 + vbObjectError
        ' Raise the error for further handling
        Err.Raise 11041 + vbObjectError, "XeroAuthenticator.StoreTenants", auth_ErrorDescription
    End If
End Sub

''
' Loads saved tenant information from an Excel sheet.
'
' @method LoadTenants
'
' This function performs the following steps:
' 1. Clears the existing dictionary of tenant IDs.
' 2. Checks if the worksheet containing tenant information exists. If not, logs a failure and exits the subroutine.
' 3. Checks if the ListObject (table) containing tenant information exists. If not, logs a failure and exits the subroutine.
' 4. Verifies if there are any rows in the table. If not, logs a failure and exits the subroutine.
' 5. Retrieves the most recent get date from the top row of the table for comparison.
' 6. Iterates through each row in the table, loading only the most recent tenant ID into the dictionary.
' 7. Logs the success of loading the tenants into the dictionary.
' 8. Handles any errors that occur during the process, logging the error details and raising an error.
''
Private Sub LoadTenants()
    On Error GoTo LoadTenants_Cleanup
    
    Dim wb As Workbook
    Dim sh As Worksheet
    Dim TenantTable As ListObject
    Dim NewestGetDate As Date
    
    ' Clear the 'DTenantsId' property
    If DTenantsId.Count > 0 Then
        DTenantsId.RemoveAll
    End If
    
    Set wb = ThisWorkbook
    ' Check if the sheet for loading tenant informations exists
    If Not WebHelpers.WorksheetExists(auth_SheetTenantsStorage, wb) Then
        ' Log failure due to non-existent sheet
        WebHelpers.LogDebug "Load tenants failed, sheet: " & auth_SheetTenantsStorage & "does not exist", "XeroAuthenticator.LoadTenants"
        ' Exit the subroutine
        Exit Sub
    Else
        ' Sheet exists, set the sheet variable
        Set sh = wb.Sheets(auth_SheetTenantsStorage)
    End If
    
    ' Check if the ListObject (table) for loading tenant informations exists
    If Not WebHelpers.ListObjectExists(auth_TableTenantsStorage, sh) Then
        ' Log failure due to non-existent table
        WebHelpers.LogDebug "Load tenants failed, ListObject(Table): " & auth_TableTenantsStorage & " does not exist", "XeroAuthenticator.LoadTenants"
        ' Exit the subroutine
        Exit Sub
    Else
        ' Table exists, set the table variable
        Set TenantTable = sh.ListObjects(auth_TableTenantsStorage)
    End If
    
    ' Check if there are rows in the ListObject (table)
    If Not TenantTable.ListRows.Count > 0 Then
        ' Log failure if there are no rows in the table
        WebHelpers.LogDebug "Load tenants failed, ListObject(Table) " & auth_TableTenantsStorage & " does not have any rows/records", "XeroAuthenticator.LoadTenants"
        ' Exit the subroutine
        Exit Sub
    Else
        ' The most recent record is at the top row (index = 1)
        ' Load the most recent get date for comparison later
        NewestGetDate = TenantTable.ListRows(1).Range.Cells(1, 1).Value
    End If
    
    ' Load tenants into the 'DTenantsId' property
    Dim TenantRow As ListRow
    For Each TenantRow In TenantTable.ListRows
        ' Only load the most recent tenant ID
        If CDate(TenantRow.Range.Cells(1, 1).Value) <> NewestGetDate Then
            Exit For
        End If
        
        ' Example row/record structure from TenantTable:
        ' Column Header: Value
        '  1    "Get_Datetime": "07/05/2024  17:06:11"
        '  2    "id": "e1eede29-f875-4a5d-8470-17f6a29a88b1"
        '  3    "authEventId": "d99ecdfe-391d-43d2-b834-17636ba90e8d"
        '  4    "tenantId": "70784a63-d24b-46a9-a4db-0e70a274b056"
        '  5    "tenantType": "ORGANISATION"
        '  6    "tenantName": "Maple Florist"
        '  7    "createdDateUtc": "2019-07-09T23:40:30.1833130"
        '  8    "updatedDateUtc": "2020-05-15T01:35:13.8491980"
        
        ' Based on StoreTenants method (10 May 2024)
        
        ' Store in dictionary with "tenantName" as the key and "tenantId" as the value
        DTenantsId(CStr(TenantRow.Range.Cells(1, 6).Value)) = CStr(TenantRow.Range.Cells(1, 4).Value)
    Next TenantRow
    
    ' Log success
    WebHelpers.LogDebug "LoadTenants success, available tenants: " & CStr(DTenantsId.Count), "XeroAuthenticator.LoadTenants"
        
LoadTenants_Cleanup:
    Set wb = Nothing
    Set sh = Nothing
    Set TenantTable = Nothing
    Set TenantRow = Nothing
    
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        
        auth_ErrorDescription = "An error occurred while loading tenants from Excel sheet." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
    
        WebHelpers.LogError auth_ErrorDescription, "XeroAuthenticator.LoadTenants", 11041 + vbObjectError
        Err.Raise 11041 + vbObjectError, "XeroAuthenticator.LoadTenants", auth_ErrorDescription
    End If
End Sub

Private pDTenantsId As Dictionary


''
' Displays a user form that allows the user to select the Xero organization/tenant for the API request.
'
' @method SelectTenant
' @param {WebClient} Client - The WebClient instance to use for retrieving the list of tenants if necessary.
' @return {String} The selected tenant ID, or an empty string if the user cancels the selection.
'
' This function performs the following steps:
' 1. Loads the list of tenants from the saved Excel sheet using the 'LoadTenants' method.
' 2. If no tenants are loaded, calls the 'GetTenant' method to retrieve the list of tenants from the Xero API.
' 3. Initializes and displays the 'SelectXeroOrgForm' user form with the list of tenants.
' 4. If the user requests an update, calls the 'GetTenant' method again and refreshes the user form with the updated list.
' 5. If the user cancels the selection, raises an error and displays a message.
' 6. Returns the selected tenant ID if a tenant is selected.
'
' Note: This function uses and updates the 'DTenantsId' property.
''
Public Function SelectTenant(Client As WebClient) As String
    On Error GoTo auth_Cleanup
    
    ' Load the list of tenants from the saved Excel sheet
    LoadTenants
    
    ' If no tenants are loaded, retrieve the list from the Xero API
    If Not DTenantsId.Count > 0 Then
        GetTenant Client
    End If

    ' Initialize the 'SelectXeroOrgForm' user form
    Dim SelectForm1 As SelectXeroOrgForm
    Set SelectForm1 = New SelectXeroOrgForm
    
    ' Assign the list of tenants to the user form
    SelectForm1.OrgList = DTenantsId.Keys
    
    ' Display the user form
    SelectForm1.show
    
    ' Handle user requests for updating the tenant list
    If SelectForm1.UserRequestUpdate Then
        ' Get and update the tenant list
        GetTenant Client
        SelectForm1.OrgList = DTenantsId.Keys
        
        ' Reset the form property to false
        SelectForm1.UserRequestUpdate = False
        
        ' Show the form again with the updated list
        SelectForm1.show
    End If
    
    ' Check if the user canceled the selection
    If SelectForm1.UserCancel Then
        ' Notify the user and raise an error
        MsgBox "You canceled! The process is stopped.", vbInformation + vbOKOnly, "Xero Report Generator - Microsoft Excel"
        Err.Raise 11040 + vbObjectError, "SelectXeroOrgForm", "User canceled selection form"
    End If
    
    ' Return the selected tenant ID
    SelectTenant = DTenantsId(CStr(SelectForm1.ComboBox1.Value))

auth_Cleanup:
    ' Unload the user form and handle errors
    If Not SelectForm1 Is Nothing Then
        Unload SelectForm1
    End If
    
    ' Rethrow error
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        
        auth_ErrorDescription = "An error occurred during the user selection process for a Xero organization." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
    
        WebHelpers.LogError auth_ErrorDescription, "XeroAuthenticator.SelectTenant", 11041 + vbObjectError
        Err.Raise 11041 + vbObjectError, "XeroAuthenticator.SelectTenant", auth_ErrorDescription
    End If
End Function


''
' Clears/deletes all saved cache (tenants & tokens) from this Outlook StorageItem.
'
' @method ClearAllCache
'
' This function performs the following steps:
' 1. Checks the boolean flags to determine which cache to clear (tenants, tokens, or both).
' 2. If 'isClearTenant' is True, it calls the 'ClearTenants' function to clear tenant information.
' 3. If 'isClearToken' is True, it calls the 'ClearTokens' function to clear token information.
' 4. Handles any errors that occur during the process and logs them.
''
Public Sub ClearAllCache(isClearTenant As Boolean, isClearToken As Boolean)
    On Error GoTo auth_Cleanup
    
    ' Check if tenant cache needs to be cleared
    If isClearTenant Then
        ' Call the ClearTenants function
        ClearTenants
    End If
    
    ' Check if token cache needs to be cleared
    If isClearToken Then
        ' Call the ClearTokens function
        ClearTokens
    End If
    
auth_Cleanup:
    ' Error handling block
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        
        ' Construct the error description message
        auth_ErrorDescription = "An error occurred while clearing cache." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
        
        ' Log the error
        WebHelpers.LogError auth_ErrorDescription, "NinjaAuthenticator.ClearAllCache", 11041 + vbObjectError
        ' Raise the error for further handling
        Err.Raise 11041 + vbObjectError, "NinjaAuthenticator.ClearAllCache", auth_ErrorDescription
    End If
End Sub

''
' Hook for taking action before a request is executed; used for setting up required request headers for API calls.
' This method is called before every API request made to the Ninja API.
'
' Ninja official documentation on required headers for API calls:
' https://developer.xero.com/documentation/guides/oauth2/auth-flow/#6-call-the-api
'
' @param {WebClient} Client - The WebClient instance that is about to execute the request.
' @param in|out {WebRequest} Request - The request object that is about to be executed.
'
' This method performs the following steps:
' 1. Call the 'GetToken' method to ensure that the current access token is valid and not expired.
' 2. Set the 'Authorization' header with the current access token.
' 3. Call the 'SelectTenant' method to allow the user to select the Xero organization/tenant for the API request.
' 4. Set the 'xero-tenant-id' header with the selected tenant ID.
''
Private Sub IWebAuthenticator_BeforeExecute(ByVal Client As WebClient, ByRef Request As WebRequest)
    ' Ensure a valid access token is available for the API request
    Me.GetToken Client
    
    ' Set the 'Accept' header with the JSON format reference
    Request.SetHeader "Accept", "application/json"
    
    ' Set the 'Authorization' header with the current access token
    Request.SetHeader "Authorization", "Bearer " & Me.AccessToken
    
    ' Allow the user to select the Xero organization/tenant for the API request
    'Dim SelectedTenantId As String
    'SelectedTenantId = SelectTenant(Client)
    
    ' Set the 'xero-tenant-id' header with the selected tenant ID
    'Request.SetHeader "xero-tenant-id", SelectedTenantId
End Sub

''
' Requests the list of authorized Xero organizations/tenants from the Xero API.
' Official documentation: https://developer.xero.com/documentation/guides/oauth2/auth-flow/#5-check-the-tenants-youre-authorized-to-access
'
' @method GetTenant
' @param {WebClient} Client - The WebClient instance to use for the API request.
'
' This method performs the following steps:
' 1. Clear the 'DTenantsId' dictionary to store the new list of tenants.
' 2. Prepare the API request to retrieve the list of authorized tenants.
' 3. Execute the API request using the provided WebClient instance.
' 4. If the request is successful (200 status code), store the tenant information in the 'DTenantsId' dictionary and on an Excel sheet.
' 5. If the response is empty (no connected organizations), display a message and raise an error.
' 6. Handle any other errors that occur during the API request.
'
' Note: This method uses the 'AccessToken' property and updates the 'DTenantsId' dictionary.
''
Public Sub GetTenant(Client As WebClient)
    On Error GoTo auth_Cleanup
    
    Dim auth_TenantClient As WebClient
    Dim auth_Request As New WebRequest
    Dim auth_Body As New Dictionary
    Dim auth_Response As WebResponse
    
    ' Clear the 'DTenantsId' dictionary
    If DTenantsId.Count > 0 Then
        DTenantsId.RemoveAll
    End If
    
    ' Clone the provided WebClient instance to avoid accidental interactions
    Set auth_TenantClient = Client.Clone
    Set auth_TenantClient.Authenticator = Nothing
    auth_TenantClient.BaseUrl = "https://api.xero.com/"
    
    ' Prepare the API request to retrieve the list of authorized tenants
    auth_Request.Resource = "connections"
    auth_Request.Method = WebMethod.HttpGet
    auth_Request.RequestFormat = WebFormat.FormUrlEncoded
    auth_Request.ResponseFormat = WebFormat.Json
    auth_Request.SetHeader "Authorization", "Bearer " & Me.AccessToken
    
    Set auth_Response = auth_TenantClient.Execute(auth_Request)
    
    ' Process the API response
    If auth_Response.StatusCode = WebStatusCode.Ok Then
        If auth_Response.Data.Count > 0 Then
            ' Load response into "DTenantsId" property
            Dim TenantDict As Dictionary
            For Each TenantDict In auth_Response.Data
                DTenantsId(CStr(TenantDict("tenantName"))) = CStr(TenantDict("tenantId"))
            Next TenantDict
            
            ' Store response into excel sheet
            StoreTenants auth_Response
            
            WebHelpers.LogDebug "GetTenant success, available tenants: " & CStr(DTenantsId.Count), "XeroAuthenticator.GetTenant"
        Else
            MsgBox "Cannot find any Xero organization connected to your account, try to re-login and authorize a connection!", vbCritical + vbOKOnly, "Xero Report Generator - Microsoft Excel"
            Err.Raise 11041 + vbObjectError, "XeroAuthenticator.GetTenant", _
                auth_Response.StatusCode & ": " & auth_Response.Content & vbNewLine & _
                "Empty response; no connected organization, need to authorize and connect through Login"
        End If
    Else
        ' Raise an error if the API request is not successful
        Err.Raise 11041 + vbObjectError, "XeroAuthenticator.GetTenant", _
            auth_Response.StatusCode & ": " & auth_Response.Content
    End If
    
auth_Cleanup:
    ' Clean up objects and handle errors
    Set auth_TenantClient = Nothing
    Set auth_Request = Nothing
    Set auth_Response = Nothing
    
    ' Rethrow error
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        
        auth_ErrorDescription = "An error occurred while retrieving new tenants." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
    
        WebHelpers.LogError auth_ErrorDescription, "XeroAuthenticator.GetTenant", 11041 + vbObjectError
        Err.Raise 11041 + vbObjectError, "XeroAuthenticator.GetTenant", auth_ErrorDescription
    End If
End Sub

''
' Clears/deletes all saved tenant information from this Excel workbook.
'
' @method ClearTenants
'
' This function performs the following steps:
' 1. Checks if the sheet and table from which to clear tenant information exist.
' 2. If they do not exist, notifies the user and exits.
' 3. If found, proceeds to delete the table and its contents.
' 4. Notifies the user of the result, performs cleanup, and handles errors.
'
' Note: This function uses the 'auth_SheetTenantsStorage' and 'auth_TableTenantsStorage' constants.
''
Private Sub ClearTenants()
    On Error GoTo ClearTenants_Cleanup
    
    Dim wb As Workbook
    Dim sh As Worksheet
    Dim TenantTable As ListObject
    
    Set wb = ThisWorkbook
    ' Check if the sheet from which to clear tenant information exists
    If Not WebHelpers.WorksheetExists(auth_SheetTenantsStorage, wb) Then
        ' Log failure due to non-existent sheet
        WebHelpers.LogDebug "Clear tenants cache failed, sheet: " & auth_SheetTenantsStorage & "does not exist", "XeroAuthenticator.ClearTenants"
        ' Notify the user
        MsgBox "Clearing Xero organizations cache failed; Xero organizations cache did not exist.", vbCritical + vbOKOnly, "Xero Report Generator - Microsoft Excel"
        ' Exit the subroutine as the sheet does not exist
        Exit Sub
    Else
        ' Sheet exists, set the sheet variable
        Set sh = wb.Sheets(auth_SheetTenantsStorage)
    End If
    
    ' Check if the ListObject (table) from which to clear tenant information exists
    If Not WebHelpers.ListObjectExists(auth_TableTenantsStorage, sh) Then
        ' Log failure due to non-existent table
        WebHelpers.LogDebug "Clear tenants cache failed, ListObject(Table): " & auth_TableTenantsStorage & " does not exist", "XeroAuthenticator.ClearTenants"
        ' Notify the user
        MsgBox "Clearing Xero organizations cache failed; Xero organizations cache did not exist.", vbCritical + vbOKOnly, "Xero Report Generator - Microsoft Excel"
        ' Exit the subroutine as the table does not exist
        Exit Sub
    Else
        ' Table exists, set the table variable
        Set TenantTable = sh.ListObjects(auth_TableTenantsStorage)
    End If
    
    ' Delete the table (ListObject) and all of its rows/records
    TenantTable.Delete
    
    ' Log success of clearing tenant information
    WebHelpers.LogDebug "Clearing tenants cache succeeded; all records of tenant information were deleted.", "XeroAuthenticator.ClearTenants"
    ' Notify the user of successful deletion
    MsgBox "Clearing Xero organizations cache succeeded; all records of Xero organizations were deleted.", vbInformation + vbOKOnly, "Xero Report Generator - Microsoft Excel"

ClearTenants_Cleanup:
    ' Cleanup object variables
    Set wb = Nothing
    Set sh = Nothing
    Set TenantTable = Nothing
    
    ' Error handling block
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        
        ' Construct the error description message
        auth_ErrorDescription = "An error occurred while clearing tenants cache from Excel sheet." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
    
        ' Log the error
        WebHelpers.LogError auth_ErrorDescription, "XeroAuthenticator.ClearTenants", 11041 + vbObjectError
        ' Raise the error for further handling
        Err.Raise 11041 + vbObjectError, "XeroAuthenticator.ClearTenants", auth_ErrorDescription
    End If
End Sub


''
' Limits the number of rows in a table, deleting rows if the count exceeds a specified limit.
' Used for limiting the rows/records in the tokens table within the StoreTokens procedure.
'
' @method LimitTableRows
' @param {ListObject} Table - The table to limit.
' @param {Integer} IRowsLimit - The maximum number of rows allowed.
'
' This function performs the following steps:
' 1. Checks if the row limit and the number of rows in the table are greater than zero.
' 2. If either is zero, logs the issue and exits.
' 3. If the number of rows is already within the limit, logs the status and exits.
' 4. If the number of rows exceeds the limit, enters a loop to delete rows from the bottom until the limit is met.
' 5. Logs the successful completion of the row limiting process.
''
Private Sub LimitTableRows(Table As ListObject, IRowsLimit As Integer)
    
    Dim TblRows As ListRows
    Set TblRows = Table.ListRows
    
    ' Exit sub if the limit or table rows count is zero
    If Not IRowsLimit > 0 Or Not TblRows.Count > 0 Then
        WebHelpers.LogDebug "Table " & Table.name & " - limit number or table rows can not be zero - " & _
                            "IRowsLimit: " & IRowsLimit & _
                            ", TblRows.Count: " & TblRows.Count, _
                            "XeroAuthenticator.LimitTableRows"
        Exit Sub
    ' Exit sub if the table rows count is already under or equal to the limit
    ElseIf Not TblRows.Count > IRowsLimit Then
        WebHelpers.LogDebug "Table " & Table.name & " - rows is already under/equal to the limit", _
                            "XeroAuthenticator.LimitTableRows"
        Exit Sub
    End If
    
    ' Loop while the table rows count is greater than the limit
    Do While TblRows.Count > IRowsLimit
        ' Delete the bottom row from the table
        TblRows(TblRows.Count).Delete
    Loop
    
    ' Log successful row limiting
    WebHelpers.LogDebug "Table " & Table.name & " - successfully limited to only " & TblRows.Count & " rows", _
                        "XeroAuthenticator.LimitTableRows"
End Sub

