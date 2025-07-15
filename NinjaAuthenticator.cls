VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NinjaAuthenticator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' NinjaAuthenticator v0.0.1
' @author NeumaC
' https://github.com/NeumaC/NinjaRMM_VBA
'
' Ninja Authenticator
' Custom IWebAuthenticator for "Ninja" authentication for NinjaRMM API
' Source (vba-xero-api) credits: Musa Yohanes - https://github.com/Muyoouu/vba-xero-api
' Source (IWebAuthenticator) credits: Tim Hall - https://github.com/VBA-tools/VBA-Web
'
' Ninja authentication docs:
' - https://app.ninjarmm.com/apidocs-beta/authorization/overview
'
' Other docs:
' - Register Native Applications: https://app.ninjarmm.com/apidocs-beta/authorization/create-applications/native-apps
' - Authorization Code Flow with Proof Key for Code Exchange (PKCE): https://app.ninjarmm.com/apidocs-beta/authorization/flows/authorization-code-flow-with-proof-key-for-code-exchange-pkce
'
' @example
' ```VB.net
' Dim Auth As NinjaAuthenticator
' Set Auth = New NinjaAuthenticator
' Auth.Setup "Your Client Id", "Your Client Secret"
'
' ' Add scopes, for example request monitoring (read) access
' ' Request refresh token with 'offline_access' scope included
' Auth.AddScope "offline_access"
' Auth.AddScope "Monitoring"
'
' ' Manually open up Ninja login
' ' (called automatically on first request otherwise)
' Auth.Login
'
' ' Add authenticator to client
' Set Client.Authenticator = Auth
' ```
'
' @class NinjaAuthenticator
' @implements IWebAuthenticator v4.*
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Implements IWebAuthenticator
Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

' URL for login process
Private Const auth_AuthorizationUrl As String = "https://eu.ninjarmm.com/ws/oauth/authorize"
' Redirect URL registered inside Xero
Private Const auth_RedirectUrl As String = "http://127.0.0.1:12345/auth"
' StorageItem name to store reusable tokens (refresh token - valid for 60 days)
Private Const auth_StorageItemTokensStorage As String = "StorageItem_NinjaTokens"
' Folder(Outlook.Folder) name to store Ninja token info (and ClientID)
Private Const auth_TokensStorage As String = "Tickets"

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

Public ClientId As String
Public CodeChallenge As String
Public CodeVerifier As String
Public ApiKey As String
Public AuthorizationCode As String
Public AccessToken As String
Public AccessTokenExpiration As Date
Public RefreshToken As String
Public RefreshTokenExpiration As Date
Public Scopes As Variant

' ============================================= '
' Public Methods
' ============================================= '

''
' Sets up the authenticator properties required for Ninja API authentication.
'
' @method Setup
' @param {String} ClientId - The client ID provided by Ninja for your application.
'
' This method initializes the 'ClientId' property with the provided value.
' These properties are used during the authentication and token exchange processes.
''

Public Sub Setup(ClientId As String)
    Me.ClientId = ClientId
    Me.CodeVerifier = BCrypt.Random(TextSize:=32, TrueBase64:=True, UrlSafeToggle:=True)

    Dim auth_VerifierBytes() As Byte
    Dim auth_Hash() As Byte

    ' Convert the CodeVerifier to raw ASCII bytes before hashing
    auth_VerifierBytes = StrConv(Me.CodeVerifier, vbFromUnicode)
    
    auth_Hash = BCrypt.HashData(auth_VerifierBytes, bcSha256)
    Me.CodeChallenge = BCrypt.Hash(Text:=Me.CodeVerifier, BcryptHashAlgorithmId:=bcSha256, UrlSafeToggle:=True)
End Sub

''
' Direct user to login and authenticate in Xero website via browser.
' Official documentation: https://developer.xero.com/documentation/guides/oauth2/auth-flow/#1-send-a-user-to-authorize-your-app
'
' @method Login
'
' The login process involves the following steps:
' 1. Open a browser instance using the ChromeDevProtocols framework (https://github.com/longvh211/Chromium-Automation-with-CDP-for-VBA).
' 2. Wait for the user to complete the login process, either by approval or encountering an error.
' 3. Extract the authentication code from the browser URL if the login is approved.
' 4. Handle any errors that occur during the login process.
'
' If the login is successful, the 'AuthorizationCode' property is set with the extracted authentication code.
' If an error occurs or the login is canceled by the user, an appropriate error message is displayed.
''
Public Sub Login()
    On Error GoTo auth_Cleanup
    
    ' Open browser using ChromeDevProtocols framework
    Dim objBrowser As New CDPBrowser
    objBrowser.start name:="edge", appUrl:=Me.GetLoginUrl, cleanActive:=True, reAttach:=True
    
    ' Start local TCP Listener for authorization code retrieval
    Dim listenerOutput As String
    listenerOutput = tcp_StartListener
    
    ' Boolean values to track login completion
    Dim auth_IsApprove As Boolean
    Dim auth_IsError As Boolean
    
    ' Wait for login to complete, either by approval or error
    Do
        auth_IsApprove = auth_LoginIsApproval(listenerOutput)
        auth_IsError = auth_LoginIsError(objBrowser)
        ' Sleep to avoid overloading the system
        objBrowser.sleep 1
    Loop Until auth_IsApprove Or auth_IsError
    
    ' Handle errors or extract the authentication code
    If auth_IsError Then
        Err.Raise 11040 + vbObjectError, "OAuthDialog", auth_LoginExtractError(objBrowser)
    Else
        Me.AuthorizationCode = auth_LoginExtractCode(listenerOutput)
        If Me.AuthorizationCode = "" Then
            MsgBox "Login might be canceled! process is stopped", vbCritical + vbOKOnly, "NinjaRMM Ticket Connector - Microsoft Outlook"
            Err.Raise 11040 + vbObjectError, "OAuthDialog", "Login was canceled"
        End If
    End If
    
auth_Cleanup:
    ' Error handling and cleanup
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        auth_ErrorDescription = "An error occurred while logging in." & vbNewLine
        If Err.Number - vbObjectError <> 11040 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
        
        ' Close browser
        If Not objBrowser Is Nothing Then: objBrowser.quit
        Set objBrowser = Nothing
        
        WebHelpers.LogError auth_ErrorDescription, "NinjaAuthenticator.Login", 11040 + vbObjectError
        Err.Raise 11040 + vbObjectError, "NinjaAuthenticator.Login", auth_ErrorDescription
    Else
        ' Close browser
        If Not objBrowser Is Nothing Then: objBrowser.quit
        Set objBrowser = Nothing
        
        If Not auth_IsApprove Then
            auth_ErrorDescription = "Login did not complete." & vbNewLine
            
            WebHelpers.LogError auth_ErrorDescription, "NinjaAuthenticator.Login", 11040 + vbObjectError
            Err.Raise 11040 + vbObjectError, "NinjaAuthenticator.Login", auth_ErrorDescription
        Else
            WebHelpers.LogDebug "Login succeeded - auth code: " & Me.AuthorizationCode, "NinjaAuthenticator.Login"
            ' Inform user that the login process succeeded
            MsgBox "Ninja login succeeded!", vbInformation + vbOKOnly, "NinjaRMM Ticket Connector - Microsoft Outlook"
        End If
    End If
End Sub

''
' Logs out by clearing all cached tokens and authorizations, effectively resetting the login session.
' After calling this method, the user will need to log in again to obtain new tokens and authorizations.
'
' @method Logout
''
Public Sub Logout()
    Me.AuthorizationCode = ""
    Me.AccessToken = ""
    Me.AccessTokenExpiration = CDate(0)
    Me.RefreshToken = ""
    Me.RefreshTokenExpiration = CDate(0)
'    Me.CodeChallenge = ""
'    Me.CodeVerifier = ""
End Sub

''
' Adds a new scope to the list of authorized scopes for the Xero API.
' To get a list of available scopes, visit https://developer.xero.com/documentation/guides/oauth2/scopes/
' The 'Scopes' property is used during the authentication process to request access to specific areas of the Xero API.
'
' @method AddScope
' @param {String} Scope - The scope to add to the authorized scopes list.
''
Public Sub AddScope(Scope As String)
    Dim auth_Scopes As Variant
    
    ' Retrieve the current list of authorized scopes
    auth_Scopes = Me.Scopes
    
    ' Initialize a new array if the 'Scopes' property is empty
    If VBA.IsEmpty(auth_Scopes) Then
        ReDim auth_Scopes(0 To 0)
    Else
        ' Resize the array to accommodate the new scope
        ReDim Preserve auth_Scopes(0 To UBound(auth_Scopes) + 1)
    End If
    
    ' Add the new scope to the end of the array
    auth_Scopes(UBound(auth_Scopes)) = Scope
    
    ' Update the 'Scopes' property with the modified array
    Me.Scopes = auth_Scopes
End Sub

''
' Obtains a new or refreshed access token and refresh token, if necessary.
'
' @method GetToken
' @param {WebClient} Client - The WebClient instance to use for token requests.
'
' This method follows the following logic:
' 1. If the current access token exists and is not expired, no action is taken.
' 2. If the current refresh token exists and is not expired, the 'GetRefreshedToken' method is called to obtain a new access token and refresh token.
' 3. If no valid access token or refresh token exists:
'    a. If the 'AuthorizationCode' is empty, the 'Login' method is called to obtain a new authorization code.
'    b. The 'GetNewToken' method is called to obtain a new access token and refresh token using the authorization code.
' 4. The new or refreshed tokens are stored by calling the 'StoreTokens' method.
'
' Note: This method uses and updates the 'AccessToken', 'AccessTokenExpiration', 'RefreshToken', 'RefreshTokenExpiration', and 'AuthorizationCode' properties.
''
Public Sub GetToken(client As WebClient)
        
    ' Check if the current access token exists and is not expired
    If Me.AccessToken <> "" And Not Me.AccessTokenExpiration < Now Then
        WebHelpers.LogDebug "AccessToken already exist and not expired, skipping GetToken process", "NinjaAuthenticator.GetToken"
        Exit Sub
    ' Check if the current refresh token exists and is not expired
    ElseIf Me.RefreshToken <> "" And Not Me.RefreshTokenExpiration < Now Then
        WebHelpers.LogDebug "RefreshToken exist and not expired, attempting to refresh AccessToken", "NinjaAuthenticator.GetToken"
        GetRefreshedToken client
    ' If no valid tokens exist, obtain a new access token and refresh token
    Else
        ' Login first if the authorization code doesn't exist
        If Me.AuthorizationCode = "" Then
            WebHelpers.printMsg "NinjaAuthenticator.GetToken", "Auth code not found, redirecting to Login"
            Me.Login
        End If
        WebHelpers.LogDebug "Tokens not found or expired, getting new tokens", "NinjaAuthenticator.GetToken"
        GetNewToken client
    End If
    
    ' Store the new or refreshed tokens
    StoreTokens
End Sub



''
' Clears/deletes all saved cache (tokens) from this Outlook StorageItem.
'
' @method ClearAllCache
'
' This function performs the following steps:
' 1. Checks the boolean flags to determine which cache to clear (only tokens at this time).
' 2. If 'isClearToken' is True, it calls the 'ClearTokens' function to clear token information.
' 3. Handles any errors that occur during the process and logs them.
''
Public Sub ClearAllCache(isClearToken As Boolean)
    On Error GoTo auth_Cleanup
    
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
''
Private Sub IWebAuthenticator_BeforeExecute(ByVal client As WebClient, ByRef Request As WebRequest)
    ' Ensure a valid access token is available for the API request
    Me.GetToken client
    
    ' Set the 'Accept' header with the JSON format reference
    Request.SetHeader "Accept", "application/json"
    
    ' Set the 'Authorization' header with the current access token
    Request.SetHeader "Authorization", "Bearer " & Me.AccessToken
    
End Sub

''
' Hook for taking action after request has been executed
'
' @param {WebClient} Client The client that executed request
' @param {WebRequest} Request The request that was just executed
' @param in|out {WebResponse} Response to request
''
Private Sub IWebAuthenticator_AfterExecute(ByVal client As WebClient, ByVal Request As WebRequest, ByRef Response As WebResponse)
    ' e.g. Handle 401 Unauthorized or other issues
End Sub

''
' Hook for updating http before send
'
' @param {WebClient} Client
' @param {WebRequest} Request
' @param in|out {WinHttpRequest} Http
''
Private Sub IWebAuthenticator_PrepareHttp(ByVal client As WebClient, ByVal Request As WebRequest, ByRef Http As Object)
    ' e.g. Update option, headers, etc.
End Sub

''
' Hook for updating cURL before send
'
' @param {WebClient} Client
' @param {WebRequest} Request
' @param in|out {String} Curl
''
Private Sub IWebAuthenticator_PrepareCurl(ByVal client As WebClient, ByVal Request As WebRequest, ByRef Curl As String)
    ' e.g. Add flags to cURL
End Sub

''
' Formats the URL for the Ninja login process, including the required scopes and other parameters.
'
' @method GetLoginUrl
' @return {String} The formatted login URL.
'
' The formatted URL includes the following parameters:
' - response_type: Set to "code" to indicate that an authorization code will be returned.
' - client_id: The client ID provided by Ninja for authentication.
' - redirect_uri: The predetermined redirect URL to which the user will be directed after authentication.
' - state: An optional state parameter, currently set to a hardcoded value of "123".
' - scope: The list of scopes (permissions) requested for the application. If no scopes are specified, an empty string is used.
'
' Docs: https://app.ninjarmm.com/apidocs-beta/authorization/flows/authorization-code-flow-with-proof-key-for-code-exchange-pkce
''
Public Function GetLoginUrl() As String
    ' Use WebRequest object as helper to format URL
    Dim auth_Request As New WebRequest
    auth_Request.Resource = auth_AuthorizationUrl
    
    ' Add required parameters
    auth_Request.AddQuerystringParam "response_type", "code"
    auth_Request.AddQuerystringParam "client_id", Me.ClientId
    auth_Request.AddQuerystringParam "redirect_uri", auth_RedirectUrl
    auth_Request.AddQuerystringParam "state", "123"
    
    ' Add scopes, or an empty string if no scopes are specified
    If Not VBA.IsEmpty(Me.Scopes) Then
        auth_Request.AddQuerystringParam "scope", VBA.Join(Me.Scopes, " ")
    Else
        auth_Request.AddQuerystringParam "scope", ""
    End If
    
    ' Add code challenge and code challenge method for PKCE
    auth_Request.AddQuerystringParam "code_challenge", Me.CodeChallenge
    auth_Request.AddQuerystringParam "code_challenge_method", "S256"
    
    ' Get the formatted URL and clean up
    GetLoginUrl = auth_Request.FormattedResource
    Set auth_Request = Nothing
End Function

' ============================================= '
' Private Methods
' ============================================= '

''
' Checks if the current URL of the browser matches the predetermined redirect URL stored in the 'auth_RedirectUrl' constant.
'
' @method auth_LoginIsApproval
' @param {CDPBrowser} objBrowser - The browser object to check the URL from.
' @return {Boolean} True if the current URL matches the redirect URL, False otherwise.
''
'Private Function auth_LoginIsApproval(objBrowser As CDPBrowser) As Boolean
'    Dim auth_BrowserUrlParts As Dictionary
'    Dim auth_RedirectUrlParts As Dictionary
'
'    Set auth_BrowserUrlParts = WebHelpers.GetUrlParts(objBrowser.url)
'    Set auth_RedirectUrlParts = WebHelpers.GetUrlParts(auth_RedirectUrl)
'
'    ' Check if the 'Host' part of both URLs match.
'    If auth_BrowserUrlParts("Host") = auth_RedirectUrlParts("Host") Then
'        auth_LoginIsApproval = True
'    Else
'        auth_LoginIsApproval = False
'    End If
'End Function

''
' Checks if an output file from the TCP listener exists.
'
' @method auth_LoginIsApproval
' @param {CDPBrowser} objBrowser - The browser object to check the URL from.
' @return {Boolean} True if the current URL matches the redirect URL, False otherwise.
''
Private Function auth_LoginIsApproval(outputPath As String) As Boolean

    ' Check if the 'Host' part of both URLs match.
    If tcp_PollListener(outputPath) <> "" Then
        auth_LoginIsApproval = True
    Else
        auth_LoginIsApproval = False
    End If
End Function

''
' Checks if the current HTML page rendered in the browser contains an error message in the title.
'
' @method auth_LoginIsError
' @param {CDPBrowser} objBrowser - The browser object to check for errors.
' @return {Boolean} True if the current page title contains the word 'error', False otherwise.
''
Private Function auth_LoginIsError(objBrowser As CDPBrowser) As Boolean
    Dim auth_Element As CDPElement
    
    ' Wait for the page to fully load before checking for errors.
    objBrowser.wait till:=ReadyState.isComplete
    
    ' Search for the 'error' word in the HTML title.
    Set auth_Element = objBrowser.getElementByQuery("title", True)
    
    ' If the title element exists, check if it contains the word 'error'.
    If auth_Element.isExist Then
        auth_LoginIsError = InStr(1, CStr(auth_Element.innerText), "error", vbTextCompare) > 0
    Else
        ' If the title element does not exist, assume no error.
        auth_LoginIsError = False
    End If
End Function

'''
'' Extracts the authentication code from the current URL of the browser if it matches the predetermined redirect URL.
''
'' @method auth_LoginExtractCode
'' @param {CDPBrowser} objBrowser - The browser object to extract the URL from.
'' @return {String} The extracted authentication code, or an empty string if the URL does not match the redirect URL.
'''
'Private Function auth_LoginExtractCode(objBrowser As CDPBrowser) As String
'    ' Check if the current URL matches the predetermined redirect URL.
'    If auth_LoginIsApproval(objBrowser) Then
'        Dim auth_UrlParts As Dictionary
'        Dim auth_ParsedUrlParams As Dictionary
'
'        ' Extract the authorization code from the 'code' parameter in the URL query string.
'        Set auth_UrlParts = WebHelpers.GetUrlParts(objBrowser.url)
'        Set auth_ParsedUrlParams = ParseUrlEncoded(auth_UrlParts("Querystring"))
'        auth_LoginExtractCode = auth_ParsedUrlParams("code")
'    Else
'        ' If the URL does not match the redirect URL, return an empty string.
'        auth_LoginExtractCode = ""
'    End If
'End Function

''
' Extracts the authentication code from the TCP listener output file.
'
' @method auth_LoginExtractCode
' @param {string} outputPath - The output file to extract the URL from.
' @return {String} The extracted authentication code, or an empty string if the URL does not match the redirect URL.
''
Private Function auth_LoginExtractCode(outputPath As String) As String
    ' Check if the current URL matches the predetermined redirect URL.
    If auth_LoginIsApproval(outputPath) Then
        Dim auth_UrlParts As Dictionary
        Dim auth_ParsedUrlParams As Dictionary
        
        ' Extract the authorization code from the 'code' parameter in the URL query string.
        Set auth_UrlParts = WebHelpers.GetUrlParts(tcp_PollListener(outputPath))
        Set auth_ParsedUrlParams = ParseUrlEncoded(auth_UrlParts("Querystring"))
        auth_LoginExtractCode = auth_ParsedUrlParams("code")
    Else
        ' If the URL does not match the redirect URL, return an empty string.
        auth_LoginExtractCode = ""
    End If
End Function

''
' Extracts the authentication code from the listening socket.
'
' @method auth_LoginExtractCode
' @param {CDPBrowser} objBrowser - The browser object to extract the URL from.
' @return {String} The extracted authentication code, or an empty string if the URL does not match the redirect URL.
''
'Private Function auth_LoginExtractCode(objBrowser As CDPBrowser) As String
'    ' Check if the current URL matches the predetermined redirect URL.
'    If auth_LoginIsApproval(objBrowser) Then
'        Dim auth_UrlParts As Dictionary
'        Dim auth_ParsedUrlParams As Dictionary
'
'        ' Extract the authorization code from the 'code' parameter in the URL query string.
'        Set auth_UrlParts = WebHelpers.GetUrlParts(objBrowser.url)
'        Set auth_ParsedUrlParams = ParseUrlEncoded(auth_UrlParts("Querystring"))
'        auth_LoginExtractCode = auth_ParsedUrlParams("code")
'    Else
'        ' If the URL does not match the redirect URL, return an empty string.
'        auth_LoginExtractCode = ""
'    End If
'End Function

''
' Extracts the error descriptions in the current HTML page rendered in the browser.
'
' @method auth_LoginExtractError
' @param {CDPBrowser} objBrowser - The browser object to extract error descriptions from.
' @return {String} Returns extracted error descriptions if found, otherwise "Unknown".
'
' This function performs the following steps:
' 1. Uses the browser object to find an element containing the text "Error".
' 2. If the error element is found, it extracts the inner text of the element as the error code.
' 3. Then, it looks for a following sibling element of the error element to extract the error description.
' 4. If both elements are found, it concatenates the error code and description.
' 5. If no error elements are found, it returns "Unknown".
''
Private Function auth_LoginExtractError(objBrowser As CDPBrowser) As String
    Dim auth_Element As CDPElement
    
    ' Attempt to find an element containing "Error" in its text
    Set auth_Element = objBrowser.getElementByXPath("//p[contains(concat(' ', text(), ' '), ' Error ')]")
    
    If auth_Element.isExist Then
        ' Extract the error code from the element
        auth_LoginExtractError = auth_Element.innerText
        
        ' Attempt to find the following sibling element containing the error description
        Set auth_Element = objBrowser.getElementByXPath("//p[contains(concat(' ', text(), ' '), ' Error ')]/following-sibling::p")
        If auth_Element.isExist Then
            ' Concatenate the error code and the error description
            auth_LoginExtractError = auth_LoginExtractError & " - " & auth_Element.innerText
        End If
    Else
        ' Return "Unknown" if no error element is found
        auth_LoginExtractError = "Unknown"
    End If
End Function

''
' Exchanges the current AuthorizationCode for new access and refresh tokens.
' Official documentation: https://app.ninjarmm.com/apidocs-beta/authorization/flows/authorization-code-flow-with-proof-key-for-code-exchange-pkce
'
' @method GetNewToken
' @param {WebClient} Client - The WebClient instance to use for the token exchange request.
'
' This method performs the following steps:
' 1. Prepare the token request by setting the required headers, request body, and other parameters.
' 2. Execute the token request using the provided WebClient instance.
' 3. If the request is unauthorized (401 status code), attempt to re-login and update the AuthorizationCode, then retry the token request.
' 4. If the token request is successful (200 status code), update the AccessToken, AccessTokenExpiration, RefreshToken, and RefreshTokenExpiration properties.
' 5. Handle any errors that occur during the token exchange process.
'
' Note: This method uses the 'ClientId', 'CodeVerifier', 'AuthorizationCode', and 'auth_RedirectUrl' properties.
''
Private Sub GetNewToken(client As WebClient)
    On Error GoTo auth_Cleanup
    
    Dim auth_TokenClient As WebClient
    Dim auth_Request As New WebRequest
    Dim auth_Body As New Dictionary
    Dim auth_Response As WebResponse
    
    ' Clone the provided WebClient instance to avoid accidental interactions
    Set auth_TokenClient = client.Clone
    Set auth_TokenClient.Authenticator = Nothing
    auth_TokenClient.BaseUrl = "https://eu.ninjarmm.com/ws/oauth/"
    
    ' Prepare the token request
    auth_Request.Resource = "token"
    auth_Request.Method = WebMethod.HttpPost
    auth_Request.RequestFormat = WebFormat.FormUrlEncoded
    auth_Request.ResponseFormat = WebFormat.Json
    'auth_Request.SetHeader "authorization", "Basic " & Base64Encode(Me.ClientId & ":" & Me.ClientSecret)
    
    ' Prepare the request body
    auth_Body.Add "grant_type", "authorization_code"
    auth_Body.Add "client_id", Me.ClientId
    auth_Body.Add "code", Me.AuthorizationCode
    auth_Body.Add "code_verifier", Me.CodeVerifier
    auth_Body.Add "redirect_uri", auth_RedirectUrl
    auth_Body.Add "scope", "offline_access monitoring"
    Set auth_Request.Body = auth_Body
    
    Set auth_Response = auth_TokenClient.Execute(auth_Request)
    
    ' Response 401 - auth code might be expired
    ' Handle unauthorized (401) response by attempting to re-login and update the AuthorizationCode
    If auth_Response.StatusCode = WebStatusCode.Unauthorized Then
        ' Log the response details
        WebHelpers.printMsg "NinjaAuthenticator.GetNewToken", auth_Response.StatusDescription & " - " & auth_Response.Data("details")
        
        ' Notify and ask permission to attempt re-login
        Dim msgBoxResponse As VbMsgBoxResult
        msgBoxResponse = MsgBox("Login session expired:" & vbNewLine & auth_Response.Data("details") & _
            vbNewLine & vbNewLine & "Attempt to re-login?", vbQuestion + vbYesNo, "NinjaRMM Ticket Connector - Microsoft Outlook")
        
        Select Case msgBoxResponse
            Case vbYes
                WebHelpers.printMsg "NinjaAuthenticator.GetNewToken", "Redirecting to Login..."
        
                ' Clear expired auth code and token
                Me.Logout
                
                ' Re-login and update the auth code
                Me.Login
                auth_Request.Body("code") = Me.AuthorizationCode
                
                ' Re-attempt request to get token
                Set auth_Response = auth_TokenClient.Execute(auth_Request)
            
            ' Notify stopping the get token process
            ' Leaving response status at 401 will raise error eventually
            Case vbNo
                WebHelpers.printMsg "NinjaAuthenticator.GetNewToken", "Abort get token process"
        End Select
    End If
    
    ' Update token properties if the token request is successful
    If auth_Response.StatusCode = WebStatusCode.Ok Then
        Me.AccessToken = auth_Response.Data("access_token")
        Me.AccessTokenExpiration = DateAdd("s", CLng(auth_Response.Data("expires_in")), Now)
        If auth_Response.Data.Exists("refresh_token") Then
            Me.RefreshToken = auth_Response.Data("refresh_token")
            Me.RefreshTokenExpiration = DateAdd("d", 60, Now)
        End If
        
        WebHelpers.LogDebug "New Access Token: " & Me.AccessToken & vbNewLine & vbNewLine & _
            "New Refresh Token: " & vbNewLine & Me.RefreshToken & vbNewLine, "NinjaAuthenticator.GetNewToken"
    Else
        ' Raise an error if the token request is not successful
        Err.Raise 11041 + vbObjectError, "NinjaAuthenticator.GetNewToken", _
            auth_Response.StatusCode & ": " & auth_Response.Content
    End If
    
auth_Cleanup:
    ' Clean up objects and handle errors
    Set auth_TokenClient = Nothing
    Set auth_Request = Nothing
    Set auth_Response = Nothing
    
    ' Rethrow error
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        
        auth_ErrorDescription = "An error occurred while retrieving new token." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
    
        WebHelpers.LogError auth_ErrorDescription, "NinjaAuthenticator.GetNewToken", 11041 + vbObjectError
        Err.Raise 11041 + vbObjectError, "NinjaAuthenticator.GetNewToken", auth_ErrorDescription
    End If
End Sub

''
' Exchanges the current refresh token for new access and refresh tokens.
' Official documentation: https://app.ninjarmm.com/apidocs-beta/authorization/refresh-tokens
'
' @method GetRefreshedToken
' @param {WebClient} Client - The WebClient instance to use for the token exchange request.
'
' This method performs the following steps:
' 1. Prepare the token refresh request by setting the required headers, request body, and other parameters.
' 2. Execute the token refresh request using the provided WebClient instance.
' 3. If the request is successful (200 status code), update the AccessToken, AccessTokenExpiration, RefreshToken, and RefreshTokenExpiration properties.
' 4. If the request is unauthorized (401 status code), attempt to re-login and update the AuthorizationCode, then redirect to the GetNewToken procedure.
' 5. Handle any other errors that occur during the token refresh process.
'
' Note: This method uses the 'ClientId', 'ClientSecret', and 'RefreshToken' properties.
''
Private Sub GetRefreshedToken(client As WebClient)
    On Error GoTo auth_Cleanup
    
    Dim auth_TokenClient As WebClient
    Dim auth_Request As New WebRequest
    Dim auth_Body As New Dictionary
    Dim auth_Response As WebResponse
    
    ' Clone the provided WebClient instance to avoid accidental interactions
    Set auth_TokenClient = client.Clone
    Set auth_TokenClient.Authenticator = Nothing
    auth_TokenClient.BaseUrl = "https://eu.ninjarmm.com/ws/oauth/"
    
    ' Prepare the token refresh request
    auth_Request.Resource = "token"
    auth_Request.Method = WebMethod.HttpPost
    auth_Request.RequestFormat = WebFormat.FormUrlEncoded
    auth_Request.ResponseFormat = WebFormat.Json
    'auth_Request.SetHeader "authorization", "Basic " + Base64Encode(Me.ClientId & ":" & Me.ClientSecret)
    
    ' Prepare the request body
    auth_Body.Add "grant_type", "refresh_token"
    auth_Body.Add "client_id", Me.ClientId
    auth_Body.Add "refresh_token", Me.RefreshToken
    Set auth_Request.Body = auth_Body
    
    Set auth_Response = auth_TokenClient.Execute(auth_Request)
    
    ' Update token properties if the token refresh request is successful
    If auth_Response.StatusCode = WebStatusCode.Ok Then
        Me.AccessToken = auth_Response.Data("access_token")
        Me.AccessTokenExpiration = DateAdd("s", CLng(auth_Response.Data("expires_in")), Now)
        Me.RefreshToken = auth_Response.Data("refresh_token")
        Me.RefreshTokenExpiration = DateAdd("d", 60, Now)
        
        WebHelpers.LogDebug "New Access Token: " & Me.AccessToken & vbNewLine & vbTab & _
            "New Refresh Token: " & Me.RefreshToken, "NinjaAuthenticator.GetNewToken"
    
    ' Response 401 - refresh token might be expired
    ' Handle unauthorized (401) response by attempting to re-login and update the AuthorizationCode
    ElseIf auth_Response.StatusCode = WebStatusCode.Unauthorized Then
        ' Log the response details
        WebHelpers.printMsg "NinjaAuthenticator.GetRefreshedToken", auth_Response.StatusDescription & " - " & auth_Response.Data("details")
        WebHelpers.LogDebug "NinjaAuthenticator.GetRefreshedToken", auth_Response.StatusCode & ": " & auth_Response.Content
        
        ' Notify and ask permission to attempt re-login
        Dim msgBoxResponse As VbMsgBoxResult
        msgBoxResponse = MsgBox("Login session expired:" & vbNewLine & auth_Response.Data("details") & _
            vbNewLine & vbNewLine & "Attempt to re-login?", vbQuestion + vbYesNo, "NinjaRMM Ticket Connector - Microsoft Outlook")
        
        Select Case msgBoxResponse
            Case vbYes
                WebHelpers.printMsg "NinjaAuthenticator.GetRefreshedToken", "Redirecting to Login..."
        
                ' Clear expired auth code and token
                Me.Logout
                
                ' Re-login and update the auth code
                Me.Login
                
                ' Clean up and redirect to the GetNewToken procedure
                Set auth_TokenClient = Nothing
                Set auth_Request = Nothing
                Set auth_Response = Nothing
                
                GetNewToken client
            Case vbNo
                ' Notify stopping refresh token process
                WebHelpers.printMsg "NinjaAuthenticator.GetRefreshedToken", "Abort refreshing token process"
        End Select
    Else
        ' Raise an error if the token refresh request is not successful
        Err.Raise 11041 + vbObjectError, "NinjaAuthenticator.GetRefreshedToken", _
            auth_Response.StatusCode & ": " & auth_Response.Content
    End If
    
auth_Cleanup:
    ' Clean up objects and handle errors
    Set auth_TokenClient = Nothing
    Set auth_Request = Nothing
    Set auth_Response = Nothing
    
    ' Rethrow error
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        
        auth_ErrorDescription = "An error occurred while retrieving token." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
    
        WebHelpers.LogError auth_ErrorDescription, "NinjaAuthenticator.GetNewToken", 11041 + vbObjectError
        Err.Raise 11041 + vbObjectError, "NinjaAuthenticator.GetNewToken", auth_ErrorDescription
    End If
End Sub

''
' Saves obtained access token and refresh token into an Outlook StorageItem.
'
' @method StoreTokens
'
' This function performs the following steps:
' 1. Checks if the folder for storing tokens exists. If not, creates a new folder.
' 2. Checks if the StorageItem (Outlook.StorageItem) for storing tokens exists. If not, creates a new StorageItem.
' 3. Adds/reuses an entry to the StorageItem PropertyAccessor and fills it with the token data.
' 4. Logs the success of the operation and handles any errors that occur.
''
Private Sub StoreTokens()
    On Error GoTo StoreTokens_Cleanup
    
    Dim TokenStore As Outlook.StorageItem
    Dim TokenUP As Outlook.UserProperties

    Dim inbox As Outlook.Folder
    Dim storeFolder As Outlook.Folder
    Dim ns As Outlook.NameSpace

    ' Get default inbox folder
    Set ns = Application.GetNamespace("MAPI")
    Set inbox = ns.GetDefaultFolder(olFolderInbox)

    ' Check if the folder for storing tokens exists
    If Not WebHelpers.olFolderExists(inbox, auth_TokensStorage) Then
        ' Create a new folder and set its name
        Set storeFolder = inbox.Folders.Add(auth_TokensStorage)
    Else
        ' Folder exists, set the folder variable
        Set storeFolder = inbox.Folders(auth_TokensStorage)
    End If
    
    ' Check if the StorageItem (Object) for storing tokens exists
    If Not WebHelpers.StorageItemExists(storeFolder, auth_StorageItemTokensStorage) Then
        Set TokenStore = storeFolder.GetStorage(auth_StorageItemTokensStorage, olIdentifyBySubject)
        Set TokenUP = TokenStore.UserProperties
        ' Create the StorageItem entries for storing tokens
        TokenUP.Add "Get_Datetime", olDateTime
        TokenUP.Add "Access_Token", olText
        TokenUP.Add "AT_Expired_Datetime", olDateTime
        TokenUP.Add "Refresh_Token", olText
        TokenUP.Add "RT_Expired_Datetime", olDateTime
        TokenUP.Add "Code_Verifier", olText
        
    Else
        ' StorageItem exists, set the object variable
        Set TokenStore = storeFolder.GetStorage(auth_StorageItemTokensStorage, olIdentifyBySubject)
        Set TokenUP = TokenStore.UserProperties
    End If
    
    ' Add a new row to the table and fill it with token data
    TokenUP.item("Get_Datetime").Value = Now
    TokenUP.item("Access_Token").Value = Me.AccessToken
    TokenUP.item("AT_Expired_Datetime").Value = Me.AccessTokenExpiration
    TokenUP.item("Refresh_Token").Value = Me.RefreshToken
    TokenUP.item("RT_Expired_Datetime").Value = Me.RefreshTokenExpiration
    TokenUP.item("Code_Verifier").Value = Me.CodeVerifier
                         
    ' Log success of storing the tokens
    WebHelpers.LogDebug "Stored existing tokens success, StorageItem: " & auth_StorageItemTokensStorage, "NinjaAuthenticator.StoreTokens"
    
StoreTokens_Cleanup:
    ' Cleanup object variables
    Set ns = Nothing
    Set inbox = Nothing
    Set storeFolder = Nothing
    Set TokenStore = Nothing
    Set TokenUP = Nothing
    
    ' Error handling block
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        
        ' Construct the error description message
        auth_ErrorDescription = "An error occurred while storing tokens into an StorageItem." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
        
        ' Log the error
        WebHelpers.LogError auth_ErrorDescription, "NinjaAuthenticator.StoreTokens", 11041 + vbObjectError
        ' Raise the error for further handling
        Err.Raise 11041 + vbObjectError, "NinjaAuthenticator.StoreTokens", auth_ErrorDescription
    End If
End Sub

''
' Loads saved access token and refresh token from an Excel sheet.
'
' @method LoadTokens
'
' This function performs the following steps:
' 1. Checks if the folder for storing tokens exists. If not, logs the failure and exits.
' 2. Checks if the StorageItem (Outlook.StorageItem) for storing tokens exists. If not, logs the failure and exits.
' 3. Checks if there are any entries in the StorageItem.UserProperties collection. If not, logs the failure and exits.
' 4. Loads the tokens into the auth object.
' 5. Checks if the access token has expired. If so, logs the failure and clears the token values.
' 6. Checks if the refresh token has expired. If so, logs the failure and clears the token values.
' 7. Handles any errors that occur and logs them.
''
Private Sub LoadTokens()
    On Error GoTo LoadTokens_Cleanup
    
    Dim TokenStore As Outlook.StorageItem
    Dim TokenUP As Outlook.UserProperties

    Dim inbox As Outlook.Folder
    Dim storeFolder As Outlook.Folder
    Dim ns As Outlook.NameSpace

    ' Get default inbox folder
    Set ns = Application.GetNamespace("MAPI")
    Set inbox = ns.GetDefaultFolder(olFolderInbox)

    ' Check if the folder for storing tokens exists
    If Not WebHelpers.olFolderExists(inbox, auth_TokensStorage) Then
        ' Log failure due to non-existent folder
        WebHelpers.LogDebug "Load tokens failed, folder: " & auth_TokensStorage & "does not exist", "NinjaAuthenticator.LoadTokens"
        ' Exit the subroutine
        Exit Sub
    Else
        ' Folder exists, set the folder variable
        Set storeFolder = inbox.Folders(auth_TokensStorage)
    End If
    
    ' Check if the StorageItem (Object) for storing tokens exists
    If Not WebHelpers.StorageItemExists(storeFolder, auth_StorageItemTokensStorage) Then
        ' Log failure due to non-existent StorageItem
        WebHelpers.LogDebug "Load tokens failed, StorageItem(Folder): " & auth_StorageItemTokensStorage & " does not exist", "NinjaAuthenticator.LoadTokens"
        ' Exit the subroutine
        Exit Sub
    Else
        ' StorageItem exists, set the storage variable
        Set TokenStore = storeFolder.GetStorage(auth_StorageItemTokensStorage, olIdentifyBySubject)
        Set TokenUP = TokenStore.UserProperties
    End If
    
    ' Check if there are any entries in the user properties
    If Not TokenUP.Count > 0 Then
        ' Log failure due to no entries in the collection
        WebHelpers.LogDebug "Load tokens failed, StorageItem(Folder) " & auth_StorageItemTokensStorage & " does not have any records", "NinjaAuthenticator.LoadTokens"
        ' Exit the subroutine
        Exit Sub
    End If
    
    ' Load tokens into the auth object
    Me.AccessToken = TokenUP.item("Access_Token").Value
    Me.AccessTokenExpiration = TokenUP.item("AT_Expired_Datetime").Value
    Me.RefreshToken = TokenUP.item("Refresh_Token").Value
    Me.RefreshTokenExpiration = TokenUP.item("RT_Expired_Datetime").Value
    Me.CodeVerifier = TokenUP.item("Code_Verifier").Value
    
    ' Check if the access token has expired
    If Me.AccessTokenExpiration < Now Then
        ' Log failure due to expired access token
        WebHelpers.LogDebug "Load AccessToken failed, AccessToken expired at: " & Format(Me.AccessTokenExpiration, "dd/mm/yyyy hh:mm"), "NinjaAuthenticator.LoadTokens"
        ' Clear the access token values
        Me.AccessToken = ""
        Me.AccessTokenExpiration = CDate(0)
    Else
        ' Log success
        WebHelpers.LogDebug "Load AccessToken success: " & Me.AccessToken, "NinjaAuthenticator.LoadTokens"
    End If
    
    ' Check if the refresh token has expired
    If Me.RefreshTokenExpiration < Now Then
        ' Log failure due to expired refresh token
        WebHelpers.LogDebug "Load RefreshToken failed, RefreshToken expired at: " & Format(Me.RefreshTokenExpiration, "dd/mm/yyyy hh:mm"), "NinjaAuthenticator.LoadTokens"
        ' Clear the refresh token values
        Me.RefreshToken = ""
        Me.RefreshTokenExpiration = CDate(0)
    Else
        ' Log success
        WebHelpers.LogDebug "Load RefreshToken success: " & Me.RefreshToken, "NinjaAuthenticator.LoadTokens"
    End If
        
LoadTokens_Cleanup:
    ' Cleanup object variables
    Set ns = Nothing
    Set inbox = Nothing
    Set storeFolder = Nothing
    Set TokenStore = Nothing
    Set TokenUP = Nothing
    
    ' Error handling block
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        
        ' Construct the error description message
        auth_ErrorDescription = "An error occurred while loading tokens from Excel sheet." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
    
        ' Log the error
        WebHelpers.LogError auth_ErrorDescription, "NinjaAuthenticator.LoadTokens", 11041 + vbObjectError
        ' Raise the error for further handling
        Err.Raise 11041 + vbObjectError, "NinjaAuthenticator.LoadTokens", auth_ErrorDescription
    End If
End Sub


''
' Clears/deletes all saved tokens from this Excel workbook.
'
' @method ClearTokens
'
' This function performs the following steps:
' 1. Checks if the folder from which to clear token information exists.
' 2. If the folder does not exist, notifies the user and exits.
' 3. If the folder is found, proceeds to check if the StorageItem (Folder) exists.
' 4. If the StorageItem does not exist, notifies the user and exits.
' 5. If the StorageItem is found, deletes the item and its contents.
' 6. Notifies the user of the result, performs cleanup, and handles errors.
'
' Note: This function uses the 'auth_TokensStorage' and 'auth_StorageItemTokensStorage' constants.
''
Private Sub ClearTokens()
    On Error GoTo ClearTokens_Cleanup
    
    Dim TokenStore As Outlook.StorageItem

    Dim inbox As Outlook.Folder
    Dim storeFolder As Outlook.Folder
    Dim ns As Outlook.NameSpace
    
    ' Get default inbox folder
    Set ns = Application.GetNamespace("MAPI")
    Set inbox = ns.GetDefaultFolder(olFolderInbox)
    
    ' Check if the folder from which to clear token information exists
    If Not WebHelpers.olFolderExists(inbox, auth_TokensStorage) Then
        ' Log failure due to non-existent sheet
        WebHelpers.LogDebug "Clear tokens cache failed, folder: " & auth_TokensStorage & "does not exist", "NinjaAuthenticator.ClearTokens"
        ' Notify the user
        MsgBox "Clearing tokens cache failed; tokens cache did not exist.", vbCritical + vbOKOnly, "NinjaRMM Ticket Connector - Microsoft Outlook"
        ' Exit the subroutine as the sheet does not exist
        Exit Sub
    Else
        ' Folder exists, set the folder variable
        Set storeFolder = inbox.Folders(auth_TokensStorage)
    End If
    
    ' Check if the StorageItem (Object) from which to clear token information exists
    If Not WebHelpers.StorageItemExists(storeFolder, auth_StorageItemTokensStorage) Then
        ' Log failure due to non-existent StorageItem
        WebHelpers.LogDebug "Clear tokens cache failed, StorageItem(Folder): " & auth_StorageItemTokensStorage & " does not exist", "NinjaAuthenticator.ClearTokens"
        ' Notify the user
        MsgBox "Clearing tokens cache failed; tokens cache did not exist.", vbCritical + vbOKOnly, "NinjaRMM Ticket Connector - Microsoft Outlook"
        ' Exit the subroutine
        Exit Sub
    Else
        ' StorageItem exists, set the storage variable
        Set TokenStore = storeFolder.GetStorage(auth_StorageItemTokensStorage, olIdentifyBySubject)
    End If
    
    ' Delete the StorageItem (Folder) and all of its user properties/records
    TokenStore.Delete
    
    ' Log success of clearing token information
    WebHelpers.LogDebug "Clearing tokens cache succeeded; all records of tokens were deleted.", "NinjaAuthenticator.ClearTokens"
    ' Notify the user of successful deletion
    MsgBox "Clearing tokens cache succeeded; all records of tokens were deleted.", vbInformation + vbOKOnly, "NinjaRMM Ticket Connector - Microsoft Outlook"

ClearTokens_Cleanup:
    ' Cleanup object variables
    Set ns = Nothing
    Set inbox = Nothing
    Set storeFolder = Nothing
    Set TokenStore = Nothing
    
    ' Error handling block
    If Err.Number <> 0 Then
        Dim auth_ErrorDescription As String
        
        ' Construct the error description message
        auth_ErrorDescription = "An error occurred while clearing tokens cache from Outlook folder." & vbNewLine
        If Err.Number - vbObjectError <> 11041 Then
            auth_ErrorDescription = auth_ErrorDescription & _
                Err.Number & VBA.IIf(Err.Number < 0, " (" & VBA.LCase$(VBA.Hex$(Err.Number)) & ")", "") & ": "
        End If
        auth_ErrorDescription = auth_ErrorDescription & Err.Description
        
        ' Log the error
        WebHelpers.LogError auth_ErrorDescription, "NinjaAuthenticator.ClearTokens", 11041 + vbObjectError
        ' Raise the error for further handling
        Err.Raise 11041 + vbObjectError, "NinjaAuthenticator.ClearTokens", auth_ErrorDescription
    End If
End Sub

Private Sub Class_Initialize()
    'Me.Scopes = Array("")
    LoadTokens
End Sub


