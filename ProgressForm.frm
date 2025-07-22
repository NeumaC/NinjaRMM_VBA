VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressForm 
   Caption         =   "ProgressForm"
   ClientHeight    =   360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3000
   OleObjectBlob   =   "ProgressForm.frx":0000
End
Attribute VB_Name = "ProgressForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare PtrSafe Function FindWindowA Lib "user32" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" _
    (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr

Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" _
    (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr

Private Declare PtrSafe Function DrawMenuBar Lib "user32" _
    (ByVal hWnd As LongPtr) As Long

Private Declare PtrSafe Function GetSystemMetrics Lib "user32" _
    (ByVal nIndex As Long) As Long

Private Const GWL_STYLE As Long = -16
Private Const WS_CAPTION As LongPtr = &HC00000
Private Const WS_THICKFRAME As LongPtr = &H40000
Private Const WS_BORDER As LongPtr = &H800000

Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1

Private Sub UserForm_Activate()
    Dim hWnd As LongPtr
    Dim lStyle As LongPtr

    ' Titelleiste/Rahmen entfernen
    hWnd = FindWindowA(vbNullString, Me.Caption)
    If hWnd <> 0 Then
        lStyle = GetWindowLongPtr(hWnd, GWL_STYLE)
        lStyle = lStyle And Not WS_CAPTION
        lStyle = lStyle And Not WS_THICKFRAME
        lStyle = lStyle And Not WS_BORDER
        SetWindowLongPtr hWnd, GWL_STYLE, lStyle
        DrawMenuBar hWnd
    End If
End Sub

Private Sub UserForm_Initialize()
    lblProgress.Width = 0
End Sub

Public Sub UpdateStatus(ByVal Text As String)
    lblStatus.Caption = Text
    DoEvents
End Sub

Public Sub UpdateProgress(ByVal Value As Long, ByVal MaxValue As Long)
    Dim pct As Double
    If MaxValue > 0 Then
        pct = Value / MaxValue
    Else
        pct = 0
    End If
    lblProgress.Width = lblBarBack.Width * pct
    DoEvents
End Sub
