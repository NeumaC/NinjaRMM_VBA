Attribute VB_Name = "ProgressHelper"
Option Explicit

Private mTotal As Long
Private mCurrent As Long
Private mProgressForm As ProgressForm

Public Sub ProgressStart(ByVal TotalItems As Long)
    mTotal = TotalItems
    mCurrent = 0
    Set mProgressForm = New ProgressForm
    mProgressForm.UpdateProgress 0, mTotal
    mProgressForm.UpdateStatus ""
    mProgressForm.show vbModeless
End Sub

Public Sub ProgressUpdate(ByVal CurrentItem As Long, ByVal Message As String)
    If mProgressForm Is Nothing Then Exit Sub
    mProgressForm.UpdateStatus Message
    mProgressForm.UpdateProgress CurrentItem, mTotal
End Sub

Public Sub ProgressStep(ByVal Message As String)
    mCurrent = mCurrent + 1
    ProgressUpdate mCurrent, Message
End Sub

Public Sub ProgressEnd()
    If Not mProgressForm Is Nothing Then
        Unload mProgressForm
        Set mProgressForm = Nothing
    End If
    mCurrent = 0
End Sub
