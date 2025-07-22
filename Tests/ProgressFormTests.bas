Attribute VB_Name = "ProgressFormTests"
Option Explicit

' ----------------------------------------------------------------------
'   Tests for ProgressForm
' ----------------------------------------------------------------------

' Shows the ProgressForm modelessly and checks if it is centered relative
' to the active Outlook window. Results are printed in the Immediate window.
Public Sub TestProgressFormPosition()
    Dim pf As ProgressForm
    Dim wnd As Object
    Dim expectedLeft As Double
    Dim expectedTop As Double

    Set pf = New ProgressForm
    pf.show vbModeless
    DoEvents

    Set wnd = Application.ActiveWindow
    expectedLeft = wnd.Left + (wnd.Width - pf.Width) / 2
    expectedTop = wnd.Top + (wnd.Height - pf.Height) / 2

    Debug.Print "ProgressForm.Left:", pf.Left
    Debug.Print "Expected Left:", expectedLeft
    Debug.Print "ProgressForm.Top:", pf.Top
    Debug.Print "Expected Top:", expectedTop

    If Abs(pf.Left - expectedLeft) <= 2 And Abs(pf.Top - expectedTop) <= 2 Then
        Debug.Print "ProgressForm position is centered correctly."
    Else
        Debug.Print "ProgressForm position is not centered as expected."
    End If

    Unload pf
End Sub
