Attribute VB_Name = "MouseOver"
Option Explicit

 ' Function found in the API Viewer
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

 ' Type found in the API Viewer and referenced by the GetCursorPos ^
Public Type POINTAPI
        X As Long
        Y As Long
End Type

Function MouseMove(FormName As Form, ControlName As Control, ControlLeft As Integer, ControlRight As Integer, ControlTop As Integer, ControlBottom As Integer)
    
     ' Return the mouse position if the screen
    Dim MousePos As POINTAPI
    Dim RetValue As Boolean
    RetValue = GetCursorPos(MousePos)
    
     ' Convert from Twips to Pixels
    Dim frmX, frmY
    frmX = MousePos.X - FormName.ScaleX(FormName.Left, vbTwips, vbPixels)
    frmY = MousePos.Y - FormName.ScaleY(FormName.Top, vbTwips, vbPixels)
    
     ' Light up the control
    If frmX < ControlLeft Then
        ControlName.BackColor = &H8000& ' This can be changed to represent the function you want
    ElseIf frmX > ControlRight Then     ' for example, Instead of the Controls Background color,
        ControlName.BackColor = &H8000& ' You might want to change the Controls picture.
    ElseIf frmY < ControlTop Then
        ControlName.BackColor = &H8000&
    ElseIf frmY > ControlBottom Then
        ControlName.BackColor = &H8000&
    Else
        ControlName.BackColor = &HFF00&
    End If
    
End Function

Function ReturnPixels(FormName As Form, Label As Label)

     ' Return the mouse position if the screen
    Dim MousePos As POINTAPI
    Dim RetValue As Boolean
    RetValue = GetCursorPos(MousePos)
    
     ' Convert from Twips to Pixels
    Dim frmX, frmY
    frmX = MousePos.X - FormName.ScaleX(FormName.Left, vbTwips, vbPixels)
    frmY = MousePos.Y - FormName.ScaleY(FormName.Top, vbTwips, vbPixels)
    
     ' Show the X, Y values in pixel on the form to figure _
     out the the value of the intergers while developing
    Label.Caption = "X: " & frmX & ", Y: " & frmY
    
End Function
