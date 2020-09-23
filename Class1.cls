VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim Surveille As Control
Public WithEvents PB As PictureBox
Attribute PB.VB_VarHelpID = -1



  
Public Property Set MousePos(ByRef MP As PictureBox)
'    Set Surveille = MP
'    SetCapture MP.hwnd
'    Set PB = MP
End Property

Public Sub LookFor(aControl As Variant)
'    Set Surveille = aControl
'    SetCapture PB.hwnd
End Sub

Public Sub ForceMouseOut()
    'ReleaseCapture
    'RaiseEvent MouseOut
End Sub

Private Sub PB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Dim Xi, Yi As Single
'    Dim xyCursor As POINTAPI
'    GetCursorPos xyCursor
'    ScreenToClient PB.hwnd, xyCursor
'
'    Xi = xyCursor.X * Screen.TwipsPerPixelX
'    Yi = xyCursor.Y * Screen.TwipsPerPixelY
'
'    If Xi < 0 Or Yi < 0 Or Xi > PB.Width Or Yi > PB.Height Then
'          ReleaseCapture
'          EventMouseOut
'          'RaiseEvent MouseOut
'    Else
'        EventMouseMove
'        'RaiseEvent MouseMove(Button, Shift, Xi, Yi)
'    End If
End Sub

Public Sub EventMouseMove()
    Form1.Picture3.BackColor = RGB(255, 0, 0)
End Sub

Public Sub EventMouseOut()
    Form1.Picture3.BackColor = Form1.BackColor
End Sub
