VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   2805
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Icon"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3840
      Width           =   975
   End
   Begin VB.PictureBox picCursor 
      Height          =   375
      Left            =   3720
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   8
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      Height          =   3765
      Left            =   0
      ScaleHeight     =   3705
      ScaleWidth      =   2715
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   4
         Left            =   240
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":0152
         ScaleHeight     =   255
         ScaleWidth      =   2295
         TabIndex        =   7
         Top             =   2280
         Width           =   2295
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   3
         Left            =   240
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":0684
         ScaleHeight     =   255
         ScaleWidth      =   2295
         TabIndex        =   6
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   3240
         Width           =   2175
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":0BB6
         ScaleHeight     =   255
         ScaleWidth      =   2175
         TabIndex        =   4
         Top             =   2760
         Width           =   2175
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   2
         Left            =   240
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":10E8
         ScaleHeight     =   255
         ScaleWidth      =   2295
         TabIndex        =   3
         Top             =   1320
         Width           =   2295
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   1
         Left            =   240
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":161A
         ScaleHeight     =   255
         ScaleWidth      =   2295
         TabIndex        =   2
         Top             =   840
         Width           =   2295
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   0
         Left            =   240
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":1B4C
         ScaleHeight     =   255
         ScaleWidth      =   2295
         TabIndex        =   1
         Top             =   360
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" _
  (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long

Private LastIndex As Integer

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Enum MouseTekst
    Mouse0 = 0
    Mouse1 = 1
    Mouse2 = 2
    Mouse3 = 3
    Mouse4 = 4
End Enum

Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, _
qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean

Private Const BF_LEFT = &H1
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOM = &H8

Private Const BF_MIDDLE = &H800        ' Fill in the middle
Private Const BF_SOFT = &H1000         ' For softer buttons
Private Const BF_ADJUST = &H2000       ' Calculate the space left over
Private Const BF_FLAT = &H4000         ' For flat rather than 3D borders
Private Const BF_MONO = &H8000         ' For monochrome borders


Private Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Private Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Private Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Private Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Const BF_DIAGONAL = &H10

Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENOUTER = &H2
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_SUNKENINNER = &H8

Private Const BDR_OUTER = &H3
Private Const BDR_INNER = &HC
Private Const BDR_RAISED = &H5
Private Const BDR_SUNKEN = &HA

Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Function GetMouseTekst(ByVal MT As MouseTekst, ByVal InCtrl As Boolean) As String
    Select Case MT
        Case 0
            If InCtrl Then
                GetMouseTekst = "Open"
            Else
                GetMouseTekst = "Open"
            End If
        Case 1
            If InCtrl Then
                GetMouseTekst = "Paint"
            Else
                GetMouseTekst = "Paint"
            End If
        Case 2
            If InCtrl Then
                GetMouseTekst = "Spell"
            Else
                GetMouseTekst = "Spell"
            End If
        Case 3
            If InCtrl Then
                GetMouseTekst = "Print"
            Else
                GetMouseTekst = "Print"
            End If
        Case 4
            If InCtrl Then
                GetMouseTekst = "Redo"
            Else
                GetMouseTekst = "Redo"
            End If
    End Select
End Function

Private Sub Form_Load()
    Dim i As Byte
    For i = 0 To 4
        Me.Picture2(i).MouseIcon = Me.picCursor.Picture
        DrawCloseControl Me.Picture2(i), GetMouseTekst(i, False), BF_FLAT
    Next
End Sub

Private Sub Picture2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Xi, Yi As Single
    Dim xyCursor As POINTAPI
    SetCapture Me.Picture2(Index).hwnd
    GetCursorPos xyCursor
    ScreenToClient Me.Picture2(Index).hwnd, xyCursor

    Xi = xyCursor.X * Screen.TwipsPerPixelX
    Yi = xyCursor.Y * Screen.TwipsPerPixelY

    If Xi < 0 Or Yi < 0 Or Xi > Me.Picture2(Index).Width Or Yi > Me.Picture2(Index).Height Then
        LastIndex = -1
        ReleaseCapture
        DrawCloseControl Me.Picture2(Index), GetMouseTekst(Index, False), BF_FLAT
    Else
        If LastIndex = Index Then Exit Sub
        LastIndex = Index
        DrawCloseControl Me.Picture2(Index), GetMouseTekst(Index, True), EDGE_RAISED
    End If
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Xi, Yi As Single
    Dim xyCursor As POINTAPI
    SetCapture Me.Picture3.hwnd
    GetCursorPos xyCursor
    ScreenToClient Picture3.hwnd, xyCursor

    Xi = xyCursor.X * Screen.TwipsPerPixelX
    Yi = xyCursor.Y * Screen.TwipsPerPixelY

    If Xi < 0 Or Yi < 0 Or Xi > Picture3.Width Or Yi > Picture3.Height Then
          ReleaseCapture
          DrawCloseControl Me.Picture3, "Out of control", BF_FLAT
    Else
        DrawCloseControl Me.Picture3, "MouseDown", EDGE_RAISED
    End If
End Sub

Public Sub DrawCloseControl(picControl As PictureBox, _
    strCaption As String, Optional vntEdge, Optional ByVal ButWidth As Single = 250, Optional ByVal FontIsBold As Boolean = True, Optional StandFont As Boolean = True)
    On Error GoTo fout
    Dim r As RECT
    Dim intOffset%
    vntEdge = IIf(IsMissing(vntEdge), EDGE_RAISED, vntEdge)
    With picControl
        .AutoRedraw = True
        .Appearance = 0
        .ScaleMode = vbPixels
        .BorderStyle = 0
        .BackColor = vb3DFace
        .Cls
        .ForeColor = RGB(0, 0, 255)
        .CurrentX = 50
        .CurrentY = 2
        r.Left = .ScaleLeft
        r.Top = .ScaleTop
        If Me.Check1.Value = vbChecked Then
            r.Right = .ScaleWidth \ 7.5
        Else
            r.Right = .ScaleWidth
        End If
        r.Bottom = .ScaleHeight
        If vntEdge = EDGE_SUNKEN Then intOffset = 2
    End With
    picControl.Print strCaption
    DrawEdge picControl.hdc, r, CLng(vntEdge), BF_RECT
    If picControl.AutoRedraw Then picControl.Refresh
    Exit Sub
fout:
    MsgBox Err.Description
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Xi, Yi As Single
    Dim xyCursor As POINTAPI
    SetCapture Me.Text1.hwnd
    GetCursorPos xyCursor
    ScreenToClient Text1.hwnd, xyCursor

    Xi = xyCursor.X * Screen.TwipsPerPixelX
    Yi = xyCursor.Y * Screen.TwipsPerPixelY

    If Xi < 0 Or Yi < 0 Or Xi > Text1.Width Or Yi > Text1.Height Then
          ReleaseCapture
          Me.Text1.BackColor = RGB(255, 255, 255)
          Me.Text1.Text = "Out"
    Else
        Me.Text1.BackColor = RGB(100, 100, 255)
        Me.Text1.Text = "In"
    End If
End Sub
