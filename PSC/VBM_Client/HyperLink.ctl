VERSION 5.00
Begin VB.UserControl HyperLink 
   AutoRedraw      =   -1  'True
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1950
   ScaleHeight     =   375
   ScaleWidth      =   1950
   ToolboxBitmap   =   "HyperLink.ctx":0000
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'Kein
      FillColor       =   &H00FF0000&
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   1935
      TabIndex        =   0
      Top             =   0
      Width           =   1935
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   45
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2400
      Top             =   3480
   End
End
Attribute VB_Name = "HyperLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type POINTAPI
    X                       As Long
    Y                       As Long
End Type
Private Type RECT
    Left                    As Long
    Top                     As Long
    Right                   As Long
    Bottom                  As Long
End Type
Public Event OnClick()
Private MouseIn         As Boolean
Private strCaption      As String
Private BolFontBold     As Boolean
Private olecBackColor   As OLE_COLOR
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, _
                                                ByVal X As Long, _
                                                ByVal Y As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long

Public Property Get BackColor() As OLE_COLOR

    BackColor = olecBackColor

End Property

Public Property Let BackColor(temp As OLE_COLOR)

    olecBackColor = temp
    UserControl.BackColor = temp
    Picture1.BackColor = temp

End Property

Public Property Get Caption() As String

    strCaption = Label1.Caption
    Caption = strCaption

End Property

Public Property Let Caption(ByVal strNewCaption As String)

    strCaption = strNewCaption
    Label1.Caption = strCaption
    PropertyChanged "Caption"
    Resize

End Property

Public Property Get FontBold() As Boolean

    BolFontBold = Label1.FontBold
    '''''''''''''''''''''''''''''''''''''''''''''' Let & Get
    FontBold = BolFontBold

End Property

Public Property Let FontBold(ByVal bolNewFontBold As Boolean)

    BolFontBold = bolNewFontBold
    Label1.FontBold = BolFontBold '''
    PropertyChanged "FontBold"

End Property

Private Sub Label1_Change()

    Resize

End Sub

Private Sub Label1_Click()

    RaiseEvent OnClick

End Sub

Private Sub Label1_MouseDown(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)

    Label1.ForeColor = vbBlack

End Sub

Private Sub Label1_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)

    If Not MouseIn Then
        Timer1.Enabled = True
        MouseIn = True
        ' MouseIn
        MouseComeIn
    End If

End Sub

Private Sub Label1_MouseUp(Button As Integer, _
                           Shift As Integer, _
                           X As Single, _
                           Y As Single)

    Label1.ForeColor = vbBlue

End Sub

Private Sub MouseComeIn()

    Label1.FontUnderline = True

End Sub

Private Sub MouseComeOut()

    Label1.FontUnderline = False

End Sub

Private Sub Resize()

    UserControl.Width = Label1.Width + Label1.Left
    UserControl.Height = Label1.Height
    With Picture1
        .Height = Label1.Height
        .Width = Label1.Width + Label1.Left
        .Top = 0
        .Left = 0
    End With 'Picture1
    Label1.Left = 0
    Label1.Top = 0

End Sub

Private Sub Timer1_Timer()

Dim P As POINTAPI
Dim R As RECT

    ' akuelle Mauszeiger-Position ermitteln
    GetCursorPos P
    ' Position und Größe der Form
    GetWindowRect Picture1.hwnd, R
    ' Befindet sich der Mauszeiger innerhalb der Form?
    If PtInRect(R, P.X, P.Y) = 0 Then
        Timer1.Enabled = False
        ' außerhalb
        MouseIn = False
        MouseComeOut
    End If

End Sub

Private Sub UserControl_Initialize()

    Resize

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    On Error Resume Next
    With PropBag
        BolFontBold = .ReadProperty("FontBold")
        Label1.FontBold = BolFontBold
        strCaption = .ReadProperty("Caption")
        BackColor = .ReadProperty("BackColor", &HFFFFFF)
    End With 'PropBag
    Label1.Caption = strCaption
    '''''''''''''''''''''''''''''''''''''''''' Read & Write

End Sub

Private Sub UserControl_Resize()

    Resize

End Sub

Private Sub UserControl_Terminate()

    Timer1.Enabled = False

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        .WriteProperty "Caption", strCaption
        .WriteProperty "FontBold", BolFontBold
        .WriteProperty "BackColor", olecBackColor, 0
    End With 'PropBag

End Sub

''
''Public Sub SetBackColour(color As ColorConstants)
''
''
''
''Picture1.BackColor = color
''UserControl.BackColor = color
''End Sub
''



