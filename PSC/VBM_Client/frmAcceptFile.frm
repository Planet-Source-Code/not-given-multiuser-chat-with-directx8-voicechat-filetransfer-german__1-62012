VERSION 5.00
Begin VB.Form frmAcceptFile 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Dateitransfer annehmen"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4305
   Icon            =   "frmAcceptFile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   108
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   287
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Timer tmrAutoDeny 
      Interval        =   1000
      Left            =   3480
      Top             =   1680
   End
   Begin VBMessenger9.isButton cmdYes 
      Height          =   300
      Index           =   0
      Left            =   2160
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   529
      Style           =   7
      Caption         =   "Speichern unter..."
      iNonThemeStyle  =   0
      USeCustomColors =   -1  'True
      BackColor       =   14737632
      HighlightColor  =   14737632
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      RoundedBordersByTheme=   0   'False
   End
   Begin VBMessenger9.isButton cmdNo 
      Height          =   300
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   529
      Style           =   7
      Caption         =   "Ablehnen (60)"
      iNonThemeStyle  =   0
      USeCustomColors =   -1  'True
      BackColor       =   14737632
      HighlightColor  =   14737632
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      RoundedBordersByTheme=   0   'False
   End
   Begin VBMessenger9.isButton cmdYes 
      Height          =   300
      Index           =   1
      Left            =   2160
      TabIndex        =   3
      Top             =   840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   529
      Style           =   7
      Caption         =   "Annehmen"
      iNonThemeStyle  =   0
      USeCustomColors =   -1  'True
      BackColor       =   14737632
      HighlightColor  =   14737632
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   1
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      RoundedBordersByTheme=   0   'False
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblAsk 
      BackStyle       =   0  'Transparent
      Caption         =   "Möchten Sie die Datei [xyz] (xykb) von Benutzer xyzxyz annehmen?"
      Height          =   735
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmAcceptFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public AcceptString   As String
Public DenyString     As String


Private Sub cmdNo_Click()

    frmMain.SendData DenyString
    modDeclaration.SendingOrReceivingFile = False
    Unload Me

End Sub

Private Sub cmdYes_Click(Index As Integer)

    tmrAutoDeny.Enabled = False
    If Index = 0 Then
        modDeclaration.PathOfFileToSendOrReceive = modCommonControl.BrowseForFolder
    Else 'NOT INDEX...
        modDeclaration.PathOfFileToSendOrReceive = modDeclaration.SavedOptions.FileTransferPath
    End If
    frmMain.SendData AcceptString
    Unload Me

End Sub

Private Sub Form_Load()

    modFunctions.MakeTopMost hwnd
    cmdYes(1).Enabled = modFunctions.DirExists(modDeclaration.SavedOptions.FileTransferPath)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)

    If Not UnloadMode = 1 Then
        cmdNo_Click
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmAcceptFile = Nothing

End Sub

Private Sub tmrAutoDeny_Timer()

Static counter As Integer

    counter = counter + 1
    cmdNo.Caption = "Ablehnen (" & CStr(60 - counter) & ")"
    If counter = 60 Then
        cmdNo_Click
        Unload frmAcceptFile
    End If

End Sub



