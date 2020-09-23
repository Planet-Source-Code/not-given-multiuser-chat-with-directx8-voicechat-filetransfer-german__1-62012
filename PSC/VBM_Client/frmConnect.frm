VERSION 5.00
Begin VB.Form frmConnect 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "VBM9 Verbinden..."
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3540
   Icon            =   "frmConnect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   3540
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox txtUserInfo 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   1440
      TabIndex        =   4
      Top             =   1035
      Width           =   1935
   End
   Begin VB.TextBox txtUserInfo 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox txtUserInfo 
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VBMessenger9.isButton cmdConnect 
      Height          =   300
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      Style           =   7
      Caption         =   "Verbinden"
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
   Begin VBMessenger9.isButton cmdEnd 
      Height          =   300
      Left            =   1800
      TabIndex        =   7
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      Style           =   7
      Caption         =   "Beenden"
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
   Begin VB.Label lblUserInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Serveradresse / InternetAdresse"
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label lblUserInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Passwort:"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblUserInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Benutzername:"
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConnect_Click()

'~ IP-Adresse herausfinden

    If LCase$(Left$(txtUserInfo(2).Text, 4)) = "http" Then
        modDeclaration.ServerIP = modFunctions.GetURL(txtUserInfo(2).Text)
    Else 'NOT LCASE$(LEFT$(TXTUSERINFO(2).TEXT,...
        modDeclaration.ServerIP = txtUserInfo(2).Text
    End If
    '~ Adresse validieren
    If Not modFunctions.IsValidIP(modDeclaration.ServerIP) Then
        Exit Sub
    Else 'NOT NOT...
        modDeclaration.Username = txtUserInfo(0)
        modDeclaration.UserPass = txtUserInfo(1)
    End If
    With frmMain
        .tmrConnect.Enabled = True
        .wsc.Close
        Do Until .wsc.State = 0
            DoEvents
        Loop
    End With 'frmMain
    modDeclaration.SignIn = True
    Unload Me

End Sub

Private Sub cmdEnd_Click()

    Unload Me

End Sub

Private Sub Form_Load()

'~ Benutzername, Passwort, Serveradresse laden / entschlüsseln

    txtUserInfo(0).Text = modFunctions.Rot13(GetSetting(App.Title, "UserData", "Name", vbNullString))
    txtUserInfo(1).Text = modFunctions.Rot13(GetSetting(App.Title, "UserData", "Pass", vbNullString))
    txtUserInfo(2).Text = modFunctions.Rot13(GetSetting(App.Title, "UserData", "Serv", vbNullString))

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)

'~ Benutzername, Passwort, Serveradresse speichern / verschlüsseln

    With App
        SaveSetting .Title, "UserData", "Name", modFunctions.Rot13(txtUserInfo(0).Text)
        SaveSetting .Title, "UserData", "Pass", modFunctions.Rot13(txtUserInfo(1).Text)
        SaveSetting .Title, "UserData", "Serv", modFunctions.Rot13(txtUserInfo(2).Text)
    End With 'App

End Sub



