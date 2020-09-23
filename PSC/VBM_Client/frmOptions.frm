VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "VBMessenger - Optionen"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6345
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   360
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   423
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox chkOption 
      Caption         =   "Nachrichten speichern"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   10
      Top             =   2760
      Value           =   1  'Aktiviert
      Width           =   6135
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2010
      Width           =   3015
   End
   Begin VB.CheckBox chkOption 
      Caption         =   "Blinkende Fenster aktivieren"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Value           =   1  'Aktiviert
      Width           =   6015
   End
   Begin VB.CheckBox chkOption 
      Caption         =   "Fenster nur blinken lassen, wenn es sich nicht im Vordergrund befindet"
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Value           =   1  'Aktiviert
      Width           =   6015
   End
   Begin VB.CheckBox chkOption 
      Caption         =   "Sounds nur abspielen, wenn sich das Fenster im Vordergrund befindet"
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Value           =   1  'Aktiviert
      Width           =   6015
   End
   Begin VB.CheckBox chkOption 
      Caption         =   "Sound abspielen"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Value           =   1  'Aktiviert
      Width           =   2055
   End
   Begin VBMessenger9.isButton cmdOk 
      Height          =   345
      Left            =   3240
      TabIndex        =   0
      Top             =   4920
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   609
      Style           =   7
      Caption         =   "Einstellungen übernehmen"
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
   Begin VBMessenger9.isButton cmdGetPath 
      Height          =   300
      Left            =   4920
      TabIndex        =   11
      Top             =   2010
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   529
      Style           =   7
      Caption         =   "..."
      IconSize        =   0
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
   Begin VBMessenger9.isButton cmdCancel 
      Height          =   345
      Left            =   480
      TabIndex        =   12
      Top             =   4920
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   609
      Style           =   7
      Caption         =   "Einstellungen verwerfen"
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
   Begin ComctlLib.Slider sldTrigger 
      Height          =   255
      Left            =   2400
      TabIndex        =   13
      Top             =   4320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   327682
      Max             =   99
      SelStart        =   10
      TickStyle       =   3
      Value           =   10
   End
   Begin ComctlLib.Slider sldPlayVolume 
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   3840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   327682
      Min             =   -10000
      Max             =   0
      TickStyle       =   3
   End
   Begin ComctlLib.Slider sldRecVolume 
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   4320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   327682
      Min             =   -10000
      Max             =   0
      TickStyle       =   3
   End
   Begin ComctlLib.Slider sldQuality 
      Height          =   255
      Left            =   2400
      TabIndex        =   20
      Top             =   3840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   327682
      Min             =   1
      Max             =   100
      SelStart        =   80
      TickStyle       =   3
      Value           =   80
   End
   Begin VB.Label Label3 
      Caption         =   " Qualität"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   21
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Aufnahme Lautstärke"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Wiedergabe Lautstärke"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label lblOptions 
      Caption         =   "Empfindlichkeit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   17
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label lblTopic 
      Caption         =   "Voice-Chat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A85E33&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   16
      Top             =   3360
      Width           =   6255
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      Index           =   7
      X1              =   0
      X2              =   424
      Y1              =   313
      Y2              =   313
   End
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      Index           =   6
      X1              =   0
      X2              =   424
      Y1              =   312
      Y2              =   312
   End
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      Index           =   5
      X1              =   0
      X2              =   424
      Y1              =   112
      Y2              =   112
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      Index           =   4
      X1              =   0
      X2              =   424
      Y1              =   113
      Y2              =   113
   End
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      Index           =   3
      X1              =   0
      X2              =   424
      Y1              =   160
      Y2              =   160
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   0
      X2              =   424
      Y1              =   161
      Y2              =   161
   End
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   0
      X2              =   424
      Y1              =   216
      Y2              =   216
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   0
      X2              =   424
      Y1              =   217
      Y2              =   217
   End
   Begin VB.Label Label4 
      Caption         =   "Eingehende Dateien"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblTopic 
      Caption         =   "Nachrichten History"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A85E33&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   6255
   End
   Begin VB.Label lblTopic 
      Caption         =   "Verzeichnisse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A85E33&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   6255
   End
   Begin VB.Label lblTopic 
      Caption         =   "Einkommende Nachrichten"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A85E33&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkOption_Click(Index As Integer)

    chkOption(1).Enabled = chkOption(0).Value = vbChecked
    chkOption(3).Enabled = chkOption(2).Value = vbChecked

End Sub

Private Sub cmdCancel_Click()

'~ VoiceChat wieder zurücksetzen

    RecordingVolume modDeclaration.SavedOptions.lVoiceRecVol
    VolumeOut modDeclaration.SavedOptions.lSoundVol
    Quality modDeclaration.SavedOptions.lQuality
    TriggerVoiceVolume modDeclaration.SavedOptions.lTriggerVal
    Unload Me

End Sub

Private Sub cmdGetPath_Click()

    txtPath.Text = modCommonControl.BrowseForFolder

End Sub

Private Sub cmdOk_Click()

'~ und ab damit...

    modDeclaration.SavedOptions.Sound = IIf(chkOption(0).Value = vbChecked, True, False)
    modDeclaration.SavedOptions.SoundLimited = IIf(chkOption(1).Value = vbChecked, True, False)
    modDeclaration.SavedOptions.Flash = IIf(chkOption(2).Value = vbChecked, True, False)
    modDeclaration.SavedOptions.FlashLimited = IIf(chkOption(3).Value = vbChecked, True, False)
    modDeclaration.SavedOptions.FileTransferPath = txtPath.Text
    modDeclaration.SavedOptions.SaveHistory = IIf(chkOption(4).Value = vbChecked, True, False)
    modDeclaration.SavedOptions.lQuality = sldQuality.Value
    modDeclaration.SavedOptions.lVoiceRecVol = sldRecVolume.Value
    modDeclaration.SavedOptions.lTriggerVal = sldTrigger.Value
    modDeclaration.SavedOptions.lSoundVol = sldPlayVolume.Value
    Unload Me

End Sub

Private Sub Form_Load()

    chkOption(0).Value = IIf(modDeclaration.SavedOptions.Sound, vbChecked, vbUnchecked)
    chkOption(1).Value = IIf(modDeclaration.SavedOptions.SoundLimited, vbChecked, vbUnchecked)
    chkOption(2).Value = IIf(modDeclaration.SavedOptions.Flash, vbChecked, vbUnchecked)
    chkOption(3).Value = IIf(modDeclaration.SavedOptions.FlashLimited, vbChecked, vbUnchecked)
    txtPath.Text = modDeclaration.SavedOptions.FileTransferPath
    chkOption(4).Value = IIf(modDeclaration.SavedOptions.SaveHistory, vbChecked, vbUnchecked)
    chkOption(1).Enabled = chkOption(0).Value = vbChecked
    chkOption(3).Enabled = chkOption(2).Value = vbChecked
    sldPlayVolume.Value = modDeclaration.SavedOptions.lSoundVol
    sldRecVolume.Value = modDeclaration.SavedOptions.lVoiceRecVol
    sldTrigger.Value = modDeclaration.SavedOptions.lTriggerVal
    sldQuality.Value = modDeclaration.SavedOptions.lQuality

End Sub

Private Sub Quality(Quality As Long)

Dim DVCC As DVCLIENTCONFIG

    On Error Resume Next
    DVCC = dvClient.GetClientConfig
    DVCC.lBufferQuality = Quality
    dvClient.SetClientConfig DVCC

End Sub

Private Sub RecordingVolume(Volume As Long)

Dim DVCC As DVCLIENTCONFIG

    On Error Resume Next
    DVCC = dvClient.GetClientConfig
    DVCC.lRecordVolume = Volume
    dvClient.SetClientConfig DVCC

End Sub

Private Sub sldPlayVolume_Change()

    sldPlayVolume_Scroll

End Sub

Private Sub sldPlayVolume_Scroll()

    VolumeOut sldPlayVolume.Value

End Sub

Private Sub sldQuality_Change()

    sldQuality_Scroll

End Sub

Private Sub sldQuality_Scroll()

    Quality sldQuality.Value

End Sub

Private Sub sldRecVolume_Change()

    sldRecVolume_Scroll

End Sub

Private Sub sldRecVolume_Scroll()

    RecordingVolume sldRecVolume.Value

End Sub

Private Sub sldTrigger_Change()

    sldTrigger_Scroll

End Sub

Private Sub sldTrigger_Scroll()

    TriggerVoiceVolume sldTrigger.Value

End Sub

Private Sub TriggerVoiceVolume(Volume As Long)

Dim DVCC As DVCLIENTCONFIG

    On Error Resume Next
    DVCC = dvClient.GetClientConfig
    DVCC.lThreshold = Volume
    dvClient.SetClientConfig DVCC

End Sub

Private Sub VolumeOut(Volume As Long)

Dim DVCC As DVCLIENTCONFIG

    On Error Resume Next
    DVCC = dvClient.GetClientConfig
    DVCC.lPlaybackVolume = Volume
    dvClient.SetClientConfig DVCC

End Sub


