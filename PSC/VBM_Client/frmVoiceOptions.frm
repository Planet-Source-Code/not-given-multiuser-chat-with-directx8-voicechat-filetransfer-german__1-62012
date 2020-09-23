VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmVoiceOptions 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "VBMessenger VoiceChat"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   3135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin ComctlLib.Slider sldTrigger 
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   2040
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      _Version        =   327682
      Max             =   99
      SelStart        =   10
      TickStyle       =   3
      Value           =   10
   End
   Begin ComctlLib.Slider sldPlayVolume 
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      _Version        =   327682
      Min             =   -10000
      Max             =   0
      TickStyle       =   3
   End
   Begin ComctlLib.Slider sldRecVolume 
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      _Version        =   327682
      Min             =   -10000
      Max             =   0
      TickStyle       =   3
   End
   Begin VB.Label Label3 
      Caption         =   "Einstellungen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1455
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
      Left            =   240
      TabIndex        =   5
      Top             =   1800
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
      TabIndex        =   4
      Top             =   480
      Width           =   2175
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
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
   End
End
Attribute VB_Name = "frmVoiceOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
sldPlayVolume.Value = modDeclaration.SavedOptions.lSoundVol
sldRecVolume.Value = modDeclaration.SavedOptions.lVoiceRecVol
sldTrigger.Value = modDeclaration.SavedOptions.lTriggerVal
End Sub

Private Sub sldPlayVolume_Change()
sldPlayVolume_Scroll
End Sub

Private Sub sldPlayVolume_Scroll()
VolumeOut sldPlayVolume.Value
modDeclaration.SavedOptions.lSoundVol = sldPlayVolume.Value
End Sub

Private Sub sldRecVolume_Change()
sldRecVolume_Scroll
End Sub

Private Sub sldRecVolume_Scroll()
RecordingVolume sldRecVolume.Value
modDeclaration.SavedOptions.lVoiceRecVol = sldRecVolume.Value
End Sub

Private Sub sldTrigger_Change()
sldTrigger_Scroll
End Sub

Private Sub sldTrigger_Scroll()
TriggerVoiceVolume sldTrigger.Value
modDeclaration.SavedOptions.lTriggerVal = sldTrigger.Value
End Sub


Private Sub VolumeOut(Volume As Long)
On Error Resume Next

Dim DVCC As DVCLIENTCONFIG
DVCC = dvClient.GetClientConfig
DVCC.lPlaybackVolume = Volume

dvClient.SetClientConfig DVCC

End Sub

Private Sub RecordingVolume(Volume As Long)
On Error Resume Next

Dim DVCC As DVCLIENTCONFIG
DVCC = dvClient.GetClientConfig
DVCC.lRecordVolume = Volume
dvClient.SetClientConfig DVCC

End Sub

Private Sub TriggerVoiceVolume(Volume As Long)
On Error Resume Next

Dim DVCC As DVCLIENTCONFIG
DVCC = dvClient.GetClientConfig
DVCC.lThreshold = Volume
dvClient.SetClientConfig DVCC

End Sub
