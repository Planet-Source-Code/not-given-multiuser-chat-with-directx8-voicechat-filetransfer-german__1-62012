VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmReceiveFile 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Datei empfangen"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5145
   Icon            =   "frmReceiveFile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   165
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   343
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Timer tmrTimeOut 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   4680
      Top             =   2040
   End
   Begin ComctlLib.ProgressBar ProgressBar 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Timer tmrSpeed 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5520
      Top             =   2040
   End
   Begin Projekt1.isButton cmdCancel 
      Default         =   -1  'True
      Height          =   345
      Left            =   1560
      TabIndex        =   1
      Top             =   2040
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   609
      Style           =   7
      Caption         =   "Abbrechen"
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
   Begin VB.Label lblInfo 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "??? MB"
      Height          =   255
      Index           =   7
      Left            =   1800
      TabIndex        =   2
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label lblInfo 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "??? kb/s"
      Height          =   255
      Index           =   6
      Left            =   1800
      TabIndex        =   3
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label lblInfo 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "??? MB"
      Height          =   255
      Index           =   5
      Left            =   1800
      TabIndex        =   4
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label lblInfo 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "???"
      Height          =   255
      Index           =   4
      Left            =   1800
      TabIndex        =   5
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label lblInfo 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Speicherort"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblInfo 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Größe:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblInfo 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Geschwindigkeit"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblInfo 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Empfangen"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1695
   End
End
Attribute VB_Name = "frmReceiveFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' für den Empfangsvorgang
Dim nBytesTotal As Long
Dim nBytesRead As Long
Dim nFile As Integer

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
frmReceiveFile.Caption = "Datei empfangen (" & modDeclaration.FilesReceiving.FileSenderReceiver & ")"
tmrTimeOut.Enabled = True
End Sub

Public Sub DataArrival(Data As String)
' Wenn Daten ankommen...

tmrTimeOut.Enabled = False
tmrTimeOut.Enabled = True

  Dim sData As String
  Dim sTemp As String
  Static sFile As String
  
  sData = Right$(Data, Len(Data) - InStr(Data, "|"))
  
  If Left$(sData, 12) = "<begin size=" Then
    ' Aha... eine neue Datei wird gesendet
    tmrSpeed.Enabled = True
    sData = Mid$(sData, 13)
    sTemp = Left$(sData, InStr(sData, ">") - 1)
    sData = Mid$(sData, InStr(sData, ">") + 1)
    
    ' Dateigröße und Dateiname ermitteln
    If InStr(sTemp, ";") > 0 Then
      nBytesTotal = Val(Left$(sTemp, InStr(sTemp, ";") - 1))
      sFile = Mid$(sTemp, InStr(sTemp, ";") + 1)
    Else
      nBytesTotal = Val(sTemp)
    End If
    
    sFile = modDeclaration.FilesReceiving.SaveToPath & "\" & sFile
    
    lblInfo(4).Caption = modFunctions.CompactPath(frmReceiveFile, sFile, lblInfo(4))
    lblInfo(5).Caption = modFunctions.BytesToString(nBytesTotal)
    
    ' ggf. Datei löschen, falls bereits existiert
    On Error Resume Next
    Kill sFile
    On Error GoTo 0
    
    ' Datei im Binary-Mode öffnen
    nFile = FreeFile
    Open sFile For Binary As #nFile
    
    ' bisher gelesene Bytes zurücksetzen
    nBytesRead = 0
  End If
  
  If Len(sData) > 0 And nFile > 0 Then
    ' bisher empfangene Daten...
    nBytesRead = nBytesRead + Len(sData)
    
    ' Daten in Datei speichern
    Put #nFile, , sData
    
    ' evtl. Fortschritt anzeigen
   lblInfo(7).Caption = modFunctions.BytesToString(nBytesRead)
   
   ProgressBar.Value = Int(nBytesRead / nBytesTotal * 100)
    
    ' Wenn alle Bytes empfangen wurden, Datei schließen
    If nBytesRead = nBytesTotal Then
      Close #nFile
      nFile = 0
        tmrSpeed.Enabled = False
        MsgBox "Die Datei wurde vollständig empfangen!", vbInformation
        Unload Me
        
    End If
  End If

    DoEvents
End Sub

Private Sub Form_Unload(Cancel As Integer)
Close #nFile
modDeclaration.FilesReceiving.SendingReceiving = False
Set frmReceiveFile = Nothing
End Sub

Private Sub tmrSpeed_Timer()
Static TimeCounter As Long
TimeCounter = TimeCounter + 1 'Anzahl der sekunden
lblInfo(6).Caption = CStr(Round((nBytesRead / 1024) / TimeCounter, 2)) & " kb/sec"
End Sub

Private Sub tmrTimeOut_Timer()
MsgBox "Der Dateitransfer konnte nicht beendet werden!", vbExclamation, "Timeout"
nFile = 0
tmrSpeed.Enabled = False
Unload Me
End Sub
