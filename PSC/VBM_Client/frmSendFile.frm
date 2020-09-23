VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmSendFile 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Datei senden"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5145
   Icon            =   "frmSendFile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   165
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   343
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Timer start 
      Interval        =   1000
      Left            =   5160
      Top             =   2040
   End
   Begin VB.Timer tmrSpeed 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5640
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
   Begin Projekt1.isButton cmdCancel 
      Default         =   -1  'True
      Height          =   345
      Left            =   1560
      TabIndex        =   9
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
      Caption         =   "??? MB"
      Height          =   255
      Index           =   7
      Left            =   1800
      TabIndex        =   8
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label lblInfo 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "??? kb/s"
      Height          =   255
      Index           =   6
      Left            =   1800
      TabIndex        =   7
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label lblInfo 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "??? MB"
      Height          =   255
      Index           =   5
      Left            =   1800
      TabIndex        =   6
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label lblInfo 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Gesendet:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblInfo 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Geschwindigkeit"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblInfo 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Größe:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblInfo 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Dateiname:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmSendFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim nFilePos As Long

Public Sub WinsockSendBinaryFile(ByVal sFile As String)
  Dim F As Integer
  Dim sBuffer As String
  Dim nFileSize As Long
  Dim nBytesToRead As Long
  Dim i As Long
  
  ' Größe der einzelnen Datenpakete
  Const BlockSize = 1000
  
  tmrSpeed.Enabled = True
  
  ' Datei im Binary-Mode öffnen
  F = FreeFile
  Open sFile For Binary As #F
  
  ' Dateiname extrahieren
  If InStr(sFile, "\") > 0 Then
    sFile = Mid$(sFile, InStrRev(sFile, "\") + 1)
  End If
  
  ' Dateigröße
  nFileSize = LOF(F)
  
  ' Sendevorgang starten

    ' Empfänger mitteln, welche Datei und wieviele
    ' Daten gesendet werden
    frmMain.SendData "t|" & modDeclaration.FilesSending.FileSenderReceiver & "|<begin size=" & CStr(nFileSize) & ";" & sFile & ">"
    
    ' Datei blockweise senden
    Do While nFilePos < nFileSize
      nBytesToRead = BlockSize
      If nFilePos + nBytesToRead > nFileSize Then
        nBytesToRead = nFileSize - nFilePos
      End If
      
      ' Datenblock lesen
      sBuffer = Space$(nBytesToRead)
      Get #F, , sBuffer
      
      ' Datenblock senden
      frmMain.SendData "t|" & modDeclaration.FilesSending.FileSenderReceiver & "|" & sBuffer
      
      
      
      ' Fortschritt aktualisieren
      nFilePos = nFilePos + nBytesToRead
      lblInfo(7).Caption = modFunctions.BytesToString(nFilePos)
      ProgressBar.Value = Int(nFilePos / nFileSize * 100)
      
    ' Wichtig!
    For i = 0 To 300
    DoEvents ' --> kleine Bremse, damit die anderen Benutzer sich auch noch unterhalten können ;)
    Next i
    
    If cmdCancel.Enabled = False Then Exit Do
      
    Loop
    
    modDeclaration.FilesSending.SendingReceiving = False

  
  ' Datei schließen (Sendevorgang beendet)
  Close #F
  
  tmrSpeed.Enabled = False
  MsgBox "Die Datei wurde vollständig gesendet!", vbInformation
  Unload frmSendFile
End Sub



Private Sub cmdCancel_Click()
cmdCancel.Enabled = False
End Sub

Private Sub Form_Load()
frmSendFile.Caption = "Datei senden (" & modDeclaration.FilesSending.FileSenderReceiver & ")"
lblInfo(4).Caption = modFunctions.ExtractFilename(modDeclaration.FilesSending.Filename)
lblInfo(5).Caption = modFunctions.GetFileLen(modDeclaration.FilesSending.Filename)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
cmdCancel.Enabled = False
Cancel = 1
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmSendFile = Nothing
End Sub

Private Sub start_Timer()
WinsockSendBinaryFile modDeclaration.FilesSending.Filename
start.Enabled = False
End Sub

Private Sub tmrSpeed_Timer()
Static TimeCounter As Long
TimeCounter = TimeCounter + 1 'Anzahl der sekunden
lblInfo(6).Caption = CStr(Round((nFilePos / 1024) / TimeCounter, 2)) & " kb/sec"
End Sub

