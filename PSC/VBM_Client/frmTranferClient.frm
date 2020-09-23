VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmTransfer 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "()"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5880
   Icon            =   "frmTranferClient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer tmrSuccess 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   240
      Top             =   4200
   End
   Begin VB.Timer tmrEnd 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   720
      Top             =   3240
   End
   Begin VB.Timer tmrConnectionCheck 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   240
      Top             =   3720
   End
   Begin VB.Timer tmrSpeed 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   720
      Top             =   3720
   End
   Begin ComctlLib.ProgressBar ProgressBar 
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSWinsockLib.Winsock wscFile 
      Left            =   1200
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   91
   End
   Begin VB.Timer tmrStart 
      Interval        =   1
      Left            =   240
      Top             =   3240
   End
   Begin VBMessenger9.isButton cmdCancel 
      Default         =   -1  'True
      Height          =   345
      Left            =   1920
      TabIndex        =   9
      Top             =   2160
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
      BackStyle       =   0  'Transparent
      Caption         =   "Übertragen"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Geschwindigkeit"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Größe"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Speicherort"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "???"
      Height          =   255
      Index           =   4
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "??? MB"
      Height          =   255
      Index           =   5
      Left            =   1800
      TabIndex        =   2
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "??? kb/s"
      Height          =   255
      Index           =   6
      Left            =   1800
      TabIndex        =   1
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "??? MB"
      Height          =   255
      Index           =   7
      Left            =   1800
      TabIndex        =   0
      Top             =   1200
      Width           =   3975
   End
End
Attribute VB_Name = "frmTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Listening As Boolean
Private Sending As Boolean
'~-------
Private nFilePos As Long
Private nBytesTotal As Long
Private nBytesRead As Long
Private nFile As Integer
Private FileFullyReceivedorSend As Boolean
Private StopLoop As Boolean

Private Sub cmdCancel_Click()

    UnloadfrmTransfer

End Sub

Private Sub Form_Load()

    bLoadedfrmTransfer = True
    modDeclaration.SendingOrReceivingFile = True
    If SendFile Then
        Caption = "Sende Datei an " & modDeclaration.ReceiverOrSender
    Else 'SENDFILE = FALSE/0
        Caption = "Empfange Datei von " & modDeclaration.ReceiverOrSender
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)

    If Not UnloadMode = 1 Then
        Cancel = 1
        cmdCancel_Click
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    modDeclaration.SendingOrReceivingFile = False
    frmMain.cmdSendFile.Value = False
    bLoadedfrmTransfer = False

End Sub

'~-------
Private Sub tmrConnectionCheck_Timer()

' checkt, ob die Verbindung noch steht

    If wscFile.State = 0 Or wscFile.State = 8 Then
        cmdCancel_Click
    End If

End Sub

' ##################################################################
' ~~~ Connecten!!!
' ##################################################################
Private Sub tmrEnd_Timer()

' beendet alles
'~ Timer & Winsock stoppen

Dim i As Integer

    tmrConnectionCheck.Enabled = False
    tmrSpeed.Enabled = False
    tmrStart.Enabled = False
    tmrSuccess.Enabled = False
    Close #nFile
    wscFile.Close
    ' ----
    If SendFile Then
        If FileFullyReceivedorSend Then
            MsgBox "Die Datei " & PathOfFileToSendOrReceive & " wurde vollständig gesendet!", vbInformation
        Else 'FILEFULLYRECEIVEDORSEND = FALSE/0
            MsgBox "Die Datei " & PathOfFileToSendOrReceive & " wurde NICHT vollständig gesendet!", vbCritical
        End If
    Else 'SENDFILE = FALSE/0
        If FileFullyReceivedorSend Then
            MsgBox "Die Datei " & PathOfFileToSendOrReceive & " wurde vollständig empfangen!", vbInformation
        Else 'FILEFULLYRECEIVEDORSEND = FALSE/0
            MsgBox "Die Datei " & PathOfFileToSendOrReceive & " wurde NICHT vollständig empfangen!", vbCritical
        End If
    End If
    Set frmTransfer = Nothing
    Unload Me

End Sub

Private Sub tmrSpeed_Timer()

' rechnet den durchschnittsspeed aus

Static TimeCounter As Integer

    TimeCounter = TimeCounter + 1
    If SendFile Then
        lblInfo(6).Caption = CStr(Round((nFilePos / 1024) / TimeCounter, 2)) & " kb/s"
    Else 'SENDFILE = FALSE/0
        lblInfo(6).Caption = CStr(Round((nBytesRead / 1024) / TimeCounter, 2)) & " kb/s"
    End If

End Sub

Private Sub tmrStart_Timer()

' startet alles wie oben beschrieben

    tmrStart.Enabled = False
    ' wichtig!
    wscFile.Close
    Do Until wscFile.State = 0
        DoEvents
    Loop
    If SendFile Then
        wscFile.LocalPort = 18252
        wscFile.Listen
        Listening = True
    Else 'SENDFILE = FALSE/0
        wscFile.LocalPort = 0
        wscFile.Connect RemoteIP, 18252
        Listening = False
    End If
    tmrSuccess.Enabled = True

End Sub

Private Sub tmrSuccess_Timer()

Static TimeOut As Integer

    ' prüft, ob eine Verbindung zustande gekommen ist
    TimeOut = TimeOut + 1
    If TimeOut = 20 Then
        tmrSuccess.Enabled = False
        tmrEnd.Enabled = True
        Exit Sub
    End If
    If Not wscFile.State = 7 Then
        Listening = Not Listening
        wscFile.Close
        Do Until wscFile.State = 0
            DoEvents
        Loop
        If Listening Then
            wscFile.LocalPort = 18252
            wscFile.Listen
        Else 'LISTENING = FALSE/0
            wscFile.LocalPort = 0
            wscFile.Connect RemoteIP, 18252
        End If
    Else 'NOT NOT...
        tmrSuccess.Enabled = False
        tmrSpeed.Enabled = True
        tmrConnectionCheck.Enabled = True
        If SendFile Then WinsockSendBinaryFile PathOfFileToSendOrReceive
    End If

End Sub

Public Sub UnloadfrmTransfer()

    Sending = False
    StopLoop = True
    If SendFile = False Or tmrSuccess.Enabled = False Then
        tmrEnd.Enabled = True
    Else 'NOT SENDFILE...
        tmrEnd.Enabled = True
    End If

End Sub

Public Sub WinsockSendBinaryFile(ByVal sFile As String)

Dim F            As Integer
Dim sBuffer      As String
Dim nFileSize    As Long
Dim nBytesToRead As Long
Const BlockSize = 1000

    On Error GoTo errh
    ' Größe der einzelnen Datenpakete
    nFilePos = 0
    ' ### Pfad gekürzt anzeigen
    lblInfo(4).Caption = modFunctions.CompactPath(frmTransfer, sFile, lblInfo(4))
    ' Datei im Binary-Mode öffnen
    F = FreeFile
    Open sFile For Binary As #F
    ' Dateiname extrahieren
    If InStr(sFile, "\") > 0 Then
        sFile = Mid$(sFile, InStrRev(sFile, "\") + 1)
    End If
    ' Dateigröße
    nFileSize = LOF(F)
    ' ### Dateigröße optimal anzeigen
    lblInfo(5).Caption = modFunctions.BytesToString(nFileSize)
    ' Sendevorgang starten
    With wscFile
        ' Empfänger mitteln, welche Datei und wieviele
        ' Daten gesendet werden
        .SendData "<begin size=" & CStr(nFileSize) & ";" & sFile & ">"
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
            Do While Sending: DoEvents: Loop
            ' sehr WICHTIG
            ' sonst sammeln sich alle Daten im Buffer... Das Programm zeigt dann an, dass die Datei
            ' gesendet wurde, obwohl noch aus dem buffer gesendet wird...
            If StopLoop Then
                GoTo errh
                Exit Do
            End If
            Sending = True
            .SendData sBuffer
            ' Fortschritt aktualisieren
            nFilePos = nFilePos + nBytesToRead
            ' ### bereits gesendete Bytes optimal anzeigen
            lblInfo(7).Caption = modFunctions.BytesToString(nFilePos)
            ProgressBar.Value = Int(nFilePos / nFileSize * 100)
            ' Wichtig!
            DoEvents
        Loop
    End With 'WSCFILE
    ' Datei schließen (Sendevorgang beendet)
    Close #F
    Do While Sending: DoEvents: Loop ' wartet bis das letzte packet gesendet wurde
    FileFullyReceivedorSend = True
    tmrConnectionCheck.Enabled = False
    tmrEnd.Enabled = True
    ' anders gehts nicht !!! hier ein unload me und die form lädt sich nach dem unload wieder wie von selbst :*(
    ' hat ca. 7 stunden geadauert das rauszufinden *arg*

Exit Sub

errh:
    Close #F
    tmrEnd.Enabled = True
    ' anders gehts nicht !!! hier ein unload me und die form lädt sich nach dem unload wieder wie von selbst :*(
    ' hat ca. 7 stunden geadauert das rauszufinden *arg*

End Sub

Private Sub wscFile_ConnectionRequest(ByVal requestID As Long)

    wscFile.Close
    Do Until wscFile.State = 0
        DoEvents
    Loop
    wscFile.Accept requestID

End Sub

Private Sub wscFile_DataArrival(ByVal bytesTotal As Long)

Dim sData    As String
Dim sTemp    As String
Static sFile As String

    ' Daten holen
    wscFile.GetData sData, vbString
    If Left$(sData, 12) = "<begin size=" Then
        ' Aha... eine neue Datei wird gesendet
        sData = Mid$(sData, 13)
        sTemp = Left$(sData, InStr(sData, ">") - 1)
        sData = Mid$(sData, InStr(sData, ">") + 1)
        ' Dateigröße und Dateiname ermitteln
        If InStr(sTemp, ";") > 0 Then
            nBytesTotal = Val(Left$(sTemp, InStr(sTemp, ";") - 1))
            sFile = Mid$(sTemp, InStr(sTemp, ";") + 1)
        Else 'NOT INSTR(STEMP,...
            nBytesTotal = Val(sTemp)
        End If
        sFile = PathOfFileToSendOrReceive & "\" & sFile
        sFile = modFunctions.GetNextFreeFilename(sFile)
        PathOfFileToSendOrReceive = sFile
        ' ### Speicherort "gekürzt" anzeigen
        lblInfo(4).Caption = modFunctions.CompactPath(frmTransfer, sFile, lblInfo(4))
        ' ### Größe optimal anzeigen
        lblInfo(5).Caption = modFunctions.BytesToString(nBytesTotal)
        ' ggf. Datei löschen, falls bereits existiert
        On Error Resume Next
        Kill sFile
        On Error GoTo 0
        ' Datei im Binary-Mode öffnen
        nFile = FreeFile
        Open sFile For Binary As #nFile
        nBytesRead = 0
    End If
    If Len(sData) > 0 And nFile > 0 Then
        ' bisher empfangene Daten...
        nBytesRead = nBytesRead + Len(sData)
        ' Daten in Datei speichern
        Put #nFile, , sData
        ' evtl. Fortschritt anzeigen
        ' ### übertragene Bytes optimal anzeigen
        lblInfo(7).Caption = modFunctions.BytesToString(nBytesRead)
        ProgressBar.Value = Int(nBytesRead / nBytesTotal * 100)
        ' Wenn alle Bytes empfangen wurden, Datei schließen
        If nBytesRead = nBytesTotal Then
            Close #nFile
            FileFullyReceivedorSend = True
            nFile = 0
            tmrConnectionCheck.Enabled = False
            tmrEnd.Enabled = True
        End If
    End If

End Sub

Private Sub wscFile_Error(ByVal Number As Integer, _
                          Description As String, _
                          ByVal Scode As Long, _
                          ByVal Source As String, _
                          ByVal HelpFile As String, _
                          ByVal HelpContext As Long, _
                          CancelDisplay As Boolean)

    Sending = False

End Sub

Private Sub wscFile_SendComplete()

    Sending = False

End Sub



