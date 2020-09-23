VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "VBMessenger 9 Server"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6825
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   339
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox picFrame 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   3735
      Index           =   2
      Left            =   120
      ScaleHeight     =   3735
      ScaleWidth      =   6495
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CommandButton cmdRemUser 
         Caption         =   "Benutzer entfernen"
         Height          =   315
         Left            =   3240
         TabIndex        =   12
         Top             =   3360
         Width           =   1575
      End
      Begin VB.CommandButton cmdAddUser 
         Caption         =   "Benutzer hinzufügen"
         Height          =   315
         Left            =   4920
         TabIndex        =   11
         Top             =   3360
         Width           =   1575
      End
      Begin ComctlLib.ListView lvAddRem 
         Height          =   3255
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   5741
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Benutzername"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Passwort"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Admin"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Timer tmrStats 
      Interval        =   30000
      Left            =   2640
      Top             =   5640
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Unten ausrichten
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   4830
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picFrame 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   3735
      Index           =   1
      Left            =   120
      ScaleHeight     =   3735
      ScaleWidth      =   6495
      TabIndex        =   1
      Top             =   480
      Width           =   6495
      Begin VB.ListBox lstServerEvents 
         Height          =   3735
         IntegralHeight  =   0   'False
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   6495
      End
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "VBMessenger 9 Server Beenden"
      Height          =   345
      Left            =   120
      TabIndex        =   7
      Top             =   4440
      Width           =   2895
   End
   Begin VB.PictureBox picTray 
      BorderStyle     =   0  'Kein
      Height          =   240
      Left            =   240
      Picture         =   "frmMain.frx":27A2
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      Top             =   6120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Timer tmrTime 
      Interval        =   10000
      Left            =   2160
      Top             =   5640
   End
   Begin VB.Timer tmrGetData 
      Interval        =   50
      Left            =   1680
      Top             =   5640
   End
   Begin VB.Timer tmrConnectionCheck 
      Interval        =   5000
      Left            =   1200
      Top             =   5640
   End
   Begin MSWinsockLib.Winsock wsc 
      Index           =   0
      Left            =   720
      Top             =   5640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wscMain 
      Left            =   240
      Top             =   5640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   81
   End
   Begin VB.PictureBox picFrame 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   3735
      Index           =   0
      Left            =   120
      ScaleHeight     =   3735
      ScaleWidth      =   6495
      TabIndex        =   2
      Top             =   480
      Width           =   6495
      Begin VB.CommandButton cmdActualize 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Aktualisieren"
         Height          =   315
         Left            =   5160
         TabIndex        =   4
         Top             =   3360
         Width           =   1335
      End
      Begin ComctlLib.ListView lvUser 
         Height          =   3255
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   5741
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   11
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Benutzername"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Passwort"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Admin"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "BanTime"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   4
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "wscID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   5
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "VoiceEnabled"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   6
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "VoiceID"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   7
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Gesendet"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   8
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Empfangen"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   9
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Gesendet (VoiceChat)"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(11) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   10
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Empfangen (VoiceChat)"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin ComctlLib.TabStrip tsServer 
      Height          =   4335
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   7646
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Session_Info"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Server_Log"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Benutzer verwalten"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' --------------------------------------------------
'
'             VBMessenger9 Server & Client
'
' (C) by Thorben Linneweber (2005)
'
' OpenSource
'
' Darf unverändert weitergegeben werden.
' Darf nur für den privaten Gebrauch verändert werden.
' Kommerzielle Nutzung ist untersagt!
'
' Verbesserungsvorschläge und Änderungen bitte an
' thorben_linneweber@hotmail.com
'
' ----------------------------------------------------
' --------------------------------------------------
'
' VBMessenger 9 Server
'
' Befehle an den Server (aus Sicht des Servers)
' ---------------------
'
'
'    login
'    [Username][Passwort]
'    -> LoginAnfrage
'
'    message
'    [Nachricht]
'    -> [Nachricht] an Alle
'
'    privatemessage
'    [Username]
'    -> Nachricht an [Username]
'
'    nudge
'    [Username]
'    -> Nudge an [Username]
'
'    file
'    [Username][Filename][Filegröße]
'    -> sendet an [Username] eine Fileanfrage für [Filename][Filegröße]
'
'
'    acceptfile
'    [Username][AcceptString(true\false\alreadyreceiving)]
'    -> sendet an [Username] den [AcceptString(true\false\alreadyreceiving)] für eine FileTransferanfrage
'
'    typing
'    [TypeString(true\false]
'    -> sendet an alle Benutzer den [TypeString(true\false]
'
'
'    kick
'    [Username]
'    -> kickt [Username] (geht nur, wenn der Sender Adminrechte hat)
'
'    ban
'    [Username]
'    -> banned [Username] (geht nur, wenn der Sender Adminrechte hat)
'
'    makeadmin
'    [Username]
'    -> macht [Username] zum Admin (geht nur, wenn der Sender Adminrechte hat)
'
'    giveupadmin
'    []
'    -> gibt die Adminrechte ab
'
'    kickvoice
'    [Username]
'    -> schmeißt [Username] aus dem VoiceChat
'
'    loginadmin
'    []
'    -> stellt die festen Adminrechte wieder her
'
'    askvoice
'    []
'    -> Anfrage für VoiceChat
'
'
' Rückgabewerte vom Server (Events) (aus Sicht des Clienten)
' ------------------------
'
'    login
'    [AcceptString(accept\deny)]
'    -> LoginAnfrage-Antwort
'
'    message
'    [Benutzername][Nachricht][Admin(true\false)]
'    -> [Nachricht] an Alle von [Benutzername]
'
'    privatemessage
'    [Benutzername][Nachricht][Admin(true\false)]
'    -> private [Nachricht] von [Benutzername]
'
'    privatemessageb
'    [Benutzername][Nachricht][Admin(true\false)]
'    -> Bestätigung von einer privaten [Nachricht]
'
'    nudge
'    [Benutzername][Admin(true\false)]
'    -> Nudge von [Benutzername]
'
'    nudgeb
'    [Benutzername][Admin(true\false)]
'    -> Nudge-Sende-Bestätigung
'
'    fileb
'    [Benutzername][Filename][Filesize][Admin(true\false)]
'    -> FileAnfragesendebestätigung
'
'    acceptfile
'    [Benutzername][AcceptString(true\false\alreadyreceiving)][Admin(true\false)]
'    -> Filetransfer Anfragebestätigung von [Benutzername], [AcceptString(true\false\alreadyreceiving)]
'
'    acceptfileb
'    [Benutzername][AcceptString(true\false\alreadyreceiving)][Admin(true\false)]
'    -> FiletransferAnfrageAntwortBestätigung
'
'    typing
'    [Username][TypeString(true\false)][Admin(true\false)]
'    -> [Username] tippt \ tippt nicht ; [TypeString(true\false)]
'
'    voice
'    [Username][Username][VoiceStr(enabled\disabled]
'    -> [Username] tippt \ tippt nicht ; [TypeString(true\false)]
'
'    askvoice
'    [AllowStr(true\false)]
'    -> Antwort auf Voiceanfrage
'
'    kick
'    [Username][Username2]
'    -> [Username] hat [Username2] gekickt
'
'    ban
'    [Username][Username2]
'    -> [Username] hat [Username2] gebannt
'
'    makeadmin
'    [Username][Username2]
'    -> [Username] hat [Username2] zum Admin gemacht
'
'    giveupadmin
'    [Username]
'    -> [Username] hat seine Adminrechte abgelegt
'
'    kickvoice
'    [Username][Username2]
'    -> [Username] kickt [Username2] aus dem VoiceChat
'
'
'
'
Option Explicit
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private RawDataOutgoing, RawDataIncoming, ServerEvent, ServerError
#End If
Private Type TUser
    Username                    As String    ' Benutzername
    UserPass                    As Long      ' Passwort (CRC32 aus Benutzername & Passwort)
    Admin                       As Boolean   ' Isser Admin?
    wscID                       As Integer   ' Der Winsock Index
    banTime                     As Integer   ' Wie lange isser gebannt?
    VoiceEnabled                As Boolean   ' VoiceChat an?
    VoicePlayerID               As Long      ' DX8 Voice PlayerID
    wscBytesReceived            As Long       ' Bytes empfangen (Chat)
    wscBytesSend                As Long       ' Bytes gesendet  (Chat)
    dp8BytesReceivedOldSessions As Long
    ' Bytes empfangen (Voice;nicht in der aktuellen Session)
    dp8BytesSendOldSessions     As Long
    ' Bytes gesendet  (Voice;nicht in der aktuellen Session)
End Type
Private Enum eLogData
    infodata
    ServerEvent
    ServerError
End Enum
' VOICE
Implements DirectPlay8Event
Implements DirectPlayVoiceEvent8
Private moCallBack As DirectPlay8Event
Private mfExit As Boolean
Private mfTerminate As Boolean
Private mlVoiceError As Long
' VOICE
Private ExitServer          As Boolean
Private ReceiveBuffers()    As String
Private AllUser()           As TUser
Private WinsockCount        As Long
Private Const Seperator     As String = "|||"
Private Const LVM_FIRST                        As Long = &H1000
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE     As Long = (LVM_FIRST + 54)
Private Const LVS_EX_GRIDLINES                 As Long = &H1
Private Const LVS_EX_FULLROWSELECT             As Long = &H20
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        lParam As Any) As Long

Private Sub AddLogEntry(ByVal Entry As String, _
                        LogDataType As eLogData)

    Select Case LogDataType
    Case ServerEvent
        lstServerEvents.AddItem "[Server " & Time & "] " & Entry
    Case infodata
        lstServerEvents.AddItem "[Information " & Time & "] " & Entry
    Case ServerError
        lstServerEvents.AddItem "[FEHLER " & Time & "] " & Entry
    End Select
    lstServerEvents.Selected(lstServerEvents.ListCount - 1) = True

End Sub

Private Function BytesToString(Bytes As Long) As String

Dim zwFileLen As Long

    zwFileLen = Bytes
    If zwFileLen <= 1024 Then
        BytesToString = zwFileLen & " Byte"
    ElseIf zwFileLen <= 1048576 Then 'NOT ZWFILELEN...
        BytesToString = Round(zwFileLen / (1024), 2) & " KB"
    Else 'NOT ZWFILELEN...
        BytesToString = Round(zwFileLen / 1048576, 2) & " MB"
    End If

End Function

Private Sub cmdActualize_Click()

Dim dpCI As DPN_CONNECTION_INFO
Dim i    As Integer

    On Error Resume Next
    lvUser.ListItems.Clear
    For i = 0 To UBoundEx(AllUser)
        lvUser.ListItems.Add , , AllUser(i).Username
        '~ wenn voice enabled ist...
        If Not AllUser(i).VoicePlayerID = -1 Then
            dpCI = dpp.GetConnectionInfo(AllUser(i).VoicePlayerID)
            lvUser.ListItems(i + 1).SubItems(10) = BytesToString(dpCI.lBytesReceivedGuaranteed + dpCI.lBytesReceivedNonGuaranteed + AllUser(i).dp8BytesReceivedOldSessions)
            lvUser.ListItems(i + 1).SubItems(9) = BytesToString(dpCI.lBytesSentGuaranteed + dpCI.lBytesSentNonGuaranteed + AllUser(i).dp8BytesSendOldSessions)
        Else 'NOT NOT...
            lvUser.ListItems(i + 1).SubItems(10) = BytesToString(AllUser(i).dp8BytesReceivedOldSessions)
            lvUser.ListItems(i + 1).SubItems(9) = BytesToString(AllUser(i).dp8BytesSendOldSessions)
        End If
        '~ wscBytes...
        lvUser.ListItems(i + 1).SubItems(8) = BytesToString(AllUser(i).wscBytesReceived)
        lvUser.ListItems(i + 1).SubItems(7) = BytesToString(AllUser(i).wscBytesSend)
        '~ Rest
        lvUser.ListItems(i + 1).SubItems(1) = AllUser(i).UserPass
        lvUser.ListItems(i + 1).SubItems(2) = IIf(AllUser(i).Admin, "true", "false")
        lvUser.ListItems(i + 1).SubItems(3) = AllUser(i).banTime
        lvUser.ListItems(i + 1).SubItems(4) = AllUser(i).wscID
        lvUser.ListItems(i + 1).SubItems(5) = AllUser(i).VoiceEnabled
        lvUser.ListItems(i + 1).SubItems(6) = AllUser(i).VoicePlayerID
    Next i

End Sub

Private Sub cmdAddUser_Click()

Dim NewUsername As String
Dim NewPassword As String
Dim NewAdmin    As String
Dim Crc32Check  As New clsCRC
Dim fnr         As Integer
Dim zw          As String
Dim i           As Integer
Dim e           As Integer

    On Error GoTo errh
    '~ Sicherungskopie
    If FileExists(App.Path & "\user.txt") Then Call FileCopy(App.Path & "\user.txt", App.Path & "\sec_user.txt")
    '~ Neuen Benutzernamen
    NewUsername = InputBox("Geben Sie einen neuen Benutzernamen ein", "Neuer Benutzer")
    If StrPtr(NewUsername) = 0 Then Exit Sub
    '~Benutzername schon vorhanden?
    e = -1
    For i = 1 To lvAddRem.ListItems.Count
        If lvAddRem.ListItems(i).Text = NewUsername Then
            If MsgBox("Der Benutzername existiert bereits! Möchten Sie für diesen Benutzer neuere Daten eintragen?", vbExclamation + vbYesNo) = vbYes Then
                e = i
                Exit For
            Else 'NOT MSGBOX("DER BENUTZERNAME EXISTIERT BEREITS! MÖCHTEN SIE FÜR DIESEN BENUTZER NEUERE DATEN EINTRAGEN?",...
                Exit Sub
            End If
        End If
    Next i
    '~Wenn nicht, dann adden...
    If e = -1 Then
        lvAddRem.ListItems.Add , , NewUsername
        e = lvAddRem.ListItems.Count
    End If
    '~ Neues Passwort
    NewPassword = InputBox("Geben Sie ein Passwort für den Benutzer ein", "Neuer Benutzer")
    If StrPtr(NewPassword) = 0 Then Exit Sub
    '~ Admin-Rechte
    If MsgBox("Soll der Benutzer feste Administratorrechte haben?", vbYesNo) = vbYes Then
        NewAdmin = "true"
    Else 'NOT MSGBOX("SOLL DER BENUTZER FESTE ADMINISTRATORRECHTE HABEN?",...
        NewAdmin = "false"
    End If
    lvAddRem.ListItems.Item(e).SubItems(1) = Crc32Check.CalculateString(NewPassword & NewUsername)
    lvAddRem.ListItems.Item(e).SubItems(2) = NewAdmin
    DoEvents
    Refresh
    DoEvents
    '~ so alles hinzugefügt...jetzt noch speichern
    fnr = FreeFile
    Open App.Path & "\user.txt" For Output As #fnr
    For i = 1 To lvAddRem.ListItems.Count
        Print #fnr, lvAddRem.ListItems(i).Text
        Print #fnr, lvAddRem.ListItems(i).SubItems(1)
        Print #fnr, lvAddRem.ListItems(i).SubItems(2)
    Next i
    Close #fnr
    If FileExists(App.Path & "\sec_user.txt") Then Kill App.Path & "\sec_user.txt"
    MsgBox "Benutzer registriert! (Neustart erforderlich)", vbInformation

Exit Sub

errh:
    MsgBox "Schwerer Ausnahmefehler beim Zugriff auf die Datei User.txt! Es wurde zuvor eine Sicherungskopie angelegt. (sec_user.txt)", vbCritical
    End

End Sub

Private Sub cmdEnd_Click()

    If MsgBox("Server wirklich beenden?", vbCritical + vbYesNo) = vbYes Then
        ExitServer = True
        Unload Me
    End If

End Sub

Private Sub cmdRemUser_Click()

Dim i   As Integer
Dim fnr As Integer

    On Error GoTo errh
    If (lvAddRem.SelectedItem Is Nothing) Then Exit Sub
    If LenB(lvAddRem.SelectedItem.Text) = 0 Then Exit Sub
    If FileExists(App.Path & "\user.txt") Then Call FileCopy(App.Path & "\user.txt", App.Path & "\sec_user.txt")
    lvAddRem.ListItems.Remove lvAddRem.SelectedItem.Index
    fnr = FreeFile
    Open App.Path & "\user.txt" For Output As #fnr
    For i = 1 To lvAddRem.ListItems.Count
        Print #fnr, lvAddRem.ListItems(i).Text
        Print #fnr, lvAddRem.ListItems(i).SubItems(1)
        Print #fnr, lvAddRem.ListItems(i).SubItems(2)
    Next i
    Close #fnr
    If FileExists(App.Path & "\sec_user.txt") Then Kill App.Path & "\sec_user.txt"
    MsgBox "Der Benutzer wurde erfolgreich entfernt. (Neustart erforderlich)", vbInformation

Exit Sub

errh:
    MsgBox "Schwerer Ausnahmefehler beim Zugriff auf die Datei User.txt! Es wurde zuvor eine Sicherungskopie angelegt. (sec_user.txt)", vbCritical
    End
    '

End Sub

Private Sub DirectPlay8Event_AddRemovePlayerGroup(ByVal lMsgID As Long, _
                                                  ByVal lPlayerID As Long, _
                                                  ByVal lGroupID As Long, _
                                                  fRejectMsg As Boolean)

    If (Not moCallBack Is Nothing) Then moCallBack.AddRemovePlayerGroup lMsgID, lPlayerID, lGroupID, fRejectMsg

End Sub

Private Sub DirectPlay8Event_AppDesc(fRejectMsg As Boolean)

    If (Not moCallBack Is Nothing) Then moCallBack.AppDesc fRejectMsg

End Sub

Private Sub DirectPlay8Event_AsyncOpComplete(dpnotify As DxVBLibA.DPNMSG_ASYNC_OP_COMPLETE, _
                                             fRejectMsg As Boolean)

    If (Not moCallBack Is Nothing) Then moCallBack.AsyncOpComplete dpnotify, fRejectMsg

End Sub

Private Sub DirectPlay8Event_ConnectComplete(dpnotify As DxVBLibA.DPNMSG_CONNECT_COMPLETE, _
                                             fRejectMsg As Boolean)

    If (Not moCallBack Is Nothing) Then moCallBack.ConnectComplete dpnotify, fRejectMsg

End Sub

Private Sub DirectPlay8Event_CreateGroup(ByVal lGroupID As Long, _
                                         ByVal lOwnerID As Long, _
                                         fRejectMsg As Boolean)

'VB requires that we must implement *every* member of this interface

    If (Not moCallBack Is Nothing) Then moCallBack.CreateGroup lGroupID, lOwnerID, fRejectMsg

End Sub

Private Sub DirectPlay8Event_CreatePlayer(ByVal lPlayerID As Long, _
                                          fRejectMsg As Boolean)

Dim dpPeer As DPN_PLAYER_INFO

    On Error Resume Next
    dpPeer = dpp.GetPeerInfo(lPlayerID)
    If (dpPeer.lPlayerFlags And DPNPLAYER_LOCAL) = DPNPLAYER_LOCAL Then
        glMyPlayerID = lPlayerID
    End If
    If (Not moCallBack Is Nothing) Then moCallBack.CreatePlayer lPlayerID, fRejectMsg

End Sub

Private Sub DirectPlay8Event_DestroyGroup(ByVal lGroupID As Long, _
                                          ByVal lReason As Long, _
                                          fRejectMsg As Boolean)

'VB requires that we must implement *every* member of this interface

    If (Not moCallBack Is Nothing) Then moCallBack.DestroyGroup lGroupID, lReason, fRejectMsg

End Sub

Private Sub DirectPlay8Event_DestroyPlayer(ByVal lPlayerID As Long, _
                                           ByVal lReason As Long, _
                                           fRejectMsg As Boolean)

Dim dpPeer  As DPN_PLAYER_INFO
Dim lOffset As Long
Dim oBuf()  As Byte
Dim lmsg    As Long
Dim dpCI    As DPN_CONNECTION_INFO

    On Error Resume Next
    If lPlayerID <> glMyPlayerID Then 'ignore removing myself
        RemovePlayer lPlayerID
    End If
    '~ bevor wir unsere Peer destroyen, heben wir seine Bytes für spätere Generationen auf ;)
    dpCI = dpp.GetConnectionInfo(lPlayerID)
    dpPeer = dpp.GetPeerInfo(lPlayerID)
    AllUser(GetArrayIDFromUsername(dpPeer.Name)).dp8BytesReceivedOldSessions = AllUser(GetArrayIDFromUsername(dpPeer.Name)).dp8BytesReceivedOldSessions + dpCI.lBytesReceivedGuaranteed + dpCI.lBytesReceivedNonGuaranteed
    AllUser(GetArrayIDFromUsername(dpPeer.Name)).dp8BytesSendOldSessions = AllUser(GetArrayIDFromUsername(dpPeer.Name)).dp8BytesSendOldSessions + dpCI.lBytesSentGuaranteed + dpCI.lBytesSentNonGuaranteed
    Debug.Print AllUser(GetArrayIDFromUsername(dpPeer.Name)).dp8BytesSendOldSessions
    '~ zur Sicherheit
    lmsg = MsgUpdatePlayerLst
    dpp.DestroyPeer lPlayerID, 0, lmsg, LenB(lmsg)
    '~ senden, dass einer von uns gegangen ist
    lmsg = MsgUpdatePlayerLst
    lOffset = NewBuffer(oBuf)
    AddDataToBuffer oBuf, lmsg, LenB(lmsg), lOffset
    dpp.SendTo DPNID_ALL_PLAYERS_GROUP, oBuf, 0, DPNSEND_NOLOOPBACK Or DPNSEND_GUARANTEED
    AddLogEntry dpPeer.Name & " hat den VoiceChat verlassen", ServerEvent
    SendToAll "voice|" & dpPeer.Name & "|disabled|" & IIf(AllUser(GetArrayIDFromUsername(dpPeer.Name)).Admin, "true", "false")
    'VB requires that we must implement *every* member of this interface
    If (Not moCallBack Is Nothing) Then moCallBack.DestroyPlayer lPlayerID, lReason, fRejectMsg

End Sub

Private Sub DirectPlay8Event_EnumHostsQuery(dpnotify As DxVBLibA.DPNMSG_ENUM_HOSTS_QUERY, _
                                            fRejectMsg As Boolean)

'VB requires that we must implement *every* member of this interface

    If (Not moCallBack Is Nothing) Then moCallBack.EnumHostsQuery dpnotify, fRejectMsg

End Sub

Private Sub DirectPlay8Event_EnumHostsResponse(dpnotify As DxVBLibA.DPNMSG_ENUM_HOSTS_RESPONSE, _
                                               fRejectMsg As Boolean)

'VB requires that we must implement *every* member of this interface

    If (Not moCallBack Is Nothing) Then moCallBack.EnumHostsResponse dpnotify, fRejectMsg

End Sub

Private Sub DirectPlay8Event_HostMigrate(ByVal lNewHostID As Long, _
                                         fRejectMsg As Boolean)

'VB requires that we must implement *every* member of this interface

    If (Not moCallBack Is Nothing) Then moCallBack.HostMigrate lNewHostID, fRejectMsg

End Sub

Private Sub DirectPlay8Event_IndicateConnect(dpnotify As DxVBLibA.DPNMSG_INDICATE_CONNECT, _
                                             fRejectMsg As Boolean)

'VB requires that we must implement *every* member of this interface

    If (Not moCallBack Is Nothing) Then moCallBack.IndicateConnect dpnotify, fRejectMsg

End Sub

Private Sub DirectPlay8Event_IndicatedConnectAborted(fRejectMsg As Boolean)

'VB requires that we must implement *every* member of this interface

    If (Not moCallBack Is Nothing) Then moCallBack.IndicatedConnectAborted fRejectMsg

End Sub

Private Sub DirectPlay8Event_InfoNotify(ByVal lMsgID As Long, _
                                        ByVal lNotifyID As Long, _
                                        fRejectMsg As Boolean)

'VB requires that we must implement *every* member of this interface

    If (Not moCallBack Is Nothing) Then moCallBack.InfoNotify lMsgID, lNotifyID, fRejectMsg

End Sub

Private Sub DirectPlay8Event_Receive(dpnotify As DxVBLibA.DPNMSG_RECEIVE, _
                                     fRejectMsg As Boolean)

Dim lmsg   As Long, lOffset As Long
Dim dpPeer As DPN_PLAYER_INFO
Dim oBuf() As Byte
Dim Address As DirectPlay8Address
Dim localAddress As DirectPlay8Address
Dim MyLocalIP As String

Set localAddress = dpp.GetLocalHostAddress(1)

MyLocalIP = CStr(localAddress.GetComponentString(DPN_KEY_HOSTNAME))

Set Address = dpp.GetPeerAddress(dpnotify.idSender)


 

    With dpnotify
        GetDataFromBuffer .ReceivedData, lmsg, LenB(lmsg), lOffset
        Select Case lmsg
        Case MsgAskToJoin
            '~ damit sich nicht jeder (direct)X-beliebige anmelden kann
            dpPeer = dpp.GetPeerInfo(dpnotify.idSender)
            
            AddLogEntry "Jemand versucht den VoiceChat mit dem Namen " & dpPeer.Name & " zu betreten (" & CStr(Address.GetComponentString(DPN_KEY_HOSTNAME)) & ")", ServerEvent
    
            
            If (CStr(Address.GetComponentString(DPN_KEY_HOSTNAME)) = CStr(wsc(AllUser(GetArrayIDFromUsername(dpPeer.Name)).wscID).RemoteHostIP)) Or (CStr(Address.GetComponentString(DPN_KEY_HOSTNAME)) = MyLocalIP) Then

                AddLogEntry dpPeer.Name & " konnte anhand seiner IP erfolgreich identifiziert werden", ServerEvent
                AllUser(GetArrayIDFromUsername(dpPeer.Name)).VoiceEnabled = True
 
                AllUser(GetArrayIDFromUsername(dpPeer.Name)).VoiceEnabled = True
                
                AllUser(GetArrayIDFromUsername(dpPeer.Name)).VoicePlayerID = dpnotify.idSender
                'Accept this connection
                lmsg = MsgAcceptJoin
                lOffset = NewBuffer(oBuf)
                AddDataToBuffer oBuf, lmsg, LenB(lmsg), lOffset
                dpp.SendTo dpnotify.idSender, oBuf, 0, DPNSEND_NOLOOPBACK
                SendToAll "voice|" & dpPeer.Name & "|enabled|" & IIf(AllUser(GetArrayIDFromUsername(dpPeer.Name)).Admin, "true", "false")
                'Notify everyone that this player has joined
                lmsg = MsgUpdatePlayerLst
                lOffset = NewBuffer(oBuf)
                AddDataToBuffer oBuf, lmsg, LenB(lmsg), lOffset
                dpp.SendTo DPNID_ALL_PLAYERS_GROUP, oBuf, 0, DPNSEND_NOLOOPBACK Or DPNSEND_GUARANTEED
                '''dpnotify.lDataSize
                Else
                
                AddLogEntry CStr(Address.GetComponentString(DPN_KEY_HOSTNAME)) & " hat versucht sich mit falschem Benutzernamen im VoiceChat anzumelden", ServerEvent
                
                
            End If
        End Select
    End With 'DPNOTIFY
    If (Not moCallBack Is Nothing) Then moCallBack.Receive dpnotify, fRejectMsg

End Sub

Private Sub DirectPlay8Event_SendComplete(dpnotify As DxVBLibA.DPNMSG_SEND_COMPLETE, _
                                          fRejectMsg As Boolean)

'VB requires that we must implement *every* member of this interface

    If (Not moCallBack Is Nothing) Then moCallBack.SendComplete dpnotify, fRejectMsg

End Sub

Private Sub DirectPlay8Event_TerminateSession(dpnotify As DxVBLibA.DPNMSG_TERMINATE_SESSION, _
                                              fRejectMsg As Boolean)

'VB requires that we must implement *every* member of this interface

    If (Not moCallBack Is Nothing) Then moCallBack.TerminateSession dpnotify, fRejectMsg
    mfTerminate = True
    '  tmrUpdate.Enabled = True

End Sub

Private Sub DirectPlayVoiceEvent8_ConnectResult(ByVal ResultCode As Long)

End Sub

Private Sub DirectPlayVoiceEvent8_CreateVoicePlayer(ByVal playerID As Long, _
                                                    ByVal flags As Long)

End Sub

Private Sub DirectPlayVoiceEvent8_DeleteVoicePlayer(ByVal playerID As Long)

'VB requires that we must implement *every* member of this interface


End Sub

Private Sub DirectPlayVoiceEvent8_DisconnectResult(ByVal ResultCode As Long)

'VB requires that we must implement *every* member of this interface


End Sub

Private Sub DirectPlayVoiceEvent8_HostMigrated(ByVal NewHostID As Long, _
                                               ByVal NewServer As DxVBLibA.DirectPlayVoiceServer8)

'VB requires that we must implement *every* member of this interface


End Sub

Private Sub DirectPlayVoiceEvent8_InputLevel(ByVal PeakLevel As Long, _
                                             ByVal RecordVolume As Long)

'VB requires that we must implement *every* member of this interface


End Sub

Private Sub DirectPlayVoiceEvent8_OutputLevel(ByVal PeakLevel As Long, _
                                              ByVal OutputVolume As Long)

'VB requires that we must implement *every* member of this interface


End Sub

Private Sub DirectPlayVoiceEvent8_PlayerOutputLevel(ByVal SourcePlayerID As Long, _
                                                    ByVal PeakLevel As Long)

'VB requires that we must implement *every* member of this interface


End Sub

Private Sub DirectPlayVoiceEvent8_PlayerVoiceStart(ByVal SourcePlayerID As Long)

'VB requires that we must implement *every* member of this interface


End Sub

Private Sub DirectPlayVoiceEvent8_PlayerVoiceStop(ByVal SourcePlayerID As Long)

'VB requires that we must implement *every* member of this interface


End Sub

Private Sub DirectPlayVoiceEvent8_RecordStart(ByVal PeakVolume As Long)

'VB requires that we must implement *every* member of this interface


End Sub

Private Sub DirectPlayVoiceEvent8_RecordStop(ByVal PeakVolume As Long)

'VB requires that we must implement *every* member of this interface


End Sub

Private Sub DirectPlayVoiceEvent8_SessionLost(ByVal ResultCode As Long)

'VB requires that we must implement *every* member of this interface


End Sub

Private Function FileExists(Path As String) As Boolean

Const NotFile As Double = vbDirectory Or vbVolume

    On Error Resume Next
    FileExists = (GetAttr(Path) And NotFile) = 0
    On Error GoTo 0

End Function

Private Sub FixedDataArrival(Index As Integer, _
                             strData As String)

Dim v As Variant

    On Error GoTo IntrusionDetection
    If Not Len(strData) = 0 Then
        v = Split(strData, "|")
        If CStr(v(0)) = "login" Then
            UserLogin CStr(v(1)), CStr(v(2)), Index
            Exit Sub
        End If
        '~ Validate User first!
        If Not IsUserLoggedIn(Index) Then
            Exit Sub
        End If
        Select Case CStr(v(0))
        Case "message"
            SendToAll "message|" & AllUser(GetArrayIDFromIndex(Index)).Username & "|" & v(1) & "|" & IIf(AllUser(GetArrayIDFromIndex(Index)).Admin, "true", "false")
            AddLogEntry AllUser(GetArrayIDFromIndex(Index)).Username & " hat eine Nachricht an alle Benutzer gesendet", ServerEvent
        Case "privatemessage"
            SendData AllUser(GetArrayIDFromUsername(CStr(v(1)))).wscID, "privatemessage|" & AllUser(GetArrayIDFromIndex(Index)).Username & "|" & v(2) & "|" & IIf(AllUser(GetArrayIDFromIndex(Index)).Admin, "true", "false")
            SendData Index, "privatemessageb|" & CStr(v(1)) & "|" & CStr(v(2)) & "|" & IIf(AllUser(GetArrayIDFromIndex(Index)).Admin, "true", "false")
            AddLogEntry AllUser(GetArrayIDFromIndex(Index)).Username & " hat eine Nachricht an " & AllUser(GetArrayIDFromUsername(CStr(v(1)))).Username & " gesendet", ServerEvent
        Case "nudge"
            SendData AllUser(GetArrayIDFromUsername(CStr(v(1)))).wscID, "nudge|" & AllUser(GetArrayIDFromIndex(Index)).Username & "|" & IIf(AllUser(GetArrayIDFromIndex(Index)).Admin, "true", "false")
            SendData Index, "nudgeb|" & CStr(v(1)) & "|" & IIf(AllUser(GetArrayIDFromIndex(Index)).Admin, "true", "false")
            AddLogEntry AllUser(GetArrayIDFromIndex(Index)).Username & " hat einen 'Nudge' an " & AllUser(GetArrayIDFromUsername(CStr(v(1)))).Username & " gesendet", ServerEvent
        Case "file"
            SendData AllUser(GetArrayIDFromUsername(CStr(v(1)))).wscID, "file|" & AllUser(GetArrayIDFromIndex(Index)).Username & "|" & CStr(v(2)) & "|" & CStr(v(3)) & "|" & wsc(Index).RemoteHostIP & "|" & IIf(AllUser(GetArrayIDFromIndex(Index)).Admin, "true", "false")
            SendData Index, "fileb|" & CStr(v(1)) & "|" & CStr(v(2)) & "|" & CStr(v(3)) & "|" & IIf(AllUser(GetArrayIDFromIndex(Index)).Admin, "true", "false")
            AddLogEntry AllUser(GetArrayIDFromIndex(Index)).Username & " möchte " & AllUser(GetArrayIDFromUsername(CStr(v(1)))).Username & " eine Datei senden", ServerEvent
        Case "acceptfile"
            SendData AllUser(GetArrayIDFromUsername(CStr(v(1)))).wscID, "acceptfile|" & AllUser(GetArrayIDFromIndex(Index)).Username & "|" & CStr(v(2)) & "|" & wsc(Index).RemoteHostIP & "|" & IIf(AllUser(GetArrayIDFromIndex(Index)).Admin, "true", "false")
            SendData Index, "acceptfileb|" & CStr(v(1)) & "|" & CStr(v(2)) & "|" & IIf(AllUser(GetArrayIDFromIndex(Index)).Admin, "true", "false")
            AddLogEntry AllUser(GetArrayIDFromIndex(Index)).Username & " hat auf die Anfrage von " & AllUser(GetArrayIDFromUsername(CStr(v(1)))).Username & " geantwortet", ServerEvent
        Case "typing"
            SendToAll "typing|" & AllUser(GetArrayIDFromIndex(Index)).Username & "|" & CStr(v(1)) & "|" & IIf(AllUser(GetArrayIDFromIndex(Index)).Admin, "true", "false")
            '
            '
            '~ Admin Privilegien ... jetzt wirds lustig ;)
        Case "kick"
            ' ~ erstmal gucken, ob derjenige überhaupt admin-rechte hat ;)
            If AllUser(GetArrayIDFromIndex(Index)).Admin Then
                wsc(AllUser(GetArrayIDFromUsername(CStr(v(1)))).wscID).Close
                AddLogEntry AllUser(GetArrayIDFromIndex(Index)).Username & " hat " & AllUser(GetArrayIDFromUsername(CStr(v(1)))).Username & " gekicked", ServerEvent
                SendToAll "kick|" & AllUser(GetArrayIDFromIndex(Index)).Username & "|" & AllUser(GetArrayIDFromUsername(CStr(v(1)))).Username
                If AllUser(GetArrayIDFromUsername(CStr(v(1)))).VoiceEnabled = True Then
                    KickVoicePlayer AllUser(GetArrayIDFromUsername(CStr(v(1)))).VoicePlayerID
                    SendToAll "kickvoice|" & AllUser(GetArrayIDFromIndex(Index)).Username & "|" & AllUser(GetArrayIDFromUsername(CStr(v(1)))).Username
                    AddLogEntry AllUser(GetArrayIDFromIndex(Index)).Username & " hat " & AllUser(GetArrayIDFromUsername(CStr(v(1)))).Username & " aus dem VoiceChat gekicked", ServerEvent
                End If
            End If
        Case "ban"
            ' ~ erstmal gucken, ob derjenige überhaupt admin-rechte hat ;)
            If AllUser(GetArrayIDFromIndex(Index)).Admin Then
                wsc(AllUser(GetArrayIDFromUsername(CStr(v(1)))).wscID).Close
                AllUser(GetArrayIDFromUsername(CStr(v(1)))).banTime = 29
                ' -> macht 5min ((29+1)*10 sec.)
                AddLogEntry AllUser(GetArrayIDFromIndex(Index)).Username & " hat einen Ban gegen " & AllUser(GetArrayIDFromUsername(CStr(v(1)))).Username & " ausgesprochen", ServerEvent
                SendToAll "ban|" & AllUser(GetArrayIDFromIndex(Index)).Username & "|" & AllUser(GetArrayIDFromUsername(CStr(v(1)))).Username
                If AllUser(GetArrayIDFromUsername(CStr(v(1)))).VoiceEnabled = True Then
                    KickVoicePlayer AllUser(GetArrayIDFromUsername(CStr(v(1)))).VoicePlayerID
                    SendToAll "kickvoice|" & AllUser(GetArrayIDFromIndex(Index)).Username & "|" & AllUser(GetArrayIDFromUsername(CStr(v(1)))).Username
                    AddLogEntry AllUser(GetArrayIDFromIndex(Index)).Username & " hat " & AllUser(GetArrayIDFromUsername(CStr(v(1)))).Username & " aus dem VoiceChat gekicked", ServerEvent
                End If
            End If
        Case "makeadmin"
            ' ~ erstmal gucken, ob derjenige überhaupt admin-rechte hat ;)
            If AllUser(GetArrayIDFromIndex(Index)).Admin Then
                If Not AllUser(GetArrayIDFromUsername(CStr(v(1)))).Admin Then
                    ' man kann nur leute zu admins machen, wenn sie noch keine admins sind
                    AllUser(GetArrayIDFromUsername(CStr(v(1)))).Admin = True
                    AddLogEntry AllUser(GetArrayIDFromIndex(Index)).Username & " hat " & AllUser(GetArrayIDFromUsername(CStr(v(1)))).Username & " Administrator Rechte gegeben.", ServerEvent
                    SendToAll "makeadmin|" & AllUser(GetArrayIDFromIndex(Index)).Username & "|" & AllUser(GetArrayIDFromUsername(CStr(v(1)))).Username
                End If
            End If
        Case "giveupadmin"
            ' ~ erstmal gucken, ob derjenige überhaupt admin-rechte hat ;)
            If AllUser(GetArrayIDFromIndex(Index)).Admin Then
                AllUser(GetArrayIDFromIndex(Index)).Admin = False
                AddLogEntry AllUser(GetArrayIDFromIndex(Index)).Username & " hat seine Administartor-Privilegien aufgegeben.", ServerEvent
                SendToAll "giveupadmin|" & AllUser(GetArrayIDFromIndex(Index)).Username
            End If
        Case "kickvoice"
            ' ~ erstmal gucken, ob derjenige überhaupt admin-rechte hat ;)
            If AllUser(GetArrayIDFromIndex(Index)).Admin Then
                If AllUser(GetArrayIDFromUsername(CStr(v(1)))).VoiceEnabled = True Then
                    KickVoicePlayer AllUser(GetArrayIDFromUsername(CStr(v(1)))).VoicePlayerID
                    SendToAll "kickvoice|" & AllUser(GetArrayIDFromIndex(Index)).Username & "|" & AllUser(GetArrayIDFromUsername(CStr(v(1)))).Username
                    AddLogEntry AllUser(GetArrayIDFromIndex(Index)).Username & " hat " & AllUser(GetArrayIDFromUsername(CStr(v(1)))).Username & " aus dem VoiceChat gekicked", ServerEvent
                End If
            End If
        Case "loginadmin"
            If RestoreOriginalAdminRights(GetArrayIDFromIndex(Index)) Then
                SendToAll "loginadmin|" & AllUser(GetArrayIDFromIndex(Index)).Username
                AddLogEntry AllUser(GetArrayIDFromIndex(Index)).Username & " hat seine Administartor-Privilegien wieder aufgenommen.", ServerEvent
            End If
        Case "askvoice"
            If AllUser(GetArrayIDFromIndex(Index)).VoicePlayerID = -1 Then '~der player hat keine Voice Session am laufen!!!
                SendData Index, "askvoice|true"
                AddLogEntry AllUser(GetArrayIDFromIndex(Index)).Username & " hat den VoiceChat betreten.", ServerEvent
            Else 'NOT NOT...
                SendData Index, "askvoice|false"
                AddLogEntry AllUser(GetArrayIDFromIndex(Index)).Username & " wurde der Zugang zum VoiceChat verwehrt.", ServerEvent
            End If
        Case Else
            AddLogEntry "Ungültiges Paket: " & strData, ServerError
        End Select
    End If

Exit Sub

IntrusionDetection:
    AddLogEntry "Ungültiges Paket: " & strData, ServerError

End Sub

Private Sub Form_Initialize()

    InitCommonControls

End Sub

Private Sub Form_Load()

Dim i As Integer

    If App.PrevInstance Then End
    frmMain.Hide
    ReDim ReceiveBuffers(0)
    ReDim SendingData(0)
    ReadUserData
    tsServer_Click
    cmdActualize_Click
    DoEvents
    modSysTray.AddTray picTray, "VBMessenger9 Server", picTray
    FullRowSelect lvUser
    FullRowSelect lvAddRem
    GridLines lvUser
    GridLines lvAddRem
    wscMain.Listen
    '
    '
    ' ~Chat
    AddLogEntry "Server gestartet", infodata
    AddLogEntry UBoundEx(AllUser) + 1 & " Benutzer geladen", infodata
    '
    '
    '~Voice
    modDServer.gsUserName = "VBMessenger_VoiceServer"
    StartHosting Me
    AddLogEntry "DirectX8 Voice Server geladen", infodata
    '~
    '
    tmrStats_Timer
    For i = 0 To UBoundEx(AllUser)
        lvAddRem.ListItems.Add , , AllUser(i).Username
        lvAddRem.ListItems.Item(i + 1).SubItems(1) = AllUser(i).UserPass
        lvAddRem.ListItems.Item(i + 1).SubItems(2) = IIf(AllUser(i).Admin, "true", "false")
    Next i

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)

    If Not ExitServer And Not vbAppWindows Then
        Cancel = True
        Me.Hide
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim i   As Integer
Dim fnr As Integer

    On Error Resume Next
    modSysTray.RemTray
    wscMain.Close
    For i = 1 To WinsockCount
        wsc(WinsockCount).Close
        Unload wsc(WinsockCount)
        DoEvents
    Next i
    AddLogEntry "Winsocks beendet", infodata
    Cleanup
    AddLogEntry "DirectX 8 Voice Server geschlossen", infodata
    On Error Resume Next
    MkDir App.Path & "\ServerLogs"
    On Error GoTo 0
    fnr = FreeFile
    Open App.Path & "\ServerLogs\ServerLog_" & GetNow & ".txt" For Output As #fnr
    For i = 0 To lstServerEvents.ListCount - 1
        Print #fnr, lstServerEvents.List(i)
        DoEvents
    Next i
    AddLogEntry "Schließe Server-Log", infodata
    Print #fnr, "[Server " & Time & "] " & "Schließe Server-Log"
    Close #fnr
    DoSleep 200

End Sub

Private Sub FullRowSelect(LV As ListView)

Dim State As Long

    State = True
    SendMessage LV.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_FULLROWSELECT, ByVal State

End Sub

Private Function GetArrayIDFromIndex(Index As Integer) As Long

'~ WinsockIndex -> ArrayIndex

Dim i As Integer

    For i = 0 To UBoundEx(AllUser)
        If AllUser(i).wscID = Index Then
            GetArrayIDFromIndex = i
            Exit Function
        End If
    Next i
    GetArrayIDFromIndex = -1

End Function

Private Function GetArrayIDFromUsername(Username As String) As Long

'~ Benutzername -> ArrayIndex

Dim i As Integer

    For i = 0 To UBoundEx(AllUser)
        If AllUser(i).Username = Username Then
            GetArrayIDFromUsername = i
            Exit Function
        End If
    Next i
    GetArrayIDFromUsername = -1

End Function

Private Function GetNow() As String

    GetNow = Replace$(Replace$(Replace$(Replace$(Replace$(Replace$(Replace$(Replace$(Replace$(Now, "/", "."), "\", "."), ":", "."), "*", "."), "?", "."), Chr$(34), "."), "<", "."), ">", "."), "|", ".")

End Function

Private Sub GridLines(LV As ListView)

Dim State As Long

    State = True
    SendMessage LV.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_GRIDLINES, ByVal State

End Sub

Private Function IsUserLoggedIn(Index As Integer) As Boolean

Dim i As Integer

    For i = 0 To UBoundEx(AllUser)
        If AllUser(i).wscID = Index Then
            IsUserLoggedIn = True
            Exit For
        End If
    Next i

End Function

Private Sub KickVoicePlayer(VoicePlayerID As Long)

Dim lmsg   As Long, lOffset As Long
Dim oBuf() As Byte

    On Error Resume Next
    lmsg = MsgUpdatePlayerLst
    dpp.DestroyPeer VoicePlayerID, 0, lmsg, LenB(lmsg)

End Sub

Private Sub picTray_MouseMove(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)

Static lngMsg As Long

    lngMsg = X / Screen.TwipsPerPixelX
    Select Case lngMsg
    Case WM_LBUTTONDBLCLK
        On Error Resume Next
        frmMain.WindowState = vbNormal
        frmMain.Show
        On Error GoTo 0
    End Select

End Sub

Private Function ReadUserData() As Boolean

Dim fnr  As Integer
Dim zw   As String
Dim line As Long

    ' TODO -> Inifiles
    fnr = FreeFile
    If FileExists(App.Path & "\user.txt") Then
        Open App.Path & "\user.txt" For Input As #fnr
        Do While Not EOF(fnr)
            ReDim Preserve AllUser(line)
            Line Input #fnr, zw
            AllUser(line).Username = zw
            Line Input #fnr, zw
            AllUser(line).UserPass = CLng(zw)
            Line Input #fnr, zw
            AllUser(line).Admin = IIf(zw = "true", True, False)
            AllUser(line).wscID = -1
            AllUser(line).VoicePlayerID = -1
            line = line + 1
        Loop
        Close #fnr
        ReadUserData = True
    End If

End Function

Private Sub RefreshUserLst()

Dim i                  As Integer

    For i = 0 To UBoundEx(AllUser)
        With AllUser(i)
            '  Debug.Print AllUser(i).Username
            If Not .wscID = -1 Then
                If Not wsc(.wscID).State = 7 Then
                    .wscID = -1  'wscid wieder auf -1 (=offline)
                    SendToAll "typing|" & .Username & "|false|" & IIf(.Admin, "true", "false")
                    ' tote schreiben nicht ;)
                    SendToAll "userlst|" & .Username & "|false"
                    ' und sind auch nicht online :D
                    Call RestoreOriginalAdminRights(i)
                    ' original admin rechte der person wiederherstellen
                    AddLogEntry .Username & " ist Offline", ServerEvent
                End If
            End If
        End With 'ALLUSER(I)
    Next i
    ' alle aussortiert, die offline sind

End Sub

Private Sub RemovePlayer(ByVal lPlayerID As Long)

Dim i As Integer

    For i = 0 To UBoundEx(AllUser)
        If AllUser(i).VoicePlayerID = lPlayerID Then
            AllUser(i).VoicePlayerID = -1
            AllUser(i).VoiceEnabled = False
        End If
    Next i

End Sub

Private Function RestoreOriginalAdminRights(ArrayIndex As Integer) As Boolean

Dim fnr  As Integer
Dim zw   As String
Dim line As Long

    ' ~gibt true zurück wenn von false auf true geändert wurde
    ' TODO -> Inifiles
    If FileExists(App.Path & "\user.txt") Then
        If AllUser(ArrayIndex).Admin Then Exit Function
        fnr = FreeFile
        Open App.Path & "\user.txt" For Input As #fnr
        Do While Not EOF(fnr)
            Line Input #fnr, zw
            Line Input #fnr, zw
            Line Input #fnr, zw
            If ArrayIndex = line Then
                AllUser(line).Admin = zw
                If zw = "true" Then
                    RestoreOriginalAdminRights = True
                End If
                Exit Do
            End If
            line = line + 1
        Loop
        Close #fnr
    End If

End Function

Private Sub SendData(Index As Integer, _
                     Text As String)

Dim ArrID As Integer

    On Error Resume Next
    ArrID = GetArrayIDFromIndex(Index)
    If Not Index = -1 Then
        If wsc(Index).State = 7 Then
            wsc(Index).SendData Text & Seperator
            AllUser(ArrID).wscBytesSend = AllUser(ArrID).wscBytesSend + Len(Text)
        End If
    End If

End Sub

Private Sub SendToAll(Text As String, _
                      Optional ExceptionIndex As Integer = -1)

Dim i As Integer
Dim e As String

    On Error Resume Next
    For i = 0 To WinsockCount
        If Not i = ExceptionIndex Then SendData i, Text
    Next i

End Sub

Private Sub SendWholeLst(Index As Integer)

Dim i As Integer

    For i = 0 To UBoundEx(AllUser)
        With AllUser(i)
            If Not .wscID = -1 Then
                If wsc(.wscID).State = 7 Then
                    SendData Index, "userlst|" & AllUser(i).Username & "|true"
                    ' index alle benutzer senden!
                End If
            End If
        End With 'ALLUSER(I)
    Next i

End Sub

Private Sub tmrConnectionCheck_Timer()

' periodisch gucken, wer online/offline ist

    RefreshUserLst

End Sub

Private Sub tmrGetData_Timer()

Dim temp  As Long
Dim Index As Integer

    ' Das prüfen kommt in einen Timer?
    ' WIESO?!
    ' Ganz einfach....
    ' angenommen, es kommt eine Nachricht an, und die wird gerade in dieser Schleife verarbeitei,
    ' allerdings im Winsock Data_Arrival. Dann gibt es große Probleme, wenn wärend dieser Bearbeitung ein
    ' weiteres Packets arrived! (hört sich komisch an, iss aber so)
    For Index = 0 To WinsockCount
        Do While InStr(1, ReceiveBuffers(Index), Seperator) > 0
            temp = InStr(1, ReceiveBuffers(Index), Seperator)
            If temp > 1 Then
                FixedDataArrival Index, Left$(ReceiveBuffers(Index), temp - 1)
            End If
            ReceiveBuffers(Index) = Mid$(ReceiveBuffers(Index), temp + Len(Seperator))
        Loop
    Next Index

End Sub

Private Sub tmrStats_Timer()

Static oldTrafficD As Long

Static oldTrafficU As Long
Dim TrafficD       As Long
Dim TrafficU       As Long
Dim i              As Integer
Dim dpCI           As DPN_CONNECTION_INFO
    'On Error Resume Next
    For i = 0 To UBoundEx(AllUser)
        If Not AllUser(i).VoicePlayerID = -1 Then
            dpCI = dpp.GetConnectionInfo(AllUser(i).VoicePlayerID)
            TrafficD = TrafficD + dpCI.lBytesReceivedGuaranteed + dpCI.lBytesReceivedNonGuaranteed + AllUser(i).dp8BytesReceivedOldSessions + AllUser(i).wscBytesReceived
            TrafficU = TrafficU + dpCI.lBytesSentGuaranteed + dpCI.lBytesSentNonGuaranteed + AllUser(i).dp8BytesSendOldSessions + AllUser(i).wscBytesSend
        Else 'NOT NOT...
            TrafficD = TrafficD + AllUser(i).dp8BytesReceivedOldSessions + AllUser(i).wscBytesReceived
            TrafficU = TrafficU + AllUser(i).dp8BytesSendOldSessions + AllUser(i).wscBytesSend
        End If
        DoEvents
    Next i
    StatusBar.SimpleText = "Gesendet: " & BytesToString(TrafficU) & " (" & Round((TrafficU - oldTrafficU) / 1024 / 30, 2) & " kb/s) Empfangen: " & BytesToString(TrafficD) & " (" & Round((TrafficD - oldTrafficD) / 1024 / 30, 2) & "kb/s)"
    oldTrafficD = TrafficD
    oldTrafficU = TrafficU

End Sub

Private Sub tmrTime_Timer()

Dim i As Integer

    For i = 0 To UBoundEx(AllUser)
        If Not AllUser(i).banTime = 0 Then
            AllUser(i).banTime = AllUser(i).banTime - 1
        End If
    Next i

End Sub

Private Sub tsServer_Click()

Dim i As Integer

    For i = 0 To picFrame.Count - 1
        picFrame(i).Visible = tsServer.SelectedItem.Index = (i + 1)
    Next i
    '    Select Case tsServer.SelectedItem.Index
    '    Case 1
    '        picFrame(2).Visible = False
    '        picFrame(1).Visible = True
    '        picFrame(0).Visible = False
    '    Case 2
    '        picFrame(0).Visible = True
    '        picFrame(1).Visible = False
    '    End Select

End Sub

Private Function UBoundEx(Var() As TUser) As Long

    On Error GoTo errh
    UBoundEx = UBound(Var)

Exit Function

errh:
    UBoundEx = -1

End Function

Private Sub UserLogin(Username As String, _
                      UserPass As String, _
                      Index As Integer)

Dim ArrID As Long
Dim Crc32 As New clsCRC

    ArrID = GetArrayIDFromUsername(Username)
    If Not ArrID = -1 Then
        If AllUser(ArrID).banTime = 0 Then
            If AllUser(ArrID).UserPass = CLng(Crc32.CalculateString(UserPass & Username)) Then
                If AllUser(ArrID).wscID = -1 Then
                    AllUser(ArrID).wscID = Index
                    SendData Index, "login|accept"
                    AddLogEntry Username & " hat sich erfolgreich angemeldet", ServerEvent
                    SendWholeLst Index ' so, volle breitseite für den neu angemeldeten benutzer
                    SendToAll "userlst|" & Username & "|true", Index
                Else 'NOT ALLUSER(ARRID).WSCID...
                    SendData Index, "login|deny"
                    AddLogEntry Username & " wurde die Anmeldung verweigert (Grund: Bereits angemeldet)", ServerEvent
                End If
                ' so, jetzt sagt er allen anderen, dass er online ist
            Else 'NOT ALLUSER(ARRID).USERPASS...
                SendData Index, "login|deny"
                AddLogEntry Username & " wurde die Anmeldung verweigert (Grund: Falsches Passwort)", ServerEvent
            End If
        Else 'NOT ALLUSER(ARRID).BANTIME...
            SendData Index, "login|deny"
            AddLogEntry Username & " wurde die Anmeldung verweigert (Grund: banned)", ServerEvent
        End If
    Else 'NOT NOT...
        SendData Index, "login|deny"
        AddLogEntry Username & " wurde die Anmeldung verweigert (Grund: bereits angemeldet)", ServerEvent
    End If

End Sub

Private Sub wsc_Close(Index As Integer)

    wsc(Index).Close

End Sub

Private Sub wsc_DataArrival(Index As Integer, _
                            ByVal bytesTotal As Long)

' Das EINZIG funktionierende Empfangssystem, welches die messages wirklich trennt
' (c) by Thorben Linneweber !!! :)

Dim strData As String
Dim ArrID   As Integer

    wsc(Index).GetData strData
    ArrID = GetArrayIDFromIndex(Index)
    If Not ArrID = -1 Then
        AllUser(ArrID).wscBytesReceived = AllUser(ArrID).wscBytesReceived + Len(strData)
    End If
    ReceiveBuffers(Index) = ReceiveBuffers(Index) + strData

End Sub

Private Sub wsc_Error(Index As Integer, _
                      ByVal Number As Integer, _
                      Description As String, _
                      ByVal Scode As Long, _
                      ByVal Source As String, _
                      ByVal HelpFile As String, _
                      ByVal HelpContext As Long, _
                      CancelDisplay As Boolean)

    wsc(Index).Close

End Sub

Private Sub wscMain_ConnectionRequest(ByVal requestID As Long)

    On Error Resume Next
    wsc(WinsockCount).Accept requestID
    WinsockCount = WinsockCount + 1
    ReDim Preserve ReceiveBuffers(WinsockCount)
    Load wsc(WinsockCount)

End Sub

':)Code Fixer V3.0.9 (18.07.2005 15:59:13) 200 + 1069 = 1269 Lines Thanks Ulli for inspiration and lots of code.

