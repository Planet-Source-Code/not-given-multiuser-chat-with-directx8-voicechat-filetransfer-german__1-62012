VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmVoice 
   BackColor       =   &H00E0E0E0&
   Caption         =   "VoiceChat"
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2400
   Icon            =   "frmVoice.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   317
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   160
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox picContainer 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      ScaleHeight     =   735
      ScaleWidth      =   2175
      TabIndex        =   6
      Top             =   960
      Width           =   2175
      Begin VB.Label lblConnecting 
         BackStyle       =   0  'Transparent
         Caption         =   "Verbinde..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A85E33&
         Height          =   240
         Index           =   0
         Left            =   840
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.Image imgDX8 
         Height          =   480
         Left            =   120
         Picture         =   "frmVoice.frx":058A
         Top             =   120
         Width           =   480
      End
      Begin VB.Shape shpBorder 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Innen ausgefüllt
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Ausgefüllt
         Height          =   735
         Left            =   0
         Top             =   0
         Width           =   2160
      End
   End
   Begin VB.Frame frameControls 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'Kein
      Caption         =   "Frame1"
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   3600
      Width           =   2415
      Begin ComctlLib.ProgressBar pbPlayVol 
         Height          =   165
         Left            =   480
         TabIndex        =   2
         Top             =   420
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   291
         _Version        =   327682
         Appearance      =   1
         Max             =   99
      End
      Begin ComctlLib.ProgressBar PBRecVol 
         Height          =   165
         Left            =   480
         TabIndex        =   3
         Top             =   120
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   291
         _Version        =   327682
         Appearance      =   1
         Max             =   99
      End
      Begin VBMessenger9.isButton cmdCancel 
         Height          =   300
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   529
         Style           =   7
         Caption         =   "Verlassen"
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
      Begin VBMessenger9.isButton cmdOptions 
         Height          =   300
         Left            =   1200
         TabIndex        =   5
         Top             =   720
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   529
         Style           =   7
         Caption         =   "Optionen"
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
      Begin VB.Image Image1 
         Height          =   165
         Left            =   120
         Picture         =   "frmVoice.frx":11CE
         Top             =   120
         Width           =   180
      End
      Begin VB.Image Image2 
         Height          =   165
         Left            =   120
         Picture         =   "frmVoice.frx":139C
         Top             =   405
         Width           =   135
      End
   End
   Begin VB.Timer tmrVoice 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1200
      Top             =   7320
   End
   Begin VB.Timer tmrNoConnection 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1680
      Top             =   7320
   End
   Begin ComctlLib.ListView lvVoice 
      Height          =   3600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   6350
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   327682
      SmallIcons      =   "imgLst"
      ForeColor       =   11034163
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   3969
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   ""
         Object.Width           =   0
      EndProperty
   End
   Begin ComctlLib.ImageList imgLst 
      Left            =   1800
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmVoice.frx":1512
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmVoice.frx":1864
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmVoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ~ Voice
Implements DirectPlay8Event
Implements DirectPlayVoiceEvent8
'Misc private variables
Private moCallBack As DirectPlay8Event
Private mlVoiceError As Long

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOptions_Click()

    frmOptions.Show

End Sub

Private Sub DirectPlay8Event_AddRemovePlayerGroup(ByVal lMsgID As Long, _
                                                  ByVal lPlayerID As Long, _
                                                  ByVal lGroupID As Long, _
                                                  fRejectMsg As Boolean)

'VB requires that we must implement *every* member of this interface

    If (Not moCallBack Is Nothing) Then moCallBack.AddRemovePlayerGroup lMsgID, lPlayerID, lGroupID, fRejectMsg

End Sub

Private Sub DirectPlay8Event_AppDesc(fRejectMsg As Boolean)

'VB requires that we must implement *every* member of this interface

    If (Not moCallBack Is Nothing) Then moCallBack.AppDesc fRejectMsg

End Sub

Private Sub DirectPlay8Event_AsyncOpComplete(dpnotify As DxVBLibA.DPNMSG_ASYNC_OP_COMPLETE, _
                                             fRejectMsg As Boolean)

'VB requires that we must implement *every* member of this interface

    If (Not moCallBack Is Nothing) Then moCallBack.AsyncOpComplete dpnotify, fRejectMsg

End Sub

Private Sub DirectPlay8Event_ConnectComplete(dpnotify As DxVBLibA.DPNMSG_CONNECT_COMPLETE, _
                                             fRejectMsg As Boolean)

Dim lMsg   As Long, lOffset As Long
Dim oBuf() As Byte

    If dpnotify.hResultCode = 0 Then 'Success!
        'Now let's send a message asking the host to accept our call
        lOffset = NewBuffer(oBuf)
        lMsg = MsgAskToJoin
        AddDataToBuffer oBuf, lMsg, LenB(lMsg), lOffset
        dpp.SendTo DPNID_ALL_PLAYERS_GROUP, oBuf, 0, DPNSEND_NOLOOPBACK
    Else 'NOT DPNOTIFY.HRESULTCODE...
        tmrNoConnection.Enabled = True
    End If
    'VB requires that we must implement *every* member of this interface
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
    'VB requires that we must implement *every* member of this interface
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

Dim dpPeer As DPN_PLAYER_INFO

    On Error Resume Next
    If lPlayerID <> glMyPlayerID Then 'ignore removing myself
        RemovePlayer lPlayerID
    End If
    'If Not (ChatWindow Is Nothing) Then Set moCallBack = ChatWindow 'If the chat window is open, let them know about the departure.
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

Dim lMsg   As Long, lOffset As Long
Dim dpPeer As DPN_PLAYER_INFO

    With dpnotify
        GetDataFromBuffer .ReceivedData, lMsg, LenB(lMsg), lOffset
        Select Case lMsg
        Case MsgAcceptJoin
            picContainer.Visible = False
            UpdatePlayerList
            ConnectVoice Me
        Case MsgUpdatePlayerLst
            UpdatePlayerList 'Update our list here
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
    tmrNoConnection.Enabled = True

End Sub

Private Sub DirectPlayVoiceEvent8_ConnectResult(ByVal ResultCode As Long)

Dim lTargets(0) As Long

    lTargets(0) = DVID_ALLPLAYERS
    On Error Resume Next
    'Connect the client
    dvClient.SetTransmitTargets lTargets, 0
    If Err.Number <> 0 And Err.Number <> DVERR_PENDING Then
        mlVoiceError = Err.Number
        tmrVoice.Enabled = True
        Exit Sub
    End If

End Sub

Private Sub DirectPlayVoiceEvent8_CreateVoicePlayer(ByVal PlayerID As Long, _
                                                    ByVal flags As Long)

'VB requires that we must implement *every* member of this interface


End Sub

Private Sub DirectPlayVoiceEvent8_DeleteVoicePlayer(ByVal PlayerID As Long)

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

    pbPlayVol.Value = PeakLevel

End Sub

Private Sub DirectPlayVoiceEvent8_OutputLevel(ByVal PeakLevel As Long, _
                                              ByVal OutputVolume As Long)

'VB requires that we must implement *every* member of this interface

    PBRecVol.Value = PeakLevel

End Sub

Private Sub DirectPlayVoiceEvent8_PlayerOutputLevel(ByVal SourcePlayerID As Long, _
                                                    ByVal PeakLevel As Long)

'VB requires that we must implement *every* member of this interface


End Sub

Private Sub DirectPlayVoiceEvent8_PlayerVoiceStart(ByVal SourcePlayerID As Long)

    SetPlayerStatus True, SourcePlayerID

End Sub

Private Sub DirectPlayVoiceEvent8_PlayerVoiceStop(ByVal SourcePlayerID As Long)

    SetPlayerStatus False, SourcePlayerID

End Sub

Private Sub DirectPlayVoiceEvent8_RecordStart(ByVal PeakVolume As Long)

    SetPlayerStatus True, glMyPlayerID
    'VB requires that we must implement *every* member of this interface

End Sub

Private Sub DirectPlayVoiceEvent8_RecordStop(ByVal PeakVolume As Long)

    SetPlayerStatus False, glMyPlayerID
    'VB requires that we must implement *every* member of this interface

End Sub

Private Sub DirectPlayVoiceEvent8_SessionLost(ByVal ResultCode As Long)

'VB requires that we must implement *every* member of this interface

    tmrNoConnection.Enabled = True

End Sub

Private Sub Form_Load()

    bLoadedfrmVoice = True
    modDplay.gsUserName = modDeclaration.Username
    Connect Me, modDeclaration.ServerIP
    modFunctions.FullRowSelect lvVoice
    If Not modFunctions.IsIDE Then
        With modResize2.spR
            .xMin = frmVoice.Width / Screen.TwipsPerPixelX
            .yMin = 200
            .xMax = frmVoice.Width / Screen.TwipsPerPixelX
            .yMax = Screen.Height / Screen.TwipsPerPixelY
        End With 'MODRESIZE2.SPR
        Call InitR(frmVoice)
        frmMain.oMagneticWnd.AddWindow hwnd, frmMain.hwnd
    End If

End Sub

Private Sub Form_Resize()

    lvVoice.Height = frmVoice.ScaleHeight - frameControls.Height
    frameControls.Top = frmVoice.ScaleHeight - frameControls.Height
    DoEvents

End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    'erstmal voice ausschalten
    If Not (dvClient Is Nothing) Then dvClient.Disconnect 0
   ' If Not (dvClient Is Nothing) Then dvClient.UnRegisterMessageHandler
    'dann entladen
    modDplay.Cleanup
    frmMain.oMagneticWnd.RemoveWindow hwnd
    UnHookR
    bLoadedfrmVoice = False
    'wichtig...
    Set frmVoice = Nothing

End Sub

Private Sub RemovePlayer(ByVal lPlayerID As Long)

Dim lCount As Long

    For lCount = 1 To lvVoice.ListItems.Count
        If lvVoice.ListItems.Item(lCount).SubItems(1) = lPlayerID Then lvVoice.ListItems.Remove (lCount)
    Next lCount

End Sub

Private Sub SetPlayerStatus(Speaking As Boolean, _
                            PlayerID As Long)

Dim i As Integer

    For i = 1 To lvVoice.ListItems.Count
        If lvVoice.ListItems.Item(i).SubItems(1) = PlayerID Then
            lvVoice.ListItems.Item(i).SmallIcon = IIf(Speaking, 2, 1)
            lvVoice.Refresh
            Exit For
        End If
    Next i

End Sub

Private Sub tmrNoConnection_Timer()

    tmrNoConnection.Enabled = False
    MsgBox "Die Verbindung wurde getrennt.", vbOKOnly Or vbInformation
    Unload Me

End Sub

Private Sub tmrVoice_Timer()

    tmrVoice.Enabled = False
    MsgBox "Could not start DirectPlayVoice" & vbNewLine & _
       "Error:" & CStr(mlVoiceError), vbOKOnly Or vbInformation, "No Voice"
    gfNoVoice = True
    tmrNoConnection.Enabled = True

End Sub

Public Sub UpdatePlayerList()

Dim lCount As Long, dpPeer As DPN_PLAYER_INFO
Dim lInner As Long, fFound As Boolean
Dim lTotal As Long

    lTotal = dpp.GetCountPlayersAndGroups(DPNENUM_PLAYERS)
    For lCount = 1 To lTotal
        dpPeer = dpp.GetPeerInfo(dpp.GetPlayerOrGroup(lCount))
        If Not (dpPeer.lPlayerFlags = DPNPLAYER_HOST) Then
            fFound = False
            'Make sure they're not already added
            For lInner = 1 To lvVoice.ListItems.Count
                If lvVoice.ListItems.Item(lInner).Text = CStr(dpPeer.Name) Then
                    fFound = True
                    Exit For
                End If
            Next lInner
            If Not fFound Then
                'Go ahead and add them
                lvVoice.ListItems.Add , , dpPeer.Name, , 1
                lvVoice.ListItems.Item(lvVoice.ListItems.Count).SubItems(1) = dpp.GetPlayerOrGroup(lCount)
            End If
        End If
    Next lCount

End Sub



