Attribute VB_Name = "modDServer"
Option Explicit
'Win32 declares
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Enum vbMsgType
    MsgUpdatePlayerLst
    MsgAskToJoin 'We want to ask if we can join this session
    MsgAcceptJoin 'Accept the call
End Enum
'Constants
Private Const AppGuid = "{9073823A-A565-4865-87EC-19B93B014D27}"
Private Const glDefaultPort As Long = 9897
'DirectX variables
Private dx As DirectX8
Public dpp As DirectPlay8Peer
Private dvServer As DirectPlayVoiceServer8
'Misc app variables
Public gsUserName As String
Private glAsyncEnum As Long
Public glMyPlayerID As Long
Private glHostPlayerID As Long

Public Sub Cleanup()

    On Error Resume Next
    'Stop and Destroy the server
    If Not (dvServer Is Nothing) Then
        dvServer.UnRegisterMessageHandler
        dvServer.StopSession 0
        Set dvServer = Nothing
    End If
    'Now the main session
    If Not (dpp Is Nothing) Then
        dpp.UnRegisterMessageHandler
        'Close our peer connection
        dpp.Close
        'Lose references to peer object
        Set dpp = Nothing
    End If
    'Lose references to dx object
    Set dx = Nothing
    DoSleep 500

End Sub

Public Sub DoSleep(ByVal lNumMS As Long)

Dim lCount As Long

    For lCount = 1 To lNumMS / 5
        Sleep 5
        DoEvents
    Next lCount

End Sub

Public Sub InitDPlay()

    Set dx = New DirectX8
    Set dpp = dx.DirectPlayPeerCreate

End Sub

Public Sub StartHosting(MsgForm As Form)

Dim dpa      As DirectPlay8Address
Dim oPlayer  As DPN_PLAYER_INFO
Dim oAppDesc As DPN_APPLICATION_DESC
Dim oSession As DVSESSIONDESC

    'Make sure we're ready to host
    Cleanup
    InitDPlay
    'Register the Message Handler
    dpp.RegisterMessageHandler MsgForm
    'Set the peer info
    oPlayer.lInfoFlags = DPNINFO_NAME
    oPlayer.Name = gsUserName
    dpp.SetPeerInfo oPlayer, DPNOP_SYNC
    'Create an address
    Set dpa = dx.DirectPlayAddressCreate
    'We will only be connecting via TCP/IP
    dpa.SetSP DP8SP_TCPIP
    dpa.AddComponentLong DPN_KEY_PORT, glDefaultPort
    'First set up our application description
    With oAppDesc
        .guidApplication = AppGuid
        .lMaxPlayers = 10 'We don't want to overcrowd our 'room'
        .lFlags = DPNSESSION_NODPNSVR
    End With 'OAPPDESC
    'Start our host
    dpp.Host oAppDesc, dpa
    Set dpa = Nothing
    'After we've created the session and let's start
    'the DplayVoice server
    'Create our DPlayVoice Server
    Set dvServer = dx.DirectPlayVoiceServerCreate
    'Set up the Session
    oSession.lBufferAggressiveness = DVBUFFERAGGRESSIVENESS_DEFAULT
    oSession.lBufferQuality = DVBUFFERQUALITY_DEFAULT
    oSession.lSessionType = DVSESSIONTYPE_PEER
    oSession.guidCT = vbNullString
    'Init and start the session
    dvServer.Initialize dpp, 0
    dvServer.StartSession oSession, 0
    Set dpa = Nothing

End Sub

':)Code Fixer V3.0.9 (18.07.2005 15:59:15) 21 + 90 = 111 Lines Thanks Ulli for inspiration and lots of code.
