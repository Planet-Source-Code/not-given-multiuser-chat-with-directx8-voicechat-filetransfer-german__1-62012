Attribute VB_Name = "modDplay"
Option Explicit
Public Enum vbMsgType
    MsgUpdatePlayerLst
    MsgAskToJoin 'We want to ask if we can join this session
    MsgAcceptJoin 'Accept the call
End Enum
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'Constants
Private Const AppGuid = "{9073823A-A565-4865-87EC-19B93B014D27}"
Private Const glDefaultPort As Long = 9897
'DirectX variables
Public dx As DirectX8
Public dpp As DirectPlay8Peer
Public dvClient As DirectPlayVoiceClient8
Public gsUserName As String
Private glAsyncEnum As Long
Public glMyPlayerID As Long
Private glHostPlayerID As Long
Private gfHost As Boolean
Public gfNoVoice As Boolean

Public Sub Cleanup()

    On Error Resume Next
    If Not (dvClient Is Nothing) Then
        dvClient.UnRegisterMessageHandler
        dvClient.Disconnect DVFLAGS_SYNC
        Set dvClient = Nothing
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

Public Sub Connect(MsgForm As Form, _
                   ByVal sHost As String)

Dim dpa      As DirectPlay8Address
Dim dpl      As DirectPlay8Address
Dim oPlayer  As DPN_PLAYER_INFO
Dim oAppDesc As DPN_APPLICATION_DESC

    'Try to connect to the host
    'Make sure we're ready to connect
    Cleanup
    InitDPlay
    gfHost = False
    'Register the Message Handler
    dpp.RegisterMessageHandler MsgForm
    'Set the peer info
    oPlayer.lInfoFlags = DPNINFO_NAME
    oPlayer.Name = gsUserName
    dpp.SetPeerInfo oPlayer, DPNOP_SYNC
    'Now try to enum hosts
    'Create an address
    Set dpa = dx.DirectPlayAddressCreate
    'We will only be connecting via TCP/IP
    dpa.SetSP DP8SP_TCPIP

    dpa.AddComponentString DPN_KEY_HOSTNAME, sHost 'We will try to connect to this host
    dpa.AddComponentLong DPN_KEY_PORT, glDefaultPort
    Set dpl = dx.DirectPlayAddressCreate
    'We will only be connecting via TCP/IP
    dpl.SetSP DP8SP_TCPIP
    'First set up our application description
    With oAppDesc
        .guidApplication = AppGuid
    End With 'OAPPDESC
    'Try to connect to this host
    On Error Resume Next
    DoSleep 500 'Give a slight pause to clean up any loose ends
    dpp.Connect oAppDesc, dpa, dpl, 0, ByVal 0&, 0
    If Err.Number <> 0 Then 'Woah, an error
        MsgBox "There was an error trying to connect to this machine.", vbOKOnly Or vbInformation, "Unavailable"
    End If
    Set dpa = Nothing
    Set dpl = Nothing

End Sub

Public Sub ConnectVoice(MsgForm As Form)

Dim oSound  As DVSOUNDDEVICECONFIG
Dim oClient As DVCLIENTCONFIG
Dim dvSetup As DirectPlayVoiceTest8

    'Make sure we haven't determined there would be no voice in this app
    If gfNoVoice Then Exit Sub
    'Now create a client as well (so we can both talk and listen)
    Set dvClient = dx.DirectPlayVoiceClientCreate
    'Now let's create a client event..
    dvClient.Initialize dpp, 0
    dvClient.StartClientNotification MsgForm
    'Set up our client and sound structs
    'geändert von thorben
    ''''oClient.lFlags = DVCLIENTCONFIG_AUTOVOICEACTIVATED Or DVCLIENTCONFIG_AUTORECORDVOLUME
    oClient.lFlags = DVCLIENTCONFIG_MANUALVOICEACTIVATED
    oClient.lBufferAggressiveness = DVBUFFERAGGRESSIVENESS_DEFAULT
    oClient.lBufferQuality = modDeclaration.SavedOptions.lQuality
    oClient.lNotifyPeriod = DVNOTIFYPERIOD_MINPERIOD
    'geändert von thorben
    '''''oClient.lThreshold = DVTHRESHOLD_UNUSED
    oClient.lThreshold = modDeclaration.SavedOptions.lTriggerVal
    oClient.lRecordVolume = modDeclaration.SavedOptions.lVoiceRecVol
    oClient.lPlaybackVolume = modDeclaration.SavedOptions.lSoundVol
    '''CONST_DVTHRESHOLD
    '''~~~~~~~~~~~~~~~~
    '''Used in the lThreshold member of the DVCLIENTCONFIG type to specify the input level used to trigger voice transmission if the DVCLIENTCONFIG_MANUALVOICEACTIVATED flag is specified in the lFlags member. When the flag is specified, this value can be set to anywhere in the range of DVTHRESHOLD_MIN to DVTHRESHOLD_MAX. Additionally, DVTHRESHOLD_DEFAULT can be set to use a default value.
    '''
    '''If DVCLIENTCONFIG_MANUALVOICEACTIVATED or DVCLIENTCONFIG_AUTOVOICEACTIVATED is not specified in the lFlags member of this structure (indicating push-to-talk mode), this value must be set to DVTHRESHOLD_UNUSED.
    '''
    '''Enum CONST_DVTHRESHOLD
    '''    DVTHRESHOLD_DEFAULT = -1 (&HFFFFFFFF)
    '''    DVTHRESHOLD_MAX = 99 (&H63)
    '''    DVTHRESHOLD_MIN = 0
    '''    DVTHRESHOLD_UNUSED = -2 (&HFFFFFFFE)
    '''End Enum
    '''
    '''Constants
    '''
    '''DVTHRESHOLD_DEFAULT
    '''    Default threshold value.
    '''DVTHRESHOLD_MAX
    '''    Maximum threshold value.
    '''DVTHRESHOLD_MIN
    '''    Minimum threshold value.
    '''DVTHRESHOLD_UNUSED
    '''    Must be set for push-to-talk mode.
    '''
    oSound.hwndAppWindow = frmVoice.hwnd
    '   On Error Resume Next
    'Connect the client
    dvClient.Connect oSound, oClient, 0
    If Err.Number = DVERR_RUN_SETUP Then    'The audio tests have not been run on this
        'machine.  Run them now.
        'we need to run setup first
        Set dvSetup = dx.DirectPlayVoiceTestCreate
        dvSetup.CheckAudioSetup vbNullString, vbNullString, frmVoice.hwnd, 0
         'Check the default devices since that's what we'll be using
        If Err.Number = DVERR_COMMANDALREADYPENDING Then
            MsgBox "Could not start DirectPlayVoice.  The Voice Networking wizard is already open.", vbOKOnly Or vbInformation, "No Voice"
            gfNoVoice = True
            ' Form1.chkVoice.Value = vbUnchecked
            '  Form1.chkVoice.Enabled = False
            Exit Sub
        End If
        If Err.Number = DVERR_USERCANCEL Then
            MsgBox "Could not start DirectPlayVoice.  The Voice Networking wizard was cancelled.", vbOKOnly Or vbInformation, "No Voice"
            gfNoVoice = True
            ' NetWorkForm.chkVoice.Value = vbUnchecked
            '  NetWorkForm.chkVoice.Enabled = False
            Exit Sub
        End If
        Set dvSetup = Nothing
        dvClient.Connect oSound, oClient, 0
    ElseIf Err.Number <> 0 And Err.Number <> DVERR_PENDING Then 'NOT ERR.NUMBER...
        MsgBox "Could not start DirectPlayVoice." & vbNewLine & _
       "Error:" & CStr(Err.Number) & Err.Description, vbOKOnly Or vbInformation, "No Voice"
        gfNoVoice = True
        '  NetWorkForm.chkVoice.Value = vbUnchecked
        ' NetWorkForm.chkVoice.Enabled = False
        Exit Sub
    End If
    On Error GoTo 0

End Sub

Public Sub DoSleep(ByVal lNumMS As Long)

Dim lCount As Long

    For lCount = 1 To lNumMS \ 5
        Sleep 5
        DoEvents
    Next lCount

End Sub

Public Sub InitDPlay()

    Set dx = New DirectX8
    Set dpp = dx.DirectPlayPeerCreate

End Sub

Public Sub RunAudioAssistant(hwnd As Long)

Dim dvSetup As DirectPlayVoiceTest8

    On Error Resume Next
    InitDPlay
    Set dvSetup = dx.DirectPlayVoiceTestCreate
    dvSetup.CheckAudioSetup vbNullString, vbNullString, hwnd, 0
     'Check the default devices since that's what we'll be using

End Sub



