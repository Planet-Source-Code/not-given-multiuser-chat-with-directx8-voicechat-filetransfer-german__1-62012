Attribute VB_Name = "modDeclaration"
Option Explicit
' sounds
Public Enum eSounds
    FileOrVoiceRequest
    NudgeSendOrReceived
    UserOnline
    messagesend
End Enum
' Typing
Public Type tTyping
    Username As String
    Typing   As Boolean
End Type
' ~ Save
Public Type tSave
    Sound             As Boolean
    SoundLimited      As Boolean
    Flash             As Boolean
    FlashLimited      As Boolean
    FileTransferPath  As String
    SaveHistory       As Boolean
    Font              As FONT_CONST
    lVoiceRecVol      As Long
    lSoundVol         As Long
    lTriggerVal       As Long
    lQuality          As Long
End Type
Public Type Smiley
    CharCode As String
    rtfCode As String
End Type
Public SmilieArray() As Smiley
'~  Name & Passwort speichern
Public Username                     As String
Public UserPass                     As String
Public ServerIP                     As String
'
' File transfer
Public SendingOrReceivingFile       As Boolean
Public ReceiverOrSender             As String
Public SendFile                     As Boolean
Public PathOfFileToSendOrReceive    As String
Public RemoteIP                     As String
Public SavedOptions As tSave
' frmTransfer
Public bLoadedfrmTransfer As Boolean
Public bLoadedfrmVoice    As Boolean
' frmmain
Public Const EventColor As Long = 5526612
' SignIn Abbrechen
Public SignIn As Boolean
Public Const Seperator   As String = "|||"



