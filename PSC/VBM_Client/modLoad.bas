Attribute VB_Name = "modFunctions2"
Option Explicit
' ~ eher spezielle m9 funktionen
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, _
                                                                             ByVal uFlags As Long) As Long

Public Sub FlashAndSound()

    If modDeclaration.SavedOptions.Sound Then
        If modDeclaration.SavedOptions.SoundLimited Then
            If Not HasActiveWindow Then
                PlaySound messagesend
            End If
        Else 'MODDECLARATION.SAVEDOPTIONS.SOUNDLIMITED = FALSE/0
            PlaySound messagesend
        End If
    End If
    If modDeclaration.SavedOptions.Sound Then
        If modDeclaration.SavedOptions.FlashLimited Then
            If Not HasActiveWindow Then
                modFunctions.FlashForm frmMain.hwnd, True
            End If
        Else 'MODDECLARATION.SAVEDOPTIONS.FLASHLIMITED = FALSE/0
            modFunctions.FlashForm frmMain.hwnd, True
        End If
    End If

End Sub

Public Sub LoadOptions()

    With modDeclaration.SavedOptions
        .Sound = IIf(GetSetting(App.Title, "Options", "Sound", "true") = "true", True, False)
        .SoundLimited = IIf(GetSetting(App.Title, "Options", "SoundLimited", "true") = "true", True, False)
        .Flash = IIf(GetSetting(App.Title, "Options", "Flash", "true") = "true", True, False)
        .FlashLimited = IIf(GetSetting(App.Title, "Options", "FlashLimited", "true") = "true", True, False)
        .FileTransferPath = GetSetting(App.Title, "Options", "FilePath", "")
        .SaveHistory = IIf(GetSetting(App.Title, "Options", "SaveHistory", "true") = "true", True, False)
        ' Chatfenster Font
        .Font.FontBold = IIf(GetSetting(App.Title, "Font", "Bold", "false") = "true", True, False)
        .Font.FontColor = GetSetting(App.Title, "Font", "Color", "000") 'TODO
        .Font.FontItalic = IIf(GetSetting(App.Title, "Font", "Italic", "false") = "true", True, False)
        .Font.FontName = GetSetting(App.Title, "Font", "Name", "Arial")
        .Font.FontSize = GetSetting(App.Title, "Font", "Size", "10")
        .Font.FontStrikeThru = IIf(GetSetting(App.Title, "Font", "StrikeThru", "false") = "true", True, False)
        .Font.FontUnderline = IIf(GetSetting(App.Title, "Font", "Underline", "false") = "true", True, False)
        ' VoiceChat
        .lTriggerVal = GetSetting(App.Title, "Voice", "Trigger", "10")
        .lVoiceRecVol = GetSetting(App.Title, "Voice", "RecVol", "0")
        .lSoundVol = GetSetting(App.Title, "Voice", "SoundVol", "0")
        .lQuality = GetSetting(App.Title, "Voice", "Quality", "80")
    End With 'MODDECLARATION.SAVEDOPTIONS

End Sub

Public Sub PlaySound(Sound As eSounds)

Dim Filename As String

    Filename = AppPath & "\Sounds\"
    Select Case Sound
    Case FileOrVoiceRequest
        Filename = Filename & "file_voice.wav"
    Case NudgeSendOrReceived
        Filename = Filename & "nudge.wav"
    Case UserOnline
        Filename = Filename & "online.wav"
    Case messagesend
        Filename = Filename & "send.wav"
    End Select
    If FileExists(Filename) Then
        DoEvents
        Call sndPlaySound(Filename, &H1)
        DoEvents
    End If

End Sub

Public Sub SaveOptions()

' Optionen

    SaveSetting App.Title, "Options", "Sound", IIf(modDeclaration.SavedOptions.Sound, "true", "false")
    SaveSetting App.Title, "Options", "SoundLimited", IIf(modDeclaration.SavedOptions.SoundLimited, "true", "false")
    SaveSetting App.Title, "Options", "Flash", IIf(modDeclaration.SavedOptions.Flash, "true", "false")
    SaveSetting App.Title, "Options", "FlashLimited", IIf(modDeclaration.SavedOptions.FlashLimited, "true", "false")
    SaveSetting App.Title, "Options", "FilePath", modDeclaration.SavedOptions.FileTransferPath
    SaveSetting App.Title, "Options", "SaveHistory", IIf(modDeclaration.SavedOptions.SaveHistory, "true", "false")
    ' Chatfenster Font
    SaveSetting App.Title, "Font", "Bold", IIf(modDeclaration.SavedOptions.Font.FontBold, "true", "false")
    SaveSetting App.Title, "Font", "Color", modDeclaration.SavedOptions.Font.FontColor
    SaveSetting App.Title, "Font", "Italic", IIf(modDeclaration.SavedOptions.Font.FontItalic, "true", "false")
    SaveSetting App.Title, "Font", "Name", modDeclaration.SavedOptions.Font.FontName
    SaveSetting App.Title, "Font", "Size", modDeclaration.SavedOptions.Font.FontSize
    SaveSetting App.Title, "Font", "StrikeThru", IIf(modDeclaration.SavedOptions.Font.FontStrikeThru, "true", "false")
    SaveSetting App.Title, "Font", "Underline", IIf(modDeclaration.SavedOptions.Font.FontUnderline, "true", "false")
    ' VoiceChat
    SaveSetting App.Title, "Voice", "Trigger", modDeclaration.SavedOptions.lTriggerVal
    SaveSetting App.Title, "Voice", "RecVol", modDeclaration.SavedOptions.lVoiceRecVol
    SaveSetting App.Title, "Voice", "SoundVol", modDeclaration.SavedOptions.lSoundVol
    SaveSetting App.Title, "Voice", "Quality", modDeclaration.SavedOptions.lQuality

End Sub



