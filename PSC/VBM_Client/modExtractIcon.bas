Attribute VB_Name = "modExtractIcon"
Option Explicit
Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (pDicDesc As IconType, _
                                                                      riid As CLSIdType, _
                                                                      ByVal fown As Long, _
                                                                      lpUnk As Object) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, _
                                                                                 ByVal dwFileAttributes As Long, _
                                                                                 psfi As ShellFileInfoType, _
                                                                                 ByVal cbFileInfo As Long, _
                                                                                 ByVal uFlags As Long) As Long
Public Const Large As Long = &H100&
Private Const Small As Long = &H101&
Private Const MAXPATH As Long = 260&
Private Type IconType
    cbSize As Long
    picType As PictureTypeConstants
    hIcon As Long
End Type
Private Type CLSIdType
    id(16) As Byte
End Type
Private Type ShellFileInfoType
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAXPATH
    szTypeName As String * 80
End Type

Public Function LoadIcon(Size As Long, _
                         Extension As String) As IPictureDisp

Dim Result    As Long
Dim File      As String, Slash As String
Dim Unknown   As IUnknown
Dim Icon      As IconType
Dim CLSID     As CLSIdType
Dim ShellInfo As ShellFileInfoType
Dim FF%: FF = FreeFile

    File = Environ$("TEMP")
    If Len(File) > 3 Then File = File & "\"
    File = File & "Cache." & Extension
    Open File For Binary As FF
    Put #FF, , 255
    Close FF
    Call SHGetFileInfo(File, 0, ShellInfo, Len(ShellInfo), Size)
    Icon.cbSize = Len(Icon)
    Icon.picType = vbPicTypeIcon
    Icon.hIcon = ShellInfo.hIcon
    CLSID.id(8) = &HC0
    CLSID.id(15) = &H46
    Result = OleCreatePictureIndirect(Icon, CLSID, 1, Unknown)
    Set LoadIcon = Unknown
    Kill File

End Function



