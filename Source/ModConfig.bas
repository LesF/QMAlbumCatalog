Attribute VB_Name = "ModConfig"
Option Explicit
'-----
Public ConfigFolder As String
'-----

' Windows API for INI file access
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpAppName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
     ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
    (ByVal lpAppName As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
     ByVal lpFileName As String) As Long

' Windows API to get AppData path
Private Declare Function SHGetFolderPath Lib "shell32.dll" Alias "SHGetFolderPathA" _
    (ByVal hwnd As Long, ByVal csidl As Long, ByVal hToken As Long, _
     ByVal dwFlags As Long, ByVal pszPath As String) As Long

Private Const CSIDL_APPDATA As Long = &H1A
Private Const MAX_PATH As Long = 260

' Returns: C:\Users\<user>\AppData\Roaming
Public Function GetAppDataFolder() As String
    Dim buffer As String
    buffer = String$(MAX_PATH, vbNullChar)

    If SHGetFolderPath(0, CSIDL_APPDATA, 0, 0, buffer) = 0 Then
        GetAppDataFolder = Left$(buffer, InStr(buffer, vbNullChar) - 1)
    Else
        ' Fallback: current directory
        GetAppDataFolder = App.Path
    End If
End Function

' Returns: C:\Users\<user>\AppData\Roaming\QMClientTest\config.ini
Public Function GetConfigPath() As String
    Dim folder As String
    folder = GetAppDataFolder & "\" & ConfigFolder

    If Dir$(folder, vbDirectory) = vbNullString Then
        MkDir folder
    End If

    GetConfigPath = folder & "\config.ini"
End Function

' Read a value from the INI file
Public Function IniRead(Section As String, Key As String, Optional DefaultValue As String = "") As String
    Dim buffer As String
    buffer = String$(1024, vbNullChar)

    Dim result As Long
    result = GetPrivateProfileString(Section, Key, DefaultValue, buffer, Len(buffer), GetConfigPath)

    IniRead = Left$(buffer, result)
End Function

' Write a value to the INI file
Public Sub IniWrite(Section As String, Key As String, Value As String)
    WritePrivateProfileString Section, Key, Value, GetConfigPath
End Sub


