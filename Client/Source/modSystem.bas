Attribute VB_Name = "modSystem"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetTickCount Lib "KERNEL32" () As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal H As Long, ByVal hb As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal f As Long) As Long
Public Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)

' Used with DirectX (for text output)
Public Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
Public Declare Function GetDC Lib "USER32.DLL" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowRect Lib "USER32.DLL" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function ReleaseDC Lib "USER32.DLL" (ByVal hWnd As Long, ByVal hDC As Long) As Long

' Used with ReadINI/WriteINI.
Public Declare Function WritePrivateProfileString& Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)
Public Declare Function GetPrivateProfileString& Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal ReturnedString$, ByVal RSSize&, ByVal FileName$)

Public Function FileExists(ByVal FileName As String) As Boolean
    On Error Resume Next

    FileExists = (GetAttr(App.Path & "\" & FileName) And vbDirectory) = 0
End Function

' This is the new FileExists. I'm slowly switching it over. [Mellowz]
' When I'm done switching them, this will become the new FileExists.
Public Function FileExistsNew(ByVal FileName As String) As Boolean
    On Error Resume Next

    FileExistsNew = (GetAttr(FileName) And vbDirectory) = 0
End Function

Public Function FolderExists(ByVal FilePath As String) As Boolean
    On Error Resume Next

    If LenB(Dir(FilePath, vbDirectory)) <> 0 Then
        FolderExists = True
    End If
End Function

Public Sub WriteINI(INISection As String, INIKey As String, INIValue As String, INIFile As String)
    Call WritePrivateProfileString(INISection, INIKey, INIValue, INIFile)
End Sub

Public Function ReadINI(INISection As String, INIKey As String, INIFile As String) As String
    Dim StringBuffer As String
    Dim StringBufferSize As Long

    StringBuffer = Space$(255)
    StringBufferSize = Len(StringBuffer)

    StringBufferSize = GetPrivateProfileString(INISection, INIKey, vbNullString, StringBuffer, StringBufferSize, INIFile)

    If StringBufferSize > 0 Then
        ReadINI = Left$(StringBuffer, StringBufferSize)
    Else
        ReadINI = vbNullString
    End If
End Function

Public Sub ListMusic(ByRef List As ListBox, ByVal sStartDir As String)
    Call ListBox_GetFilesInFolder(List, sStartDir & "*.mid")
End Sub

Public Sub ListSounds(ByRef List As ListBox, ByVal sStartDir As String)
    Call ListBox_GetFilesInFolder(List, sStartDir & "*.wav")
End Sub

Public Sub MovePicture(PB As PictureBox, Button As Integer, Shift As Integer, X As Single, y As Single)
    Dim GlobalX As Integer
    Dim GlobalY As Integer

    GlobalX = PB.Left
    GlobalY = PB.Top

    If Button = 1 Then
        PB.Left = GlobalX + X - SOffsetX
        PB.Top = GlobalY + y - SOffsetY
    End If
End Sub

Public Sub ListBox_GetFilesInFolder(ByRef List As ListBox, ByVal FilePath As String)
    Dim File As String
    
    ' Get the first file in the directory.
    File = Dir$(FilePath)

    ' Loop through all of the files.
    Do While LenB(File) <> 0

        ' Add the first file to the listbox.
        List.addItem (File)
        
        ' Get the next file in the directory.
        File = Dir$
    Loop
End Sub
