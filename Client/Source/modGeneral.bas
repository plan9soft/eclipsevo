Attribute VB_Name = "modGeneral"
Option Explicit

Public Sub Main()
    ' Check if we're debugging the application.
    If FileExists("debug") Then
        frmDebug.Visible = True
    End If

    ' Display the status form.
    frmSendGetData.Visible = True

    ' Check to make sure all the folder exist.
    Call SetStatus("Checking Folders...")
    Call CheckFolders

    ' Check to make sure all the files exist.
    Call SetStatus("Checking Files...")
    Call CheckFiles
    
    ' Initialize global variables.
    LAST_DIR = -1

    ' Load the configuration settings.
    Call SetStatus("Loading Configuration...")
    Call LoadConfig
    Call LoadColors
    Call LoadFont

    ' Prepare the socket for communication.
    Call SetStatus("Preparing Socket...")
    Call TCPInit
    
    ' Load the DirectX7 object.
    Call SetStatus("Initializing DirectX...")
    Call InitDirectX
    Call DirectMusic_Init
    Call DirectSound_Init

    ' Display the version the user is using on the client.
    frmMainMenu.lblVersion.Caption = "Version: " & App.Major & "." & App.Minor

    ' Hide the status form and display the main form.
    frmSendGetData.Visible = False
    frmMainMenu.Visible = True
End Sub

Public Sub CheckFolders()
    ' Check if the 'Maps' folder exists.
    If Not FolderExists(App.Path & "\Maps") Then
        Call MkDir$(App.Path & "\Maps")
    End If

    ' Check if the 'GFX' folder exists.
    If Not FolderExists(App.Path & "\GFX") Then
        Call MkDir$(App.Path & "\GFX")
    End If

    ' Check if the 'GUI' folder exists.
    If Not FolderExists(App.Path & "\GUI") Then
        Call MkDir$(App.Path & "\GUI")
    End If

    ' Check if the 'Music' folder exists.
    If Not FolderExists(App.Path & "\Music") Then
        Call MkDir$(App.Path & "\Music")
    End If

    ' Check if the 'SFX' folder exists.
    If Not FolderExists(App.Path & "\SFX") Then
        Call MkDir$(App.Path & "\SFX")
    End If

    ' Check if the 'Flashs' folder exists.
    If Not FolderExists(App.Path & "\Flashs") Then
        Call MkDir$(App.Path & "\Flashs")
    End If

    ' Check if the 'BGS' folder exists.
    If Not FolderExists(App.Path & "\BGS") Then
        Call MkDir$(App.Path & "\BGS")
    End If

    ' Check if the 'Data' folder exists.
    If Not FolderExists(App.Path & "\Data") Then
        Call MkDir$(App.Path & "\Data")
    End If
End Sub

Public Sub CheckFiles()
    ' Check if the 'Config.ini' file exists.
    If Not FileExists("Config.ini") Then
        Call FileCreateConfigINI
    End If

    ' Check if the 'News.ini' file exists.
    If Not FileExists("News.ini") Then
        Call FileCreateNewsINI
    End If

    ' Check if the 'Font.ini' file exists.
    If Not FileExists("Font.ini") Then
        Call FileCreateFontINI
    End If

    ' Check if the 'GUI\Colors.ini' file exists.
    If Not FileExists("GUI\Colors.txt") Then
        Call FileCreateColorsTXT
    End If
End Sub

Public Sub LoadConfig()
    Dim FileName As String

    ' Check for errors when loading data.
    On Error GoTo ErrorHandle

    ' Cache the file path.
    FileName = App.Path & "\Config.ini"

    ' Load the in-game configuration options.
    frmMirage.chkBubbleBar.Value = CLng(ReadINI("CONFIG", "SpeechBubbles", FileName))
    frmMirage.chkNpcBar.Value = CLng(ReadINI("CONFIG", "NpcBar", FileName))
    frmMirage.chkNpcName.Value = CLng(ReadINI("CONFIG", "NPCName", FileName))
    frmMirage.chkPlayerBar.Value = CLng(ReadINI("CONFIG", "PlayerBar", FileName))
    frmMirage.chkPlayerName.Value = CLng(ReadINI("CONFIG", "PlayerName", FileName))
    frmMirage.chkPlayerDamage.Value = CLng(ReadINI("CONFIG", "NPCDamage", FileName))
    frmMirage.chkNpcDamage.Value = CLng(ReadINI("CONFIG", "PlayerDamage", FileName))
    frmMirage.chkMusic.Value = CLng(ReadINI("CONFIG", "Music", FileName))
    frmMirage.chkSound.Value = CLng(ReadINI("CONFIG", "Sound", FileName))
    frmMirage.chkAutoScroll.Value = CLng(ReadINI("CONFIG", "AutoScroll", FileName))
    frmMirage.chkSwearFilter.Value = CLng(ReadINI("CONFIG", "SwearFilter", FileName))
    frmMirage.chkNightEffect.Value = CLng(ReadINI("CONFIG", "NightEffect", FileName))

    Exit Sub

ErrorHandle:
    ' Delete the existing file and create a new one.
    Call MsgBox("Failed to load the file: Config.ini.")
    Call Kill(App.Path & "\Config.ini")
    Call FileCreateConfigINI
End Sub

Public Sub FileCreateConfigINI()
    ' Create the key names and values in the configuration file.
    Call WriteINI("IPCONFIG", "IP", "127.0.0.1", App.Path & "\Config.ini")
    Call WriteINI("IPCONFIG", "PORT", CStr(4001), App.Path & "\Config.ini")
    Call WriteINI("CONFIG", "Account", vbNullString, App.Path & "\Config.ini")
    Call WriteINI("CONFIG", "Password", vbNullString, App.Path & "\Config.ini")
    Call WriteINI("CONFIG", "WebSite", vbNullString, App.Path & "\Config.ini")
    Call WriteINI("CONFIG", "SpeechBubbles", CStr(1), App.Path & "\Config.ini")
    Call WriteINI("CONFIG", "NpcBar", CStr(1), App.Path & "\Config.ini")
    Call WriteINI("CONFIG", "NPCName", CStr(1), App.Path & "\Config.ini")
    Call WriteINI("CONFIG", "NPCDamage", CStr(1), App.Path & "\Config.ini")
    Call WriteINI("CONFIG", "PlayerBar", CStr(1), App.Path & "\Config.ini")
    Call WriteINI("CONFIG", "PlayerName", CStr(1), App.Path & "\Config.ini")
    Call WriteINI("CONFIG", "PlayerDamage", CStr(1), App.Path & "\Config.ini")
    Call WriteINI("CONFIG", "MapGrid", CStr(1), App.Path & "\Config.ini")
    Call WriteINI("CONFIG", "Music", CStr(1), App.Path & "\Config.ini")
    Call WriteINI("CONFIG", "Sound", CStr(1), App.Path & "\Config.ini")
    Call WriteINI("CONFIG", "AutoScroll", CStr(1), App.Path & "\Config.ini")
    Call WriteINI("CONFIG", "Auto", CStr(0), App.Path & "\Config.ini")
    Call WriteINI("CONFIG", "NightEffect", CStr(1), App.Path & "\Config.ini")
    Call WriteINI("CONFIG", "SwearFilter", CStr(0), App.Path & "\Config.ini")
End Sub

Public Sub FileCreateNewsINI()
    ' Create the key names and values for the title and description.
    Call WriteINI("DATA", "News", vbNullString, App.Path & "\News.ini")
    Call WriteINI("DATA", "Desc", vbNullString, App.Path & "\News.ini")

    ' Create the key names and values for the RGB color code.
    Call WriteINI("COLOR", "Red", CStr(255), App.Path & "\News.ini")
    Call WriteINI("COLOR", "Green", CStr(255), App.Path & "\News.ini")
    Call WriteINI("COLOR", "Blue", CStr(255), App.Path & "\News.ini")

    ' Create the key names and values for the font data.
    Call WriteINI("FONT", "Font", "Arial", App.Path & "\News.ini")
    Call WriteINI("FONT", "Size", CStr(14), App.Path & "\News.ini")
End Sub

Public Sub FileCreateFontINI()
    ' Create the key names and values for the font data.
    Call WriteINI("FONT", "Font", "FixedSys", App.Path & "\Font.ini")
    Call WriteINI("FONT", "Size", CStr(18), App.Path & "\Font.ini")
End Sub

Public Sub LoadColors()
    Dim R1 As Long
    Dim G1 As Long
    Dim B1 As Long

    ' Check for errors when loading data.
    On Error GoTo ErrorHandle

    ' Read the chatbox back color RGB color code and create it.
    R1 = CInt(ReadINI("CHATBOX", "R", App.Path & "\GUI\Colors.txt"))
    G1 = CInt(ReadINI("CHATBOX", "G", App.Path & "\GUI\Colors.txt"))
    B1 = CInt(ReadINI("CHATBOX", "B", App.Path & "\GUI\Colors.txt"))
    frmMirage.txtChat.BackColor = RGB(R1, G1, B1)

    ' Read the chatbox text color RGB color code and create it.
    R1 = CInt(ReadINI("CHATTEXTBOX", "R", App.Path & "\GUI\Colors.txt"))
    G1 = CInt(ReadINI("CHATTEXTBOX", "G", App.Path & "\GUI\Colors.txt"))
    B1 = CInt(ReadINI("CHATTEXTBOX", "B", App.Path & "\GUI\Colors.txt"))
    frmMirage.txtMyTextBox.BackColor = RGB(R1, G1, B1)

    ' Read the spell list back color RGB color code and create it.
    R1 = CInt(ReadINI("SPELLLIST", "R", App.Path & "\GUI\Colors.txt"))
    G1 = CInt(ReadINI("SPELLLIST", "G", App.Path & "\GUI\Colors.txt"))
    B1 = CInt(ReadINI("SPELLLIST", "B", App.Path & "\GUI\Colors.txt"))
    frmMirage.lstSpells.BackColor = RGB(R1, G1, B1)

    ' Read the who's online back color RGB color code and create it.
    R1 = CInt(ReadINI("WHOLIST", "R", App.Path & "\GUI\Colors.txt"))
    G1 = CInt(ReadINI("WHOLIST", "G", App.Path & "\GUI\Colors.txt"))
    B1 = CInt(ReadINI("WHOLIST", "B", App.Path & "\GUI\Colors.txt"))
    frmMirage.lstOnline.BackColor = RGB(R1, G1, B1)

    ' Read the mini menu back color RGB color code.
    R1 = CInt(ReadINI("BACKGROUND", "R", App.Path & "\GUI\Colors.txt"))
    G1 = CInt(ReadINI("BACKGROUND", "G", App.Path & "\GUI\Colors.txt"))
    B1 = CInt(ReadINI("BACKGROUND", "B", App.Path & "\GUI\Colors.txt"))

    ' Create the mini menu back color.
    frmMirage.picInventory.BackColor = RGB(R1, G1, B1)
    frmMirage.picInventory3.BackColor = RGB(R1, G1, B1)
    frmMirage.itmDesc.BackColor = RGB(R1, G1, B1)
    frmMirage.picWhosOnline.BackColor = RGB(R1, G1, B1)
    frmMirage.picGuildAdmin.BackColor = RGB(R1, G1, B1)
    frmMirage.picGuildMember.BackColor = RGB(R1, G1, B1)
    frmMirage.picEquipment.BackColor = RGB(R1, G1, B1)
    frmMirage.picPlayerSpells.BackColor = RGB(R1, G1, B1)
    frmMirage.picOptions.BackColor = RGB(R1, G1, B1)

    ' Create the option menu back color.
    frmMirage.chkBubbleBar.BackColor = RGB(R1, G1, B1)
    frmMirage.chkNpcBar.BackColor = RGB(R1, G1, B1)
    frmMirage.chkNpcName.BackColor = RGB(R1, G1, B1)
    frmMirage.chkPlayerBar.BackColor = RGB(R1, G1, B1)
    frmMirage.chkPlayerName.BackColor = RGB(R1, G1, B1)
    frmMirage.chkPlayerDamage.BackColor = RGB(R1, G1, B1)
    frmMirage.chkNpcDamage.BackColor = RGB(R1, G1, B1)
    frmMirage.chkMusic.BackColor = RGB(R1, G1, B1)
    frmMirage.chkSound.BackColor = RGB(R1, G1, B1)
    frmMirage.chkAutoScroll.BackColor = RGB(R1, G1, B1)
    frmMirage.chkSwearFilter.BackColor = RGB(R1, G1, B1)

    Exit Sub

ErrorHandle:
    ' We failed to load the color tables.
    Call MsgBox("Failed to load the file: Colors.txt")
End Sub

Public Sub FileCreateColorsTXT()
    ' Create the key names and values for the chatbox data.
    Call WriteINI("CHATBOX", "R", CStr(152), App.Path & "\GUI\Colors.txt")
    Call WriteINI("CHATBOX", "G", CStr(146), App.Path & "\GUI\Colors.txt")
    Call WriteINI("CHATBOX", "B", CStr(120), App.Path & "\GUI\Colors.txt")

    ' Create the key names and values for the chatbox text data.
    Call WriteINI("CHATTEXTBOX", "R", CStr(152), App.Path & "\GUI\Colors.txt")
    Call WriteINI("CHATTEXTBOX", "G", CStr(146), App.Path & "\GUI\Colors.txt")
    Call WriteINI("CHATTEXTBOX", "B", CStr(120), App.Path & "\GUI\Colors.txt")

    ' Create the key names and values for the background data.
    Call WriteINI("BACKGROUND", "R", CStr(152), App.Path & "\GUI\Colors.txt")
    Call WriteINI("BACKGROUND", "G", CStr(146), App.Path & "\GUI\Colors.txt")
    Call WriteINI("BACKGROUND", "B", CStr(120), App.Path & "\GUI\Colors.txt")

    ' Create the key names and values for the spell list data.
    Call WriteINI("SPELLLIST", "R", CStr(152), App.Path & "\GUI\Colors.txt")
    Call WriteINI("SPELLLIST", "G", CStr(146), App.Path & "\GUI\Colors.txt")
    Call WriteINI("SPELLLIST", "B", CStr(120), App.Path & "\GUI\Colors.txt")

    ' Create the key names and values for the who's online data.
    Call WriteINI("WHOLIST", "R", CStr(152), App.Path & "\GUI\Colors.txt")
    Call WriteINI("WHOLIST", "G", CStr(146), App.Path & "\GUI\Colors.txt")
    Call WriteINI("WHOLIST", "B", CStr(120), App.Path & "\GUI\Colors.txt")

    ' Create the key names and values for the new character data.
    Call WriteINI("NEWCHAR", "R", CStr(152), App.Path & "\GUI\Colors.txt")
    Call WriteINI("NEWCHAR", "G", CStr(146), App.Path & "\GUI\Colors.txt")
    Call WriteINI("NEWCHAR", "B", CStr(120), App.Path & "\GUI\Colors.txt")
End Sub

Public Sub LoadFont()
    On Error GoTo ErrorHandle

    ' Get the font name.
    Font = ReadINI("FONT", "Font", App.Path & "\Font.ini")

    ' Get the font size.
    FontSize = CByte(ReadINI("FONT", "Size", App.Path & "\Font.ini"))

    ' Check if the font name exists.
    If LenB(Font) = 0 Then Font = "FixedSys"

    ' Check if the font size is valid.
    If FontSize < 1 Or FontSize > 32 Then FontSize = 18

    ' Create the font.
    Call SetFont(Font, FontSize)

    Exit Sub

ErrorHandle:
    ' Failed to load the font data, so use the default font.
    Call MsgBox("Failed to the load the font. Defaulting to FixedSys:18.", vbOKOnly)
    Call SetFont("FixedSys", 18)
End Sub
