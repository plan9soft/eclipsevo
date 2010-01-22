Attribute VB_Name = "modGeneral"
Option Explicit

Public Sub InitServer()
    Call SetStatus("Checking Folders...")
    Call FolderPath_Init
    Call FolderCheck

    Call SetStatus("Checking Files...")
    Call FilePath_Init
    Call FileCheck

    Call SetStatus("Loading Settings...")
    Call Config_FileLoad
    Call Status_FileLoad
    Call Engine_DefineGlobals
    Call Engine_DefineArrays

    Call SetStatus("Loading Scripts...")
    Call Script_LoadEngine

    Call SetStatus("Loading Sockets...")
    Call Engine_CreateSockets

    Call SetStatus("Clearing Arrays...")
    Call ClearArrows
    Call ClearTempTile
    Call ClearMaps
    Call ClearMapItems
    Call ClearMapNpcs
    Call ClearNpcs
    Call ClearItems
    Call ClearShops
    Call ClearSpells
    Call ClearExperience
    Call ClearEmoticon

    Call SetStatus("Loading Emoticons...")
    Call LoadEmoticon
    Call SetStatus("Loading Elements...")
    Call LoadElements
    Call SetStatus("Loading Arrows...")
    Call LoadArrows
    Call SetStatus("Loading Experience...")
    Call LoadExperience
    Call SetStatus("Loading Classes...")
    Call LoadClasses
    Call SetStatus("Loading Maps...")
    Call LoadMaps
    Call SetStatus("Loading Items...")
    Call LoadItems
    Call SetStatus("Loading NPCs...")
    Call LoadNpcs
    Call SetStatus("Loading Shops...")
    Call LoadShops
    Call SetStatus("Loading Spells...")
    Call LoadSpells
    Call SetStatus("Loading Guilds...")
    Call LoadGuilds

    Call SetStatus("Spawning Map Items...")
    Call SpawnAllMapsItems
    Call SetStatus("Spawning Map NPCs...")
    Call SpawnAllMapNpcs

    Call SetStatus("Preparing GUI...")
    Call Engine_UpdateGUI

    Call SetStatus("Preparing Listen Socket...")
    Call Engine_StartListening
    
    If SCRIPTING = 1 Then
        MyScript.ExecuteStatement "Scripts\Main.txt", "OnServerLoad"
    End If

    Call UpdateTitle
    Call UpdateTOP

    frmLoad.Visible = False
    frmServer.Show

    frmServer.tmrGameAI.Enabled = True
    frmServer.tmrScriptedTimer.Enabled = True
    frmServer.tmrPlayerSave.Enabled = True
    frmServer.tmrSpawnMapItems.Enabled = True
    frmServer.tmrDayNight.Enabled = True

    SERV_ISRUNNING = True
End Sub

Public Sub FolderCheck()
    ' Check if the 'Maps' folder exists.
    If Not FolderExists(FLDR_MAPS) Then
        Call MkDir(FLDR_MAPS)
    End If

    ' Check if the 'Logs' folder exists.
    If Not FolderExists(FLDR_LOGS) Then
        Call MkDir(FLDR_LOGS)
    End If

    ' Check if the 'Accounts' folder exists.
    If Not FolderExists(FLDR_ACCOUNTS) Then
        Call MkDir(FLDR_ACCOUNTS)
    End If

    ' Check if the 'NPCs' folder exists.
    If Not FolderExists(FLDR_NPCS) Then
        Call MkDir(FLDR_NPCS)
    End If

    ' Check if the 'Items' folder exists.
    If Not FolderExists(FLDR_ITEMS) Then
        Call MkDir(FLDR_ITEMS)
    End If

    ' Check if the 'Spells' folder exists.
    If Not FolderExists(FLDR_SPELLS) Then
        Call MkDir(FLDR_SPELLS)
    End If

    ' Check if the 'Shops' folder exists.
    If Not FolderExists(FLDR_SHOPS) Then
        Call MkDir(FLDR_SHOPS)
    End If

    ' Check if the 'Banks' folder exists.
    If Not FolderExists(FLDR_BANKS) Then
        Call MkDir(FLDR_BANKS)
    End If

    ' Check if the 'Classes' folder exists.
    If Not FolderExists(FLDR_CLASSES) Then
        Call MkDir(FLDR_CLASSES)
    End If
End Sub

Public Sub FileCheck()
    Dim FileID As Integer

    ' Check if the 'Data.ini' file exists.
    If Not FileExists("Data.ini") Then
        Call PutVar(FILE_DATAINI, "CONFIG", "GameName", "Eclipse")
        Call PutVar(FILE_DATAINI, "CONFIG", "WebSite", vbNullString)
        Call PutVar(FILE_DATAINI, "CONFIG", "Port", CStr(4001))
        Call PutVar(FILE_DATAINI, "CONFIG", "HPRegen", CStr(1))
        Call PutVar(FILE_DATAINI, "CONFIG", "HPTimer", CStr(1000))
        Call PutVar(FILE_DATAINI, "CONFIG", "MPRegen", CStr(1))
        Call PutVar(FILE_DATAINI, "CONFIG", "MPTimer", CStr(1000))
        Call PutVar(FILE_DATAINI, "CONFIG", "SPRegen", CStr(1))
        Call PutVar(FILE_DATAINI, "CONFIG", "SPTimer", CStr(1000))
        Call PutVar(FILE_DATAINI, "CONFIG", "NPCRegen", CStr(1))
        Call PutVar(FILE_DATAINI, "CONFIG", "Stat1", "Strength")
        Call PutVar(FILE_DATAINI, "CONFIG", "Stat2", "Defense")
        Call PutVar(FILE_DATAINI, "CONFIG", "Stat3", "Speed")
        Call PutVar(FILE_DATAINI, "CONFIG", "Stat4", "Magic")
        Call PutVar(FILE_DATAINI, "CONFIG", "PlayerCard", CStr(0))
        Call PutVar(FILE_DATAINI, "CONFIG", "Scrolling", CStr(0))
        Call PutVar(FILE_DATAINI, "CONFIG", "ScrollX", CStr(30))
        Call PutVar(FILE_DATAINI, "CONFIG", "ScrollY", CStr(30))
        Call PutVar(FILE_DATAINI, "CONFIG", "Scripting", CStr(0))
        Call PutVar(FILE_DATAINI, "CONFIG", "ScriptErrors", CStr(0))
        Call PutVar(FILE_DATAINI, "CONFIG", "PaperDoll", CStr(0))
        Call PutVar(FILE_DATAINI, "CONFIG", "SaveTime", CStr(0))
        Call PutVar(FILE_DATAINI, "CONFIG", "SpriteSize", CStr(0))
        Call PutVar(FILE_DATAINI, "CONFIG", "CustomPlayerGFX", CStr(0))
        Call PutVar(FILE_DATAINI, "CONFIG", "Level", CStr(0))
        Call PutVar(FILE_DATAINI, "CONFIG", "PKMinLvl", CStr(10))
        Call PutVar(FILE_DATAINI, "CONFIG", "Email", CStr(0))
        Call PutVar(FILE_DATAINI, "CONFIG", "Classes", CStr(1))
        Call PutVar(FILE_DATAINI, "CONFIG", "SPAttack", CStr(0))
        Call PutVar(FILE_DATAINI, "CONFIG", "SPRunning", CStr(0))
        Call PutVar(FILE_DATAINI, "MAX", "MAX_PLAYERS", CStr(50))
        Call PutVar(FILE_DATAINI, "MAX", "MAX_ITEMS", CStr(100))
        Call PutVar(FILE_DATAINI, "MAX", "MAX_NPCS", CStr(100))
        Call PutVar(FILE_DATAINI, "MAX", "MAX_SHOPS", CStr(100))
        Call PutVar(FILE_DATAINI, "MAX", "MAX_SPELLS", CStr(100))
        Call PutVar(FILE_DATAINI, "MAX", "MAX_MAPS", CStr(50))
        Call PutVar(FILE_DATAINI, "MAX", "MAX_MAP_ITEMS", CStr(20))
        Call PutVar(FILE_DATAINI, "MAX", "MAX_GUILDS", CStr(20))
        Call PutVar(FILE_DATAINI, "MAX", "MAX_GUILD_MEMBERS", CStr(10))
        Call PutVar(FILE_DATAINI, "MAX", "MAX_EMOTICONS", CStr(10))
        Call PutVar(FILE_DATAINI, "MAX", "MAX_LEVEL", CStr(500))
        Call PutVar(FILE_DATAINI, "MAX", "MAX_PARTY_MEMBERS", CStr(4))
        Call PutVar(FILE_DATAINI, "MAX", "MAX_ELEMENTS", CStr(20))
        Call PutVar(FILE_DATAINI, "MAX", "MAX_SCRIPTSPELLS", CStr(500))
        Call PutVar(FILE_DATAINI, "MAX", "MAX_PACKETS", CStr(25))
        Call PutVar(FILE_DATAINI, "MAX", "MAX_BYTES", CStr(1000))
        Call PutVar(FILE_DATAINI, "ENGINE", "EngineSpeed", CStr(500))
    End If

    ' Check if the 'Stats.ini' file exists.
    If Not FileExists("Stats.ini") Then
        Call PutVar(FILE_STATSINI, "HP", "AddPerLevel", CStr(10))
        Call PutVar(FILE_STATSINI, "HP", "AddPerStrength", CStr(10))
        Call PutVar(FILE_STATSINI, "HP", "AddPerDefense", CStr(0))
        Call PutVar(FILE_STATSINI, "HP", "AddPerMagic", CStr(0))
        Call PutVar(FILE_STATSINI, "HP", "AddPerSpeed", CStr(0))
        Call PutVar(FILE_STATSINI, "MP", "AddPerLevel", CStr(10))
        Call PutVar(FILE_STATSINI, "MP", "AddPerStrength", CStr(0))
        Call PutVar(FILE_STATSINI, "MP", "AddPerDefense", CStr(0))
        Call PutVar(FILE_STATSINI, "MP", "AddPerMagic", CStr(10))
        Call PutVar(FILE_STATSINI, "MP", "AddPerSpeed", CStr(0))
        Call PutVar(FILE_STATSINI, "SP", "AddPerLevel", CStr(10))
        Call PutVar(FILE_STATSINI, "SP", "AddPerStrength", CStr(0))
        Call PutVar(FILE_STATSINI, "SP", "AddPerDefense", CStr(0))
        Call PutVar(FILE_STATSINI, "SP", "AddPerMagic", CStr(0))
        Call PutVar(FILE_STATSINI, "SP", "AddPerSpeed", CStr(20))
    End If

    ' Check if the 'News.ini' file exists.
    If Not FileExists("News.ini") Then
        Call PutVar(FILE_NEWSINI, "DATA", "NewsTitle", "Change this message in News.ini.")
        Call PutVar(FILE_NEWSINI, "DATA", "NewsBody", "Change this message in News.ini.")
        Call PutVar(FILE_NEWSINI, "COLOR", "Red", CStr(255))
        Call PutVar(FILE_NEWSINI, "COLOR", "Green", CStr(255))
        Call PutVar(FILE_NEWSINI, "COLOR", "Blue", CStr(255))
    End If

    ' Check if the 'MOTD.ini' file exists.
    If Not FileExists("MOTD.ini") Then
        Call PutVar(FILE_MOTDINI, "MOTD", "Msg", "Change this message in MOTD.ini.")
    End If

    ' Check if the 'Tiles.ini' file exists.
    If Not FileExists("Tiles.ini") Then
        For FileID = 0 To 100
            Call PutVar(FILE_TILESINI, "Names", "Tile" & FileID, CStr(FileID))
        Next FileID
    End If

    ' Check if the 'Accounts\CharList.ini' file exists.
    If Not FileExists("Accounts\CharList.txt") Then
        FileID = FreeFile
        Open App.Path & "\Accounts\CharList.txt" For Output As #FileID
        Close #FileID
    End If
End Sub

Public Sub Status_FileLoad()
    On Error GoTo Status_FileLoad_Error

    ' Load the HP configuration settings.
    AddHP.LEVEL = CLng(GetVar(FILE_STATSINI, "HP", "AddPerLevel"))
    AddHP.STR = CLng(GetVar(FILE_STATSINI, "HP", "AddPerStrength"))
    AddHP.DEF = CLng(GetVar(FILE_STATSINI, "HP", "AddPerDefense"))
    AddHP.Magi = CLng(GetVar(FILE_STATSINI, "HP", "AddPerMagic"))
    AddHP.Speed = CLng(GetVar(FILE_STATSINI, "HP", "AddPerSpeed"))

    ' Load the MP configuration settings.
    AddMP.LEVEL = CLng(GetVar(FILE_STATSINI, "MP", "AddPerLevel"))
    AddMP.STR = CLng(GetVar(FILE_STATSINI, "MP", "AddPerStrength"))
    AddMP.DEF = CLng(GetVar(FILE_STATSINI, "MP", "AddPerDefense"))
    AddMP.Magi = CLng(GetVar(FILE_STATSINI, "MP", "AddPerMagic"))
    AddMP.Speed = CLng(GetVar(FILE_STATSINI, "MP", "AddPerSpeed"))

    ' Load the SP configuration settings.
    AddSP.LEVEL = CLng(GetVar(FILE_STATSINI, "SP", "AddPerLevel"))
    AddSP.STR = CLng(GetVar(FILE_STATSINI, "SP", "AddPerStrength"))
    AddSP.DEF = CLng(GetVar(FILE_STATSINI, "SP", "AddPerDefense"))
    AddSP.Magi = CLng(GetVar(FILE_STATSINI, "SP", "AddPerMagic"))
    AddSP.Speed = CLng(GetVar(FILE_STATSINI, "SP", "AddPerSpeed"))

    Exit Sub

Status_FileLoad_Error:
    Call MsgBox("Failed to load the file: Stats.ini.", vbOKOnly)
    End
End Sub

Public Sub Config_FileLoad()
    On Error GoTo Config_FileLoad_Error

    ' Load the core configuration settings.
    GAME_NAME = GetVar(FILE_DATAINI, "CONFIG", "GameName")
    WEB_SITE = GetVar(FILE_DATAINI, "CONFIG", "WebSite")
    GAME_PORT = CLng(GetVar(FILE_DATAINI, "CONFIG", "Port"))
    MAX_PLAYERS = CInt(GetVar(FILE_DATAINI, "MAX", "MAX_PLAYERS"))
    MAX_ITEMS = CInt(GetVar(FILE_DATAINI, "MAX", "MAX_ITEMS"))
    MAX_NPCS = CInt(GetVar(FILE_DATAINI, "MAX", "MAX_NPCS"))
    MAX_SHOPS = CInt(GetVar(FILE_DATAINI, "MAX", "MAX_SHOPS"))
    MAX_SPELLS = CInt(GetVar(FILE_DATAINI, "MAX", "MAX_SPELLS"))
    MAX_MAPS = CInt(GetVar(FILE_DATAINI, "MAX", "MAX_MAPS"))
    MAX_MAP_ITEMS = CInt(GetVar(FILE_DATAINI, "MAX", "MAX_MAP_ITEMS"))
    MAX_GUILDS = CInt(GetVar(FILE_DATAINI, "MAX", "MAX_GUILDS"))
    MAX_GUILD_MEMBERS = CInt(GetVar(FILE_DATAINI, "MAX", "MAX_GUILD_MEMBERS"))
    MAX_EMOTICONS = CInt(GetVar(FILE_DATAINI, "MAX", "MAX_EMOTICONS"))
    MAX_LEVEL = CInt(GetVar(FILE_DATAINI, "MAX", "MAX_LEVEL"))
    MAX_PARTY_MEMBERS = CInt(GetVar(FILE_DATAINI, "MAX", "MAX_PARTY_MEMBERS"))
    MAX_ELEMENTS = CInt(GetVar(FILE_DATAINI, "MAX", "MAX_ELEMENTS"))
    MAX_SCRIPTSPELLS = CInt(GetVar(FILE_DATAINI, "MAX", "MAX_SCRIPTSPELLS"))
    SCRIPTING = CByte(GetVar(FILE_DATAINI, "CONFIG", "Scripting"))
    SCRIPT_DEBUG = CByte(GetVar(FILE_DATAINI, "CONFIG", "ScriptErrors"))
    PAPERDOLL = CByte(GetVar(FILE_DATAINI, "CONFIG", "PaperDoll"))
    SPRITESIZE = CByte(GetVar(FILE_DATAINI, "CONFIG", "SpriteSize"))
    HP_REGEN = CByte(GetVar(FILE_DATAINI, "CONFIG", "HPRegen"))
    HP_TIMER = CLng(GetVar(FILE_DATAINI, "CONFIG", "HPTimer"))
    MP_REGEN = CByte(GetVar(FILE_DATAINI, "CONFIG", "MPRegen"))
    MP_TIMER = CLng(GetVar(FILE_DATAINI, "CONFIG", "MPTimer"))
    SP_REGEN = CByte(GetVar(FILE_DATAINI, "CONFIG", "SPRegen"))
    SP_TIMER = CLng(GetVar(FILE_DATAINI, "CONFIG", "SPTimer"))
    NPC_REGEN = CByte(GetVar(FILE_DATAINI, "CONFIG", "NPCRegen"))
    STAT1 = GetVar(FILE_DATAINI, "CONFIG", "Stat1")
    STAT2 = GetVar(FILE_DATAINI, "CONFIG", "Stat2")
    STAT3 = GetVar(FILE_DATAINI, "CONFIG", "Stat3")
    STAT4 = GetVar(FILE_DATAINI, "CONFIG", "Stat4")
    SP_ATTACK = CByte(GetVar(FILE_DATAINI, "CONFIG", "SPAttack"))
    SP_RUNNING = CByte(GetVar(FILE_DATAINI, "CONFIG", "SPRunning"))
    CUSTOM_SPRITE = CInt(GetVar(FILE_DATAINI, "CONFIG", "CustomPlayerGFX"))
    EMAIL_AUTH = CInt(GetVar(FILE_DATAINI, "CONFIG", "Email"))
    SAVETIME = CLng(GetVar(FILE_DATAINI, "CONFIG", "SaveTime"))
    LEVEL = CInt(GetVar(FILE_DATAINI, "CONFIG", "Level"))
    PKMINLVL = CInt(GetVar(FILE_DATAINI, "CONFIG", "PKMinLvl"))
    CLASSES = CByte(GetVar(FILE_DATAINI, "CONFIG", "Classes"))
    MAX_PACKETS = CLng(GetVar(FILE_DATAINI, "MAX", "MAX_PACKETS"))
    MAX_BYTES = CLng(GetVar(FILE_DATAINI, "MAX", "MAX_BYTES"))

    ' This is only temporary until I get the GetTickCount timers setup.
    frmServer.tmrGameAI.Interval = CLng(GetVar(FILE_DATAINI, "ENGINE", "EngineSpeed"))

    ' Define the start map, X, and Y.
    START_MAP = CLng(GetVar(FILE_DATAINI, "CONFIG", "RespawnMap"))
    START_X = CLng(GetVar(FILE_DATAINI, "CONFIG", "RespawnMapX"))
    START_Y = CLng(GetVar(FILE_DATAINI, "CONFIG", "RespawnMapY"))

    ' Load the scrolling map configuration settings.
    If GetVar(FILE_DATAINI, "CONFIG", "Scrolling") = 0 Then
        IS_SCROLLING = 0
        MAX_MAPX = 19
        MAX_MAPY = 14
    Else
        IS_SCROLLING = 1
        MAX_MAPX = CLng(GetVar(FILE_DATAINI, "CONFIG", "ScrollX"))
        MAX_MAPY = CLng(GetVar(FILE_DATAINI, "CONFIG", "ScrollY"))
    End If

    Exit Sub

Config_FileLoad_Error:
    Call MsgBox("Failed to load the file: Data.ini.", vbOKOnly)
    End
End Sub

Public Sub Engine_DefineGlobals()
    ' Weather variables.
    WeatherType = WEATHER_NONE
    WeatherLevel = 25

    ' Log the server.
    ServerLog = True
End Sub

Public Sub Engine_DefineArrays()
    Dim I As Long

    ' Re-define all of the map arrays.
    ReDim Map(1 To MAX_MAPS) As MapRec
    ReDim MapCache(1 To MAX_MAPS) As String
    ReDim MapItem(1 To MAX_MAPS, 1 To MAX_MAP_ITEMS) As MapItemRec
    ReDim MapNPC(1 To MAX_MAPS, 1 To MAX_MAP_NPCS) As MapNpcRec

    ReDim TempTile(1 To MAX_MAPS) As TempTileRec

    For I = 1 To MAX_MAPS
        ReDim Map(I).Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
        ReDim TempTile(I).DoorOpen(0 To MAX_MAPX, 0 To MAX_MAPY) As Byte
    Next I

    ' Re-define all of the player arrays.
    ReDim Player(1 To MAX_PLAYERS) As AccountRec
    ReDim PlayersOnMap(1 To MAX_MAPS) As Long

    ' Re-define all of the core arrays.
    ReDim Item(0 To MAX_ITEMS) As ItemRec
    ReDim NPC(0 To MAX_NPCS) As NpcRec
    ReDim Shop(1 To MAX_SHOPS) As ShopRec
    ReDim Spell(1 To MAX_SPELLS) As SpellRec
    ReDim Emoticons(0 To MAX_EMOTICONS) As EmoRec
    ReDim Element(0 To MAX_ELEMENTS) As ElementRec
    ReDim Party(1 To 1) As NewPartyRec ' We use a dynamic party array. [Mellowz]

    ' Re-define all of the guild arrays.
    ReDim Guild(1 To MAX_GUILDS) As GuildRec
    'For I = 1 To MAX_GUILDS
    '    ReDim Guild(I).Member(1 To MAX_GUILD_MEMBERS) As String * NAME_LENGTH
    'Next I

    ' Re-define the experience array.
    ReDim Experience(1 To MAX_LEVEL) As Long
End Sub

Public Sub Script_LoadEngine()
    On Error GoTo Script_LoadEngine_Error

    ' Check for the 'Main.txt' file.
    If SCRIPTING = 1 Then
        If Not FileExists("\Scripts\Main.txt") Then
            Call MsgBox("Main.txt not found. Scripts disabled.", vbExclamation)
            SCRIPTING = 0
        End If
    End If

    Set CTimers = New Collection

    ' Check if scripting is still enabled.
    If SCRIPTING = 1 Then
        ' Create the scripting objects.
        Set MyScript = New clsSadScript
        Set clsScriptCommands = New clsCommands

        ' We don't allow UI's in our scripts since the server is single-threaded.
        ' For example, sub-routines like MsgBox() will not be allowed in your scripts.
        MyScript.SControl.AllowUI = False

        ' The amount of time being our program forcefully stops a script from executing (in MS).
        MyScript.SControl.Timeout = 5000

        ' Load the scripts into memory.
        MyScript.ReadInCode App.Path & "\Scripts\Main.txt", "Scripts\Main.txt", MyScript.SControl
        MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True

        ' Update the GUI with the changes.
        frmServer.lblScriptOn.Caption = "Scripts: ON"
    End If

    Exit Sub

Script_LoadEngine_Error:
    Call MsgBox("Failed to load the scripting engine.", vbOKOnly)
    End
End Sub

Public Sub Engine_CreateSockets()
    Dim I As Long

    On Error GoTo Engine_CreateSockets_Error

    ' Define the seperate and end characters.
    SEP_CHAR = Chr$(0)
    END_CHAR = Chr$(237)

    ' Create the listen object.
    Set GameServer = New clsServer

    ' Load our byte headers.
    Call TCPAssignHeaders

    ' Initialize all the player sockets.
    For I = 1 To MAX_PLAYERS
        Call ClearPlayer(I)
        Call GameServer.Sockets.Add(CStr(I))
    Next I

    ' Update the GUI with the changes.
    For I = 1 To MAX_PLAYERS
        Call ShowPLR(I)
    Next I

    Exit Sub

Engine_CreateSockets_Error:
    Call MsgBox("Failed to load the networking engine.", vbOKOnly)
    End
End Sub

Public Sub Engine_UpdateGUI()
    Dim I As Long

    frmServer.MapList.Clear

    For I = 1 To MAX_MAPS
        frmServer.MapList.AddItem I & ": " & Map(I).Name
    Next I

    frmServer.MapList.Selected(0) = True
End Sub

Public Sub Engine_StartListening()
    ' Start listening.
    GameServer.StartListening

    ' The address is already in use.
    If Err.Number = 10048 Then
        Call MsgBox("The port on this address is already busy.", vbOKOnly)
        End
    End If
End Sub

Public Sub Engine_StopListening()
    ' Stop listening.
    GameServer.StopListening
End Sub

Public Sub DestroyServer()
    Dim I As Long
    Dim Temp As Long

    ' Switch the server off.
    SERV_ISRUNNING = False

    ' Save all the players currently connected.
    Call SaveAllPlayersOnline

    ' Disable all the server-side timers.
    Call SetStatus("Disabling Server-side Timers...")
    frmServer.tmrGameAI.Enabled = False
    frmServer.tmrScriptedTimer.Enabled = False
    frmServer.tmrPlayerSave.Enabled = False
    frmServer.tmrSpawnMapItems.Enabled = False
    frmServer.tmrDayNight.Enabled = False

    ' Hide the server and display the status form
    frmServer.Visible = False
    frmLoad.Visible = True

    ' Stop listening for connections and data.
    Call Engine_StopListening

    ' Unload all sockets created by the server.
    For I = 1 To MAX_PLAYERS
        Temp = I / MAX_PLAYERS * 100
        Call SetStatus("Unloading Sockets... " & Temp & "%")
        Call GameServer.Sockets.Remove(CStr(I))
    Next I

    ' Unload the game server object.
    Set GameServer = Nothing

    ' Close the server.
    End
End Sub

Sub SetStatus(ByVal Status As String)
    frmLoad.lblStatus.Caption = Status
    DoEvents
End Sub

Sub ServerLogic()
    Call CheckGiveVitals
    Call GameAI
    Call ScriptedTimer
End Sub

Sub CheckSpawnMapItems()
    Dim X As Long
    Dim Y As Long

    ' Used for map item respawning
    SpawnSeconds = SpawnSeconds + 1

    ' Respawns the map items.
    If SpawnSeconds >= 120 Then
        ' 2 minutes have passed
        For Y = 1 To MAX_MAPS
            ' Make sure no one is on the map when it respawns
            If PlayersOnMap(Y) = NO Then
                ' Clear out unnecessary junk
                For X = 1 To MAX_MAP_ITEMS
                    Call ClearMapItem(X, Y)
                Next X

                ' Spawn the items
                Call SpawnMapItems(Y)
                Call SendMapItemsToAll(Y)
            End If
        Next Y

        SpawnSeconds = 0
    End If
End Sub

Public Sub GameAI()
    Dim I As Long, X As Long, Y As Long, n As Long, x1 As Long, y1 As Long, TickCount As Long
    Dim Damage As Long, DistanceX As Long, DistanceY As Long, NPCnum As Long, Target As Long
    Dim DidWalk As Boolean

    On Error Resume Next

    For Y = 1 To MAX_MAPS
        If PlayersOnMap(Y) = YES Then
            TickCount = GetTickCount
            
            ' ////////////////////////////////////
            ' // This is used for closing doors //
            ' ////////////////////////////////////
            If TickCount > TempTile(Y).DoorTimer + 5000 Then
                For y1 = 0 To MAX_MAPY
                    For x1 = 0 To MAX_MAPX
                        If Map(Y).Tile(x1, y1).Type = TILE_TYPE_KEY Then
                            If TempTile(Y).DoorOpen(x1, y1) = YES Then
                                TempTile(Y).DoorOpen(x1, y1) = NO
                                Call SendDataToMap(Y, POut.MapKey & SEP_CHAR & x1 & SEP_CHAR & y1 & SEP_CHAR & 0 & END_CHAR)
                            End If
                        End If
                        
                        If Map(Y).Tile(x1, y1).Type = TILE_TYPE_DOOR Then
                            If TempTile(Y).DoorOpen(x1, y1) = YES Then
                                TempTile(Y).DoorOpen(x1, y1) = NO
                                Call SendDataToMap(Y, POut.MapKey & SEP_CHAR & x1 & SEP_CHAR & y1 & SEP_CHAR & 0 & END_CHAR)
                            End If
                        End If
                    Next x1
                Next y1
            End If
            
            For X = 1 To MAX_MAP_NPCS
                NPCnum = MapNPC(Y, X).num
                
                ' /////////////////////////////////////////
                ' // This is used for ATTACKING ON SIGHT //
                ' /////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(Y).NPC(X) > 0 Then
                    If MapNPC(Y, X).num > 0 Then
                        ' If the npc is a attack on sight, search for a player on the map
                        If NPC(NPCnum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or NPC(NPCnum).Behavior = NPC_BEHAVIOR_GUARD Then
                            For I = 1 To MAX_PLAYERS
                                If IsPlaying(I) Then
                                    If GetPlayerMap(I) = Y Then
                                        If MapNPC(Y, X).Target = 0 Then
                                            If GetPlayerAccess(I) <= ADMIN_MONITER Then
                                                n = NPC(NPCnum).Range
                                                
                                                DistanceX = MapNPC(Y, X).X - GetPlayerX(I)
                                                DistanceY = MapNPC(Y, X).Y - GetPlayerY(I)
                                                
                                                ' Make sure we get a positive value
                                                If DistanceX < 0 Then DistanceX = DistanceX * -1
                                                If DistanceY < 0 Then DistanceY = DistanceY * -1
                                                
                                                ' Are they in range?  if so GET'M!
                                                If DistanceX <= n Then
                                                    If DistanceY <= n Then
                                                        If NPC(NPCnum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or GetPlayerPK(I) = YES Then
                                                            If Trim(NPC(NPCnum).AttackSay) <> vbNullString Then
                                                                Call PlayerMsg(I, "A " & Trim(NPC(NPCnum).Name) & " : " & Trim(NPC(NPCnum).AttackSay), SayColor)
                                                            End If
                                                            
                                                            MapNPC(Y, X).Target = I
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next I
                        End If
                    End If
                End If
                                                                        
                ' /////////////////////////////////////////////
                ' // This is used for NPC walking/targetting //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(Y).NPC(X) > 0 Then
                    If MapNPC(Y, X).num > 0 Then
                        Target = MapNPC(Y, X).Target
                        
                        ' Check to see if its time for the npc to walk
                        If NPC(NPCnum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                            ' Check to see if we are following a player or not
                            If Target > 0 Then
                                ' Check if the player is even playing, if so follow'm
                                If IsPlaying(Target) Then
                                    If GetPlayerMap(Target) = Y Then
                                        DidWalk = False
                                        
                                        I = Int(Rnd * 4)
                                        
                                        ' Lets move the npc
                                        Select Case I
                                            Case 0
                                                ' Up
                                                If MapNPC(Y, X).Y > GetPlayerY(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_UP) Then
                                                            Call NpcMove(Y, X, DIR_UP, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                ' Down
                                                If MapNPC(Y, X).Y < GetPlayerY(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_DOWN) Then
                                                            Call NpcMove(Y, X, DIR_DOWN, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                ' Left
                                                If MapNPC(Y, X).X > GetPlayerX(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_LEFT) Then
                                                            Call NpcMove(Y, X, DIR_LEFT, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                ' Right
                                                If MapNPC(Y, X).X < GetPlayerX(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_RIGHT) Then
                                                            Call NpcMove(Y, X, DIR_RIGHT, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                            
                                            Case 1
                                                ' Right
                                                If MapNPC(Y, X).X < GetPlayerX(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_RIGHT) Then
                                                            Call NpcMove(Y, X, DIR_RIGHT, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                ' Left
                                                If MapNPC(Y, X).X > GetPlayerX(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_LEFT) Then
                                                            Call NpcMove(Y, X, DIR_LEFT, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                ' Down
                                                If MapNPC(Y, X).Y < GetPlayerY(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_DOWN) Then
                                                            Call NpcMove(Y, X, DIR_DOWN, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                ' Up
                                                If MapNPC(Y, X).Y > GetPlayerY(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_UP) Then
                                                            Call NpcMove(Y, X, DIR_UP, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                
                                            Case 2
                                                ' Down
                                                If MapNPC(Y, X).Y < GetPlayerY(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_DOWN) Then
                                                            Call NpcMove(Y, X, DIR_DOWN, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                ' Up
                                                If MapNPC(Y, X).Y > GetPlayerY(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_UP) Then
                                                            Call NpcMove(Y, X, DIR_UP, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                ' Right
                                                If MapNPC(Y, X).X < GetPlayerX(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_RIGHT) Then
                                                            Call NpcMove(Y, X, DIR_RIGHT, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                ' Left
                                                If MapNPC(Y, X).X > GetPlayerX(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_LEFT) Then
                                                            Call NpcMove(Y, X, DIR_LEFT, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                            
                                            Case 3
                                                ' Left
                                                If MapNPC(Y, X).X > GetPlayerX(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_LEFT) Then
                                                            Call NpcMove(Y, X, DIR_LEFT, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                ' Right
                                                If MapNPC(Y, X).X < GetPlayerX(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_RIGHT) Then
                                                            Call NpcMove(Y, X, DIR_RIGHT, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                ' Up
                                                If MapNPC(Y, X).Y > GetPlayerY(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_UP) Then
                                                            Call NpcMove(Y, X, DIR_UP, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                                ' Down
                                                If MapNPC(Y, X).Y < GetPlayerY(Target) Then
                                                    If DidWalk = False Then
                                                        If CanNpcMove(Y, X, DIR_DOWN) Then
                                                            Call NpcMove(Y, X, DIR_DOWN, MOVING_WALKING)
                                                            DidWalk = True
                                                        End If
                                                    End If
                                                End If
                                        End Select
                                    
                                        ' Check if we can't move and if player is behind something and if we can just switch dirs
                                        If Not DidWalk Then
                                            If MapNPC(Y, X).X - 1 = GetPlayerX(Target) Then
                                                If MapNPC(Y, X).Y = GetPlayerY(Target) Then
                                                    If MapNPC(Y, X).Dir <> DIR_LEFT Then
                                                        Call NpcDir(Y, X, DIR_LEFT)
                                                    End If
                                                End If
                                                DidWalk = True
                                            End If
                                            If MapNPC(Y, X).X + 1 = GetPlayerX(Target) Then
                                                If MapNPC(Y, X).Y = GetPlayerY(Target) Then
                                                    If MapNPC(Y, X).Dir <> DIR_RIGHT Then
                                                        Call NpcDir(Y, X, DIR_RIGHT)
                                                    End If
                                                End If
                                                DidWalk = True
                                            End If
                                            If MapNPC(Y, X).X = GetPlayerX(Target) Then
                                                If MapNPC(Y, X).Y - 1 = GetPlayerY(Target) Then
                                                    If MapNPC(Y, X).Dir <> DIR_UP Then
                                                        Call NpcDir(Y, X, DIR_UP)
                                                    End If
                                                End If
                                                DidWalk = True
                                            End If
                                            If MapNPC(Y, X).X = GetPlayerX(Target) Then
                                                If MapNPC(Y, X).Y + 1 = GetPlayerY(Target) Then
                                                    If MapNPC(Y, X).Dir <> DIR_DOWN Then
                                                        Call NpcDir(Y, X, DIR_DOWN)
                                                    End If
                                                End If
                                                DidWalk = True
                                            End If
                                            
                                            ' We could not move so player must be behind something, walk randomly.
                                            If Not DidWalk Then
                                                I = Int(Rnd * 2)
                                                If I = 1 Then
                                                    I = Int(Rnd * 4)
                                                    If CanNpcMove(Y, X, I) Then
                                                        Call NpcMove(Y, X, I, MOVING_WALKING)
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        MapNPC(Y, X).Target = 0
                                    End If
                                End If
                                
                            Else
                                I = Int(Rnd * 4)
                                If I = 1 Then
                                    I = Int(Rnd * 4)
                                    If CanNpcMove(Y, X, I) Then
                                        Call NpcMove(Y, X, I, MOVING_WALKING)
                                    End If
                                End If
                            End If
                            
                        End If
                    End If
                    
                End If
                
                ' /////////////////////////////////////////////
                ' // This is used for npcs to attack players //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(Y).NPC(X) > 0 Then
                    If MapNPC(Y, X).num > 0 Then
                        Target = MapNPC(Y, X).Target
                        
                        ' Check if the npc can attack the targeted player player
                        If Target > 0 Then
                            ' Is the target playing and on the same map?
                            If IsPlaying(Target) And GetPlayerMap(Target) = Y Then
                                ' Can the npc attack the player?
                                If CanNpcAttackPlayer(X, Target) Then
                                    If Not CanPlayerBlockHit(Target) Then
                                    
                                        Damage = NPC(NPCnum).STR - GetPlayerProtection(Target)
                                        
                                        If Damage > 0 Then
                                            If SCRIPTING = 1 Then
                                                MyScript.ExecuteStatement "Scripts\Main.txt", "PlayerHit " & Target & "," & X & "," & Damage
                                            Else
                                                Call NpcAttackPlayer(X, Target, Damage)
                                            End If
                                        Else
                                            Call BattleMsg(Target, "The " & Trim(NPC(NPCnum).Name) & " couldn't hurt you!", BRIGHTBLUE, 1)
                                            
                                            'Call PlayerMsg(Target, "The " & Trim(Npc(NpcNum).Name) & "'s hit didn't even phase you!", BrightBlue)
                                        End If
                                    Else
                                        Call BattleMsg(Target, "You blocked the " & Trim(NPC(NPCnum).Name) & "'s hit!", BRIGHTCYAN, 1)
                                        
                                        'Call PlayerMsg(Target, "Your " & Trim(Item(GetPlayerInvItemNum(Target, GetPlayerShieldSlot(Target))).Name) & " blocks the " & Trim(Npc(NpcNum).Name) & "'s hit!", BrightCyan)
                                    End If
                                End If
                            Else
                                ' Player left map or game, set target to 0
                                MapNPC(Y, X).Target = 0
                            End If
                        End If

                    End If
                End If
                
                ' ////////////////////////////////////////////
                ' // This is used for regenerating NPC's HP //
                ' ////////////////////////////////////////////
                ' Check to see if we want to regen some of the npc's hp
                If MapNPC(Y, X).num > 0 Then
                    If TickCount > GiveNPCHPTimer + 10000 Then
                        If MapNPC(Y, X).HP > 0 Then
                            MapNPC(Y, X).HP = MapNPC(Y, X).HP + GetNpcHPRegen(NPCnum)
                        
                            ' Check if they have more then they should and if so just set it to max
                            If MapNPC(Y, X).HP > GetNpcMaxHP(NPCnum) Then
                                MapNPC(Y, X).HP = GetNpcMaxHP(NPCnum)
                            End If
                        End If
                    End If
                End If

                ' //////////////////////////////////////
                ' // This is used for spawning an NPC //
                ' //////////////////////////////////////
                ' Check if we are supposed to spawn an npc or not
                If MapNPC(Y, X).num = 0 Then
                    If Map(Y).NPC(X) > 0 Then
                        If TickCount > MapNPC(Y, X).SpawnWait + (NPC(Map(Y).NPC(X)).SpawnSecs * 1000) Then
                            Call SpawnNPC(X, Y)
                        End If
                    End If
                End If
            Next X
        End If
    Next Y
    
    ' Make sure we reset the timer for npc hp regeneration
    If GetTickCount > GiveNPCHPTimer + 10000 Then
        GiveNPCHPTimer = GetTickCount
    End If

    ' Make sure we reset the timer for door closing
    If GetTickCount > KeyTimer + 15000 Then
        KeyTimer = GetTickCount
    End If

End Sub


Sub ScriptedTimer()
    Dim X As Long, n As Long
    Dim CustomTimer As clsCTimers

    n = 0
    X = CTimers.Count
    For Each CustomTimer In CTimers
        n = n + 1
        If GetTickCount > CustomTimer.tmrWait Then
            MyScript.ExecuteStatement "Scripts\Main.txt", CustomTimer.Name ' & " " & Index & "," & PointType
            If CTimers.Count < X Then
                n = n - X - CTimers.Count
                X = CTimers.Count
            End If
            If n > 0 Then
                CTimers.Item(n).tmrWait = GetTickCount + CustomTimer.Interval
            Else
                Exit For
            End If
        End If
    Next CustomTimer
End Sub

Sub CheckGiveVitals()
    Dim I As Long

    If HP_REGEN = 1 Then
        If GetTickCount >= GiveHPTimer + HP_TIMER Then
            For I = 1 To MAX_PLAYERS
                If IsPlaying(I) Then
                    If GetPlayerHP(I) < GetPlayerMaxHP(I) Then
                        Call SetPlayerHP(I, GetPlayerHP(I) + GetPlayerHPRegen(I))
                    End If
                End If
            Next I

            GiveHPTimer = GetTickCount
        End If
    End If

    If MP_REGEN = 1 Then
        If GetTickCount >= GiveMPTimer + MP_TIMER Then
            For I = 1 To MAX_PLAYERS
                If IsPlaying(I) Then
                    If GetPlayerMP(I) < GetPlayerMaxMP(I) Then
                        Call SetPlayerMP(I, GetPlayerMP(I) + GetPlayerMPRegen(I))
                    End If
                End If
            Next I

            GiveMPTimer = GetTickCount
        End If
    End If

    If SP_REGEN = 1 Then
        If GetTickCount >= GiveSPTimer + SP_TIMER Then
            For I = 1 To MAX_PLAYERS
                If IsPlaying(I) Then
                    If GetPlayerSP(I) < GetPlayerMaxSP(I) Then
                        Call SetPlayerSP(I, GetPlayerSP(I) + GetPlayerSPRegen(I))
                    End If
                End If
            Next I

            GiveSPTimer = GetTickCount
        End If
    End If
End Sub

Sub PlayerSaveTimer()
    Dim I As Long

    PLYRSAVE_TIMER = PLYRSAVE_TIMER + 1

    If SAVETIME <> 0 Then
        If PLYRSAVE_TIMER >= SAVETIME Then
            For I = 1 To MAX_PLAYERS
                If IsPlaying(I) Then
                    Call SavePlayer(I)
                End If
            Next I
    
            PlayerI = 1

            frmServer.PlayerTimer.Enabled = True
            frmServer.tmrPlayerSave.Enabled = False

            PLYRSAVE_TIMER = 0
        End If
    Else
        PLYRSAVE_TIMER = 0
    End If
End Sub

Function IsAlphaNumeric(TestString As String) As Boolean
    Dim LoopID As Integer
    Dim sChar As String

    IsAlphaNumeric = False

    If LenB(TestString) > 0 Then
        For LoopID = 1 To Len(TestString)
            sChar = Mid(TestString, LoopID, 1)
            If Not sChar Like "[0-9A-Za-z]" Then
                Exit Function
            End If
        Next

        IsAlphaNumeric = True
    End If
End Function

Public Sub FilePath_Init()
    FILE_DATAINI = App.Path & "\Data.ini"
    FILE_STATSINI = App.Path & "\Stats.ini"
    FILE_NEWSINI = App.Path & "\News.ini"
    FILE_MOTDINI = App.Path & "\MOTD.ini"
    FILE_TILESINI = App.Path & "\Tiles.ini"
End Sub

Public Sub FolderPath_Init()
    FLDR_MAPS = App.Path & "\Maps"
    FLDR_LOGS = App.Path & "\Logs"
    FLDR_ACCOUNTS = App.Path & "\Accounts"
    FLDR_NPCS = App.Path & "\NPCs"
    FLDR_ITEMS = App.Path & "\Items"
    FLDR_SPELLS = App.Path & "\Spells"
    FLDR_SHOPS = App.Path & "\Shops"
    FLDR_BANKS = App.Path & "\Banks"
    FLDR_CLASSES = App.Path & "\Classes"
End Sub

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    Randomize
    Rand = Int((High - Low + 1) * Rnd) + Low
End Function

