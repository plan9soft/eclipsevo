Attribute VB_Name = "modHandleData"
Option Explicit

Public Sub HandleData(ByVal Data As String)
    Dim Parse() As String
    Dim Name As String
    Dim Msg As String
    Dim Dir As Long
    Dim Level As Long
    Dim I As Long, n As Long, X As Long, y As Long, p As Long
    Dim shopNum As Long
    Dim z As Long
    Dim strfilename As String
    Dim CustomX As Long
    Dim CustomY As Long
    Dim CustomIndex As Long
    Dim customcolour As Long
    Dim customsize As Long
    Dim customtext As String
    Dim casestring As Byte ' Temp set to long until select case is finished.
    Dim packet As String
    Dim M As Long
    Dim j As Long
    Dim rec As DXVBLib.RECT

    ' Handle Data
    Parse = Split(Data, SEP_CHAR)

    ' Add packet info to debugger
    If frmDebug.Visible Then
        Call TextAdd(frmDebug.txtDebug, Time & " - ( " & Parse(0) & " )", True)
    End If

    casestring = CByte(Parse(0))

    Select Case CByte(Parse(0))
        Case PIn.MaxInfo
            Call Packet_MaxInfo(Parse)
            Exit Sub

        Case PIn.ClearParty
            Call Packet_ClearParty
            Exit Sub

        Case PIn.NPCHP
            Call Packet_UpdateNpcHP(CLng(Parse(1)), CLng(Parse(2)), CLng(Parse(3)))
            Exit Sub
        Case PIn.AlertMessage
            Call Packet_AlertMessage(Parse(1))
            Exit Sub
        Case PIn.PlainMessage
            Call Packet_PlainMessage(Parse(1), CLng(Parse(2)))
            Exit Sub
        Case PIn.CharacterList
            Call Packet_CharacterList(Parse)
            Exit Sub
        Case PIn.LoginOK
            Call Packet_LoginOK(CLng(Parse(1)))
            Exit Sub
        Case PIn.News
            Call Packet_News(Parse(1), Parse(5), CLng(Parse(2)), CLng(Parse(3)), CLng(Parse(4)))
            Exit Sub
        Case PIn.NewCharClasses
            Call Packet_NewCharacterClasses(Parse)
            Exit Sub
        Case PIn.ClassData
            Call Packet_ClassData(Parse)
            Exit Sub
        Case PIn.GameClock
            Call Packet_GameClock(CLng(Parse(1)), CLng(Parse(2)), CLng(Parse(3)), CLng(Parse(4)))
            Exit Sub
        Case PIn.InGame
            Call Packet_InGame
            Exit Sub
        Case PIn.PlayerInventory
            Call Packet_PlayerInventory(Parse)
            Exit Sub
        Case PIn.PlayerInventoryUpdate
            Call Packet_PlayerInventoryUpdate(CLng(Parse(1)), CLng(Parse(2)), CLng(Parse(3)), CLng(Parse(4)), CLng(Parse(5)))
            Exit Sub
        Case PIn.PlayerBank
            Call Packet_PlayerBank(Parse)
            Exit Sub
        Case PIn.PlayerBankUpdate
            Call Packet_PlayerBankUpdate(CLng(Parse(1)), CLng(Parse(2)), CLng(Parse(3)), CLng(Parse(4)))
            Exit Sub
        Case PIn.OpenBank
            Call Packet_OpenBank
            Exit Sub
        Case PIn.BankMessage
            Call Packet_BankMessage(Parse(1))
            Exit Sub
        Case PIn.PlayerWornEQ
            Call Packet_PlayerWornEQ(Parse)
            Exit Sub
        Case PIn.PlayerPoints
            Call Packet_PlayerPoints(CLng(Parse(1)))
            Exit Sub
        Case PIn.CustomSprite
            Call Packet_CustomSprite(CLng(Parse(1)), CLng(Parse(2)), CLng(Parse(3)), CLng(Parse(4)))
            Exit Sub
        Case PIn.PlayerHP
            Call Packet_PlayerHP(CLng(Parse(1)), CLng(Parse(2)))
            Exit Sub
        Case PIn.PlayerMP
            Call Packet_PlayerMP(CLng(Parse(1)), CLng(Parse(2)))
            Exit Sub
        Case PIn.PlayerSP
            Call Packet_PlayerSP(CLng(Parse(1)), CLng(Parse(2)))
            Exit Sub
        Case PIn.PlayerEXP
            Call Packet_PlayerEXP(CLng(Parse(1)), CLng(Parse(2)))
            Exit Sub
        Case PIn.SpeechBubble
            Call Packet_SpeechBubble(Parse(1), CLng(Parse(2)))
            Exit Sub
        Case PIn.ScriptBubble
            Call Packet_ScriptBubble(CLng(Parse(1)), Parse(2), CLng(Parse(3)), CLng(Parse(4)), CLng(Parse(5)), CLng(Parse(6)))
            Exit Sub
        Case PIn.PlayerStats
            Call Packet_PlayerStats(CLng(Parse(1)), CLng(Parse(2)), CLng(Parse(3)), CLng(Parse(4)), CLng(Parse(5)), CLng(Parse(6)), CLng(Parse(7)))
            Exit Sub
    End Select

    ' ::::::::::::::::::::::::
    ' :: Player data packet ::
    ' ::::::::::::::::::::::::
    If casestring = PIn.PlayerData Then
        Dim a As Long
        I = Val(Parse(1))
        Call SetPlayerName(I, Parse(2))
        Call SetPlayerSprite(I, Val(Parse(3)))
        Call SetPlayerMap(I, Val(Parse(4)))
        Call SetPlayerX(I, Val(Parse(5)))
        Call SetPlayerY(I, Val(Parse(6)))
        Call SetPlayerDir(I, Val(Parse(7)))
        Call SetPlayerAccess(I, Val(Parse(8)))
        Call SetPlayerPK(I, Val(Parse(9)))
        Call SetPlayerGuild(I, Parse(10))
        Call SetPlayerGuildAccess(I, Val(Parse(11)))
        Call SetPlayerClass(I, Val(Parse(12)))
        Call SetPlayerHead(I, Val(Parse(13)))
        Call SetPlayerBody(I, Val(Parse(14)))
        Call SetPlayerLeg(I, Val(Parse(15)))
        Call SetPlayerPaperdoll(I, Val(Parse(16)))
        Call SetPlayerLevel(I, Val(Parse(17)))

        ' Make sure they aren't walking
        Player(I).Moving = 0
        Player(I).xOffset = 0
        Player(I).yOffset = 0

        ' Check if the player is the client player, and if so reset directions
        If I = MyIndex Then
            DirUp = False
            DirDown = False
            DirLeft = False
            DirRight = False
        End If
        
        
        Exit Sub
        
    End If
    
    ' if a player leaves the map
    If casestring = PIn.MapLeave Then
        Call SetPlayerMap(CLng(Parse(1)), 0)
        Exit Sub
    End If
        
    ' if a player left the game
    If casestring = PIn.GameLeave Then
        Call ClearPlayer(Parse(1))
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::
    ' :: Player Level Packet  ::
    ' ::::::::::::::::::::::::::
    If casestring = PIn.PlayerLevel Then
        n = Val(Parse(1))
        Player(n).Level = Val(Parse(2))
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::
    ' :: Update Sprite Packet ::
    ' ::::::::::::::::::::::::::
    If casestring = PIn.SpriteUpdate Then
        I = Val(Parse(1))
        Call SetPlayerSprite(I, Val(Parse(1)))
    End If

    ' ::::::::::::::::::::::::::::
    ' :: Player movement packet ::
    ' ::::::::::::::::::::::::::::
    If (casestring = PIn.PlayerMove) Then
        I = Val(Parse(1))
        X = Val(Parse(2))
        y = Val(Parse(3))
        Dir = Val(Parse(4))
        n = Val(Parse(5))

        If Dir < DIR_UP Or Dir > DIR_RIGHT Then
            Exit Sub
        End If

        Call SetPlayerX(I, X)
        Call SetPlayerY(I, y)
        Call SetPlayerDir(I, Dir)

        Player(I).xOffset = 0
        Player(I).yOffset = 0
        Player(I).Moving = n
        
        ' Replaced with the one from TE.
        Select Case GetPlayerDir(I)
            Case DIR_UP
                Player(I).yOffset = PIC_Y
            Case DIR_DOWN
                Player(I).yOffset = PIC_Y * -1
            Case DIR_LEFT
                Player(I).xOffset = PIC_X
            Case DIR_RIGHT
                Player(I).xOffset = PIC_X * -1
        End Select

        Exit Sub
    End If

    ' :::::::::::::::::::::::::
    ' :: Npc movement packet ::
    ' :::::::::::::::::::::::::
    If (casestring = PIn.NpcMove) Then
        I = Val(Parse(1))
        X = Val(Parse(2))
        y = Val(Parse(3))
        Dir = Val(Parse(4))
        n = Val(Parse(5))

        MapNpc(I).X = X
        MapNpc(I).y = y
        MapNpc(I).Dir = Dir
        MapNpc(I).xOffset = 0
        MapNpc(I).yOffset = 0
        MapNpc(I).Moving = 1

        If n <> 1 Then
            Select Case MapNpc(I).Dir
                Case DIR_UP
                    MapNpc(I).yOffset = PIC_Y * Val(n - 1)
                Case DIR_DOWN
                    MapNpc(I).yOffset = PIC_Y * -n
                Case DIR_LEFT
                    MapNpc(I).xOffset = PIC_X * Val(n - 1)
                Case DIR_RIGHT
                    MapNpc(I).xOffset = PIC_X * -n
            End Select
        Else
            Select Case MapNpc(I).Dir
                Case DIR_UP
                    MapNpc(I).yOffset = PIC_Y
                Case DIR_DOWN
                    MapNpc(I).yOffset = PIC_Y * -1
                Case DIR_LEFT
                    MapNpc(I).xOffset = PIC_X
                Case DIR_RIGHT
                    MapNpc(I).xOffset = PIC_X * -1
            End Select
        End If

        Exit Sub
    End If

    ' :::::::::::::::::::::::::::::
    ' :: Player direction packet ::
    ' :::::::::::::::::::::::::::::
    If (casestring = PIn.PlayerDirection) Then
        I = Val(Parse(1))
        Dir = Val(Parse(2))

        If Dir < DIR_UP Or Dir > DIR_RIGHT Then
            Exit Sub
        End If

        Call SetPlayerDir(I, Dir)

        Player(I).xOffset = 0
        Player(I).yOffset = 0
        Player(I).MovingH = 0
        Player(I).MovingV = 0
        Player(I).Moving = 0
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::
    ' :: NPC direction packet ::
    ' ::::::::::::::::::::::::::
    If (casestring = PIn.NPCDirection) Then
        I = Val(Parse(1))
        Dir = Val(Parse(2))
        MapNpc(I).Dir = Dir

        MapNpc(I).xOffset = 0
        MapNpc(I).yOffset = 0
        MapNpc(I).Moving = 0
        Exit Sub
    End If

    ' :::::::::::::::::::::::::::::::
    ' :: Player XY location packet ::
    ' :::::::::::::::::::::::::::::::
    If (casestring = PIn.PlayerXY) Then
        I = Val(Parse(1))
        X = Val(Parse(2))
        y = Val(Parse(3))

        Call SetPlayerX(I, X)
        Call SetPlayerY(I, y)

        ' Make sure they aren't walking
        Player(I).Moving = 0
        Player(I).xOffset = 0
        Player(I).yOffset = 0

        Exit Sub
    End If

    If LCase$(Parse(0)) = PIn.PRemoveMembers Then
        For n = 1 To MAX_PARTY_MEMBERS
            Player(MyIndex).Party.Member(n) = 0
        Next n
        Exit Sub
    End If

    If LCase$(Parse(0)) = PIn.PUpdateMembers Then
        Player(MyIndex).Party.Member(Val(Parse(1))) = Val(Parse(2))
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::
    ' :: Player attack packet ::
    ' ::::::::::::::::::::::::::
    If (casestring = PIn.Attack) Then
        I = Val(Parse(1))

        ' Set player to attacking
        Player(I).Attacking = 1
        Player(I).AttackTimer = GetTickCount

        Exit Sub
    End If

    ' :::::::::::::::::::::::
    ' :: NPC attack packet ::
    ' :::::::::::::::::::::::
    If (casestring = PIn.NPCAttack) Then
        I = Val(Parse(1))

        ' Set player to attacking
        MapNpc(I).Attacking = 1
        MapNpc(I).AttackTimer = GetTickCount
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::
    ' :: Check for map packet ::
    ' ::::::::::::::::::::::::::
    If (casestring = PIn.CheckForMap) Then
        GettingMap = True
    
        ' Erase all players except self
        For I = 1 To MAX_PLAYERS
            If I <> MyIndex Then
                Call SetPlayerMap(I, 0)
            End If
        Next I

        ' Erase all temporary tile values
        Call ClearTempTile

        ' Get map num
        X = Val(Parse(1))

        ' Get revision
        y = Val(Parse(2))

        ' Reset the NPC damage display.
        NPCWho = 0
        DmgDamage = 0
        DmgTime = 0

        ' Reset the player damage display.
        NPCDmgDamage = 0
        NPCDmgTime = 0
        
        ' Close map editor if player leaves current map
        If InEditor Then
            ScreenMode = 0
            NightMode = 0
            GridMode = 0
            InEditor = False
            Unload frmMapEditor
            frmMapEditor.MousePointer = 1
            frmMirage.MousePointer = 1
        End If
        

        If FileExists("maps\map" & X & ".dat") Then
            ' Check to see if the revisions match
            If GetMapRevision(X) = y Then
            ' We do so we dont need the map

                ' Load the map
                Call LoadMap(X)

                Call SendData(POut.NeedMap & SEP_CHAR & CStr(0) & END_CHAR)
                Exit Sub
            End If
        End If

        ' Either the revisions didn't match or we dont have the map, so we need it
        Call SendData(POut.NeedMap & SEP_CHAR & CStr(1) & END_CHAR)
        Exit Sub
    End If

' :::::::::::::::::::::
' :: Map data packet ::
' :::::::::::::::::::::

    If casestring = PIn.MapData Then
        n = 1

        Map(Val(Parse(1))).Name = Parse(n + 1)
        Map(Val(Parse(1))).Revision = Val(Parse(n + 2))
        Map(Val(Parse(1))).Moral = Val(Parse(n + 3))
        Map(Val(Parse(1))).Up = Val(Parse(n + 4))
        Map(Val(Parse(1))).Down = Val(Parse(n + 5))
        Map(Val(Parse(1))).Left = Val(Parse(n + 6))
        Map(Val(Parse(1))).Right = Val(Parse(n + 7))
        Map(Val(Parse(1))).music = Parse(n + 8)
        Map(Val(Parse(1))).BootMap = Val(Parse(n + 9))
        Map(Val(Parse(1))).BootX = Val(Parse(n + 10))
        Map(Val(Parse(1))).BootY = Val(Parse(n + 11))
        Map(Val(Parse(1))).Indoors = Val(Parse(n + 12))
        Map(Val(Parse(1))).Weather = Val(Parse(n + 13))

        n = n + 14

        For y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                Map(Val(Parse(1))).Tile(X, y).Ground = Val(Parse(n))
                Map(Val(Parse(1))).Tile(X, y).Mask = Val(Parse(n + 1))
                Map(Val(Parse(1))).Tile(X, y).Anim = Val(Parse(n + 2))
                Map(Val(Parse(1))).Tile(X, y).Mask2 = Val(Parse(n + 3))
                Map(Val(Parse(1))).Tile(X, y).M2Anim = Val(Parse(n + 4))
                Map(Val(Parse(1))).Tile(X, y).Fringe = Val(Parse(n + 5))
                Map(Val(Parse(1))).Tile(X, y).FAnim = Val(Parse(n + 6))
                Map(Val(Parse(1))).Tile(X, y).Fringe2 = Val(Parse(n + 7))
                Map(Val(Parse(1))).Tile(X, y).F2Anim = Val(Parse(n + 8))
                Map(Val(Parse(1))).Tile(X, y).Type = Val(Parse(n + 9))
                Map(Val(Parse(1))).Tile(X, y).Data1 = Val(Parse(n + 10))
                Map(Val(Parse(1))).Tile(X, y).Data2 = Val(Parse(n + 11))
                Map(Val(Parse(1))).Tile(X, y).Data3 = Val(Parse(n + 12))
                Map(Val(Parse(1))).Tile(X, y).String1 = Parse(n + 13)
                Map(Val(Parse(1))).Tile(X, y).String2 = Parse(n + 14)
                Map(Val(Parse(1))).Tile(X, y).String3 = Parse(n + 15)
                Map(Val(Parse(1))).Tile(X, y).light = Val(Parse(n + 16))
                Map(Val(Parse(1))).Tile(X, y).GroundSet = Val(Parse(n + 17))
                Map(Val(Parse(1))).Tile(X, y).MaskSet = Val(Parse(n + 18))
                Map(Val(Parse(1))).Tile(X, y).AnimSet = Val(Parse(n + 19))
                Map(Val(Parse(1))).Tile(X, y).Mask2Set = Val(Parse(n + 20))
                Map(Val(Parse(1))).Tile(X, y).M2AnimSet = Val(Parse(n + 21))
                Map(Val(Parse(1))).Tile(X, y).FringeSet = Val(Parse(n + 22))
                Map(Val(Parse(1))).Tile(X, y).FAnimSet = Val(Parse(n + 23))
                Map(Val(Parse(1))).Tile(X, y).Fringe2Set = Val(Parse(n + 24))
                Map(Val(Parse(1))).Tile(X, y).F2AnimSet = Val(Parse(n + 25))
                n = n + 26
            Next X
        Next y

        For X = 1 To 15
            Map(Val(Parse(1))).Npc(X) = Val(Parse(n))
            Map(Val(Parse(1))).SpawnX(X) = Val(Parse(n + 1))
            Map(Val(Parse(1))).SpawnY(X) = Val(Parse(n + 2))
            n = n + 3
        Next X

        ' Save the map
        Call SaveLocalMap(Val(Parse(1)))

        ' Check if we get a map from someone else and if we were editing a map cancel it out
        If InEditor Then
            InEditor = False
            frmMapEditor.Visible = False
            frmMapEditor.Visible = False
            frmMirage.Show
' frmMirage.picMapEditor.Visible = False

            If frmMapWarp.Visible Then
                Unload frmMapWarp
            End If

            If frmMapProperties.Visible Then
                Unload frmMapProperties
            End If
        End If

        Exit Sub
    End If

    If casestring = PIn.TileCheck Then
        n = 5
        X = Val(Parse(2))
        y = Val(Parse(3))

        Select Case Val(Parse(4))
            Case 0
                Map(Val(Parse(1))).Tile(X, y).Ground = Val(Parse(n))
                Map(Val(Parse(1))).Tile(X, y).GroundSet = Val(Parse(n + 1))
            Case 1
                Map(Val(Parse(1))).Tile(X, y).Mask = Val(Parse(n))
                Map(Val(Parse(1))).Tile(X, y).MaskSet = Val(Parse(n + 1))
            Case 2
                Map(Val(Parse(1))).Tile(X, y).Anim = Val(Parse(n))
                Map(Val(Parse(1))).Tile(X, y).AnimSet = Val(Parse(n + 1))
            Case 3
                Map(Val(Parse(1))).Tile(X, y).Mask2 = Val(Parse(n))
                Map(Val(Parse(1))).Tile(X, y).Mask2Set = Val(Parse(n + 1))
            Case 4
                Map(Val(Parse(1))).Tile(X, y).M2Anim = Val(Parse(n))
                Map(Val(Parse(1))).Tile(X, y).M2AnimSet = Val(Parse(n + 1))
            Case 5
                Map(Val(Parse(1))).Tile(X, y).Fringe = Val(Parse(n))
                Map(Val(Parse(1))).Tile(X, y).FringeSet = Val(Parse(n + 1))
            Case 6
                Map(Val(Parse(1))).Tile(X, y).FAnim = Val(Parse(n))
                Map(Val(Parse(1))).Tile(X, y).FAnimSet = Val(Parse(n + 1))
            Case 7
                Map(Val(Parse(1))).Tile(X, y).Fringe2 = Val(Parse(n))
                Map(Val(Parse(1))).Tile(X, y).Fringe2Set = Val(Parse(n + 1))
            Case 8
                Map(Val(Parse(1))).Tile(X, y).F2Anim = Val(Parse(n))
                Map(Val(Parse(1))).Tile(X, y).F2AnimSet = Val(Parse(n + 1))
        End Select
        Call SaveLocalMap(Val(Parse(1)))
        Exit Sub
    End If

    If casestring = PIn.TileCheckAttribute Then
        n = 5
        X = Val(Parse(2))
        y = Val(Parse(3))

        Map(Val(Parse(1))).Tile(X, y).Type = Val(Parse(n - 1))
        Map(Val(Parse(1))).Tile(X, y).Data1 = Val(Parse(n))
        Map(Val(Parse(1))).Tile(X, y).Data2 = Val(Parse(n + 1))
        Map(Val(Parse(1))).Tile(X, y).Data3 = Val(Parse(n + 2))
        Map(Val(Parse(1))).Tile(X, y).String1 = Parse(n + 3)
        Map(Val(Parse(1))).Tile(X, y).String2 = Parse(n + 4)
        Map(Val(Parse(1))).Tile(X, y).String3 = Parse(n + 5)
        Call SaveLocalMap(Val(Parse(1)))
        Exit Sub
    End If

    ' :::::::::::::::::::::::::::
    ' :: Map items data packet ::
    ' :::::::::::::::::::::::::::
    If casestring = PIn.MapItemData Then
        n = 1

        For I = 1 To MAX_MAP_ITEMS
            SaveMapItem(I).num = Val(Parse(n))
            SaveMapItem(I).Value = Val(Parse(n + 1))
            SaveMapItem(I).Dur = Val(Parse(n + 2))
            SaveMapItem(I).X = Val(Parse(n + 3))
            SaveMapItem(I).y = Val(Parse(n + 4))

            n = n + 5
        Next I

        Exit Sub
    End If

    ' :::::::::::::::::::::::::
    ' :: Map npc data packet ::
    ' :::::::::::::::::::::::::
    If casestring = PIn.MapNPCData Then
        n = 1

        For I = 1 To 15
            SaveMapNpc(I).num = Val(Parse(n))
            SaveMapNpc(I).X = Val(Parse(n + 1))
            SaveMapNpc(I).y = Val(Parse(n + 2))
            SaveMapNpc(I).Dir = Val(Parse(n + 3))

            n = n + 4
        Next I

        Exit Sub
    End If

    ' :::::::::::::::::::::::::::::::
    ' :: Map send completed packet ::
    ' :::::::::::::::::::::::::::::::
    If casestring = PIn.MapDone Then
        ' Map = SaveMap

        For I = 1 To MAX_MAP_ITEMS
            MapItem(I) = SaveMapItem(I)
        Next I

        For I = 1 To MAX_MAP_NPCS
            MapNpc(I) = SaveMapNpc(I)
        Next I

        GettingMap = False

        ' Play music
        If Trim$(Map(GetPlayerMap(MyIndex)).music) <> "None" Then
            Call MapMusic(Map(GetPlayerMap(MyIndex)).music)
        End If

        If GameWeather = WEATHER_RAINING Then
            Call PlayBGS("rain.wav")
        End If
        If GameWeather = WEATHER_THUNDER Then
            Call PlayBGS("thunder.wav")
        End If

        Exit Sub
    End If

    ' ::::::::::::::::::::
    ' :: Social packets ::
    ' ::::::::::::::::::::
    If (casestring = PIn.SayMessage) Or (casestring = PIn.BroadcastMessage) Or (casestring = PIn.GlobalMessage) Or (casestring = PIn.PlayerMessage) Or (casestring = PIn.MapMessage) Or (casestring = PIn.AdminMessage) Then
        If frmMirage.chkSwearFilter.Value = vbChecked Then
            Parse(1) = SwearFilter_Replace(Parse(1))
        End If

        Call AddText(Parse(1), Val(Parse(2)))
        Exit Sub
    End If

    ' :::::::::::::::::::::::
    ' :: Item spawn packet ::
    ' :::::::::::::::::::::::
    If casestring = PIn.SpawnItem Then
        n = Val(Parse(1))

        MapItem(n).num = Val(Parse(2))
        MapItem(n).Value = Val(Parse(3))
        MapItem(n).Dur = Val(Parse(4))
        MapItem(n).X = Val(Parse(5))
        MapItem(n).y = Val(Parse(6))
        Exit Sub
    End If

    ' ::::::::::::::::::::::::
    ' :: Item editor packet ::
    ' ::::::::::::::::::::::::
    If (casestring = PIn.ItemEditor) Then
        InItemsEditor = True

        frmIndex.Show
        frmIndex.lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_ITEMS
            frmIndex.lstIndex.addItem I & ": " & Trim$(item(I).Name)
        Next I

        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If

    ' ::::::::::::::::::::::::
    ' :: Update item packet ::
    ' ::::::::::::::::::::::::
    If (casestring = PIn.UpdateItem) Then
        n = Val(Parse(1))

        ' Update the item
        item(n).Name = Parse(2)
        item(n).Pic = Val(Parse(3))
        item(n).Type = Val(Parse(4))
        item(n).Data1 = Val(Parse(5))
        item(n).Data2 = Val(Parse(6))
        item(n).Data3 = Val(Parse(7))
        item(n).StrReq = Val(Parse(8))
        item(n).DefReq = Val(Parse(9))
        item(n).SpeedReq = Val(Parse(10))
        item(n).MagicReq = Val(Parse(11))
        item(n).ClassReq = Val(Parse(12))
        item(n).AccessReq = Val(Parse(13))

        item(n).AddHP = Val(Parse(14))
        item(n).AddMP = Val(Parse(15))
        item(n).AddSP = Val(Parse(16))
        item(n).AddSTR = Val(Parse(17))
        item(n).AddDEF = Val(Parse(18))
        item(n).AddMAGI = Val(Parse(19))
        item(n).AddSpeed = Val(Parse(20))
        item(n).AddEXP = Val(Parse(21))
        item(n).desc = Parse(22)
        item(n).AttackSpeed = Val(Parse(23))
        item(n).Price = Val(Parse(24))
        item(n).Stackable = Val(Parse(25))
        item(n).Bound = Val(Parse(26))
        Exit Sub
    End If

    ' ::::::::::::::::::::::
    ' :: Edit item packet :: <- Used for item editor admins only
    ' ::::::::::::::::::::::
    If (casestring = PIn.EditItem) Then
        n = Val(Parse(1))

        ' Update the item
        item(n).Name = Parse(2)
        item(n).Pic = Val(Parse(3))
        item(n).Type = Val(Parse(4))
        item(n).Data1 = Val(Parse(5))
        item(n).Data2 = Val(Parse(6))
        item(n).Data3 = Val(Parse(7))
        item(n).StrReq = Val(Parse(8))
        item(n).DefReq = Val(Parse(9))
        item(n).SpeedReq = Val(Parse(10))
        item(n).MagicReq = Val(Parse(11))
        item(n).ClassReq = Val(Parse(12))
        item(n).AccessReq = Val(Parse(13))

        item(n).AddHP = Val(Parse(14))
        item(n).AddMP = Val(Parse(15))
        item(n).AddSP = Val(Parse(16))
        item(n).AddSTR = Val(Parse(17))
        item(n).AddDEF = Val(Parse(18))
        item(n).AddMAGI = Val(Parse(19))
        item(n).AddSpeed = Val(Parse(20))
        item(n).AddEXP = Val(Parse(21))
        item(n).desc = Parse(22)
        item(n).AttackSpeed = Val(Parse(23))
        item(n).Price = Val(Parse(24))
        item(n).Stackable = Val(Parse(25))
        item(n).Bound = Val(Parse(26))

        ' Initialize the item editor
        Call ItemEditorInit

        Exit Sub
    End If

    ' :::::::::::::::::::
    ' :: mouse packet  ::
    ' :::::::::::::::::::
    If (casestring = PIn.Mouse) Then
        Player(MyIndex).input = 1
        Exit Sub
    End If

    ' ::::::::::::::::::
    ' ::Weather Packet::
    ' ::::::::::::::::::
    If (casestring = PIn.MapWeather) Then
        If 0 + Val(Parse(1)) <> 0 Then
            Map(Val(Parse(1))).Weather = Val(Parse(2))
            If Val(Parse(1)) = 2 Then
                frmMirage.tmrSnowDrop.Interval = Val(Parse(3))
            ElseIf Val(Parse(1)) = 1 Then
                frmMirage.tmrRainDrop.Interval = Val(Parse(3))
            End If
        End If
        Exit Sub
    End If

    ' ::::::::::::::::::::::
    ' :: Npc spawn packet ::
    ' ::::::::::::::::::::::
    If casestring = PIn.SpawnNPC Then
        n = Val(Parse(1))

        MapNpc(n).num = Val(Parse(2))
        MapNpc(n).X = Val(Parse(3))
        MapNpc(n).y = Val(Parse(4))
        MapNpc(n).Dir = Val(Parse(5))
        MapNpc(n).Big = Val(Parse(6))

        ' Client use only
        MapNpc(n).xOffset = 0
        MapNpc(n).yOffset = 0
        MapNpc(n).Moving = 0
        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Npc dead packet ::
    ' :::::::::::::::::::::
    If casestring = PIn.NPCDead Then
        n = Val(Parse(1))

        MapNpc(n).num = 0
        MapNpc(n).X = 0
        MapNpc(n).y = 0
        MapNpc(n).Dir = 0

        ' Client use only
        MapNpc(n).xOffset = 0
        MapNpc(n).yOffset = 0
        MapNpc(n).Moving = 0
        Exit Sub
    End If

    ' :::::::::::::::::::::::
    ' :: Npc editor packet ::
    ' :::::::::::::::::::::::
    If (casestring = PIn.NPCEditor) Then
        InNpcEditor = True

        frmIndex.Show
        frmIndex.lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_NPCS
            frmIndex.lstIndex.addItem I & ": " & Trim$(Npc(I).Name)
        Next I

        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If

    ' :::::::::::::::::::::::
    ' :: Update npc packet ::
    ' :::::::::::::::::::::::
    If (casestring = PIn.UpdateNPC) Then
        n = Val(Parse(1))

        ' Update the item
        Npc(n).Name = Parse(2)
        Npc(n).AttackSay = vbNullString
        Npc(n).Sprite = Val(Parse(3))
        Npc(n).SpriteSize = Val(Parse(4))
        
        ' That's all well and good but it also resets our NPC - Pickle
        ' Npc(n).SpawnSecs = 0
        ' Npc(n).Behavior = 0
        ' Npc(n).Range = 0
        ' For i = 1 To MAX_NPC_DROPS
        ' Npc(n).ItemNPC(i).chance = 0
        ' Npc(n).ItemNPC(i).ItemNum = 0
        ' Npc(n).ItemNPC(i).ItemValue = 0
        ' Next i
        ' Npc(n).STR = 0
        ' Npc(n).DEF = 0
        ' Npc(n).speed = 0
        ' Npc(n).MAGI = 0
        
        Npc(n).Big = Val(Parse(5))
        Npc(n).MaxHP = Val(Parse(6))
        ' Npc(n).Exp = 0
        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Edit npc packet ::
    ' :::::::::::::::::::::
    If (casestring = PIn.EditNPC) Then
        n = Val(Parse(1))

        ' Update the npc
        Npc(n).Name = Parse(2)
        Npc(n).AttackSay = Parse(3)
        Npc(n).Sprite = Val(Parse(4))
        Npc(n).SpawnSecs = Val(Parse(5))
        Npc(n).Behavior = Val(Parse(6))
        Npc(n).Range = Val(Parse(7))
        Npc(n).STR = Val(Parse(8))
        Npc(n).DEF = Val(Parse(9))
        Npc(n).Speed = Val(Parse(10))
        Npc(n).MAGI = Val(Parse(11))
        Npc(n).Big = Val(Parse(12))
        Npc(n).MaxHP = Val(Parse(13))
        Npc(n).Exp = Val(Parse(14))
        Npc(n).SpawnTime = Val(Parse(15))
        Npc(n).Element = Val(Parse(16))
        Npc(n).SpriteSize = Val(Parse(17))

        ' Call GlobalMsg("At editnpc..." & Npc(n).Element)
        z = 18
        For I = 1 To MAX_NPC_DROPS
            Npc(n).ItemNPC(I).chance = Val(Parse(z))
            Npc(n).ItemNPC(I).ItemNum = Val(Parse(z + 1))
            Npc(n).ItemNPC(I).ItemValue = Val(Parse(z + 2))
            z = z + 3
        Next I

        ' Initialize the npc editor
        Call NpcEditorInit

        Exit Sub
    End If

    ' ::::::::::::::::::::
    ' :: Map key packet ::
    ' ::::::::::::::::::::
    If (casestring = PIn.MapKey) Then
        X = Val(Parse(1))
        y = Val(Parse(2))
        n = Val(Parse(3))

        TempTile(X, y).DoorOpen = n

        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Edit map packet ::
    ' :::::::::::::::::::::
    If (casestring = PIn.EditMap) Then
        Call EditorInit
        Exit Sub
    End If

    ' ::::::::::::::::::::::::
    ' :: Shop editor packet ::
    ' ::::::::::::::::::::::::
    If (casestring = PIn.ShopEditor) Then
        InShopEditor = True

        frmIndex.Show
        frmIndex.lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_SHOPS
            frmIndex.lstIndex.addItem I & ": " & Trim$(Shop(I).Name)
        Next I

        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If

    ' ::::::::::::::::::::::::
    ' :: Update shop packet ::
    ' ::::::::::::::::::::::::
    If (casestring = PIn.UpdateShop) Then
        n = Val(Parse(1))

        ' Update the shop name
        Shop(n).Name = Parse(2)
        Shop(n).FixesItems = Val(Parse(3))
        Shop(n).BuysItems = Val(Parse(4))
        Shop(n).currencyItem = Val(Parse(5))

        M = 6
        ' Get shop items
        For I = 1 To MAX_SHOP_ITEMS
            Shop(n).ShopItem(I).ItemNum = Val(Parse(M))
            Shop(n).ShopItem(I).Amount = Val(Parse(M + 1))
            Shop(n).ShopItem(I).Price = Val(Parse(M + 2))
            M = M + 3
        Next I

        Exit Sub
    End If

    ' ::::::::::::::::::::::
    ' :: Edit shop packet :: <- Used for shop editor admins only
    ' ::::::::::::::::::::::
    If (casestring = PIn.EditShop) Then

        shopNum = Val(Parse(1))

        ' Update the shop
        Shop(shopNum).Name = Parse(2)
        Shop(shopNum).FixesItems = Val(Parse(3))
        Shop(shopNum).BuysItems = Val(Parse(4))
        Shop(shopNum).currencyItem = Val(Parse(5))

        M = 6
        For I = 1 To MAX_SHOP_ITEMS
            Shop(shopNum).ShopItem(I).ItemNum = Val(Parse(M))
            Shop(shopNum).ShopItem(I).Amount = Val(Parse(M + 1))
            Shop(shopNum).ShopItem(I).Price = Val(Parse(M + 2))
            M = M + 3
        Next I

        ' Initialize the shop editor
        Call ShopEditorInit

        Exit Sub
    End If

    ' :::::::::::::::::::::::::
    ' :: Spell editor packet ::
    ' :::::::::::::::::::::::::
    If (casestring = PIn.SpellEditor) Then
        InSpellEditor = True

        frmIndex.Show
        frmIndex.lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_SPELLS
            frmIndex.lstIndex.addItem I & ": " & Trim$(Spell(I).Name)
        Next I

        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If

    ' ::::::::::::::::::::::::
    ' :: Update spell packet ::
    ' ::::::::::::::::::::::::
    If (casestring = PIn.UpdateSpell) Then
        n = Val(Parse(1))

        ' Update the spell name
        Spell(n).Name = Parse(2)
        Exit Sub
    End If

    ' :::::::::::::::::::::::
    ' :: Edit spell packet :: <- Used for spell editor admins only
    ' :::::::::::::::::::::::
    If (casestring = PIn.EditSpell) Then
        n = Val(Parse(1))

        ' Update the spell
        Spell(n).Name = Parse(2)
        Spell(n).ClassReq = Val(Parse(3))
        Spell(n).LevelReq = Val(Parse(4))
        Spell(n).Type = Val(Parse(5))
        Spell(n).Data1 = Val(Parse(6))
        Spell(n).Data2 = Val(Parse(7))
        Spell(n).Data3 = Val(Parse(8))
        Spell(n).MPCost = Val(Parse(9))
        Spell(n).Sound = Val(Parse(10))
        Spell(n).Range = Val(Parse(11))
        Spell(n).SpellAnim = Val(Parse(12))
        Spell(n).SpellTime = Val(Parse(13))
        Spell(n).SpellDone = Val(Parse(14))
        Spell(n).AE = Val(Parse(15))
        Spell(n).Big = Val(Parse(16))
        Spell(n).Element = Val(Parse(17))


        ' Initialize the spell editor
        Call SpellEditorInit

        Exit Sub
    End If

    ' ::::::::::::::::::
    ' :: Trade packet ::
    ' ::::::::::::::::::
    If (casestring = PIn.OpenShop) Then
        shopNum = Val(Parse(1))
        ' Show the shop
        Call GoShop(shopNum)
        Exit Sub
    End If

    ' :::::::::::::::::::
    ' :: Spells packet ::
    ' :::::::::::::::::::
    If (casestring = PIn.Spells) Then

        frmMirage.picPlayerSpells.Visible = True
        frmMirage.lstSpells.Clear

        ' Put spells known in player record
        For I = 1 To MAX_PLAYER_SPELLS
            Player(MyIndex).Spell(I) = Val(Parse(I))
            If Player(MyIndex).Spell(I) <> 0 Then
                frmMirage.lstSpells.addItem I & ": " & Trim$(Spell(Player(MyIndex).Spell(I)).Name)
            Else
                frmMirage.lstSpells.addItem "--- Slot Free ---"
            End If
        Next I

        frmMirage.lstSpells.ListIndex = 0
        Exit Sub
    End If

    ' ::::::::::::::::::::
    ' :: Weather packet ::
    ' ::::::::::::::::::::
    If (casestring = PIn.Weather) Then
        If Val(Parse(1)) = WEATHER_RAINING And GameWeather <> WEATHER_RAINING Then
            Call AddText("You see drops of rain falling from the sky above!", BRIGHTGREEN)
            Call PlayBGS("rain.mp3")
        End If
        If Val(Parse(1)) = WEATHER_THUNDER And GameWeather <> WEATHER_THUNDER Then
            Call AddText("You see thunder in the sky above!", BRIGHTGREEN)
            Call PlayBGS("thunder.mp3")
        End If
        If Val(Parse(1)) = WEATHER_SNOWING And GameWeather <> WEATHER_SNOWING Then
            Call AddText("You see snow falling from the sky above!", BRIGHTGREEN)
        End If

        If Val(Parse(1)) = WEATHER_NONE Then
            If GameWeather = WEATHER_RAINING Then
                Call AddText("The rain beings to calm.", BRIGHTGREEN)
                ' Right now there's no way to stop sounds! I need to add an index. [Mellowz]
            ElseIf GameWeather = WEATHER_SNOWING Then
                Call AddText("The snow is melting away.", BRIGHTGREEN)
            ElseIf GameWeather = WEATHER_THUNDER Then
                Call AddText("The thunder begins to disapear.", BRIGHTGREEN)
                ' Right now there's no way to stop sounds! I need to add an index. [Mellowz]
            End If
        End If
        GameWeather = Val(Parse(1))
        RainIntensity = Val(Parse(2))
        If MAX_RAINDROPS <> RainIntensity Then
            MAX_RAINDROPS = RainIntensity
            ReDim DropRain(1 To MAX_RAINDROPS) As DropRainRec
            ReDim DropSnow(1 To MAX_RAINDROPS) As DropRainRec
        End If
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::::::
    ' :: playername coloring packet ::
    ' ::::::::::::::::::::::::::::::::
    If (casestring = PIn.NameColor) Then
        Player(MyIndex).Color = Val(Parse(1))
        Exit Sub
    End If

    ' :::::::::::::::::::::::
    ' :: image packet      ::
    ' :::::::::::::::::::::::
    If (LCase$(Parse(0)) = PIn.Fog) Then
        rec.Top = Int(Val(Parse(4)))
        rec.Bottom = Int(Val(Parse(5)))
        rec.Left = Int(Val(Parse(6)))
        rec.Right = Int(Val(Parse(7)))
        Call DD_BackBuffer.BltFast(Val(Parse(1)), Val(Parse(2)), DD_TileSurf(Val(Parse(3))), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::
    ' :: Get Online List ::
    ' ::::::::::::::::::::::::::
    If casestring = PIn.OnlineList Then
        frmMirage.lstOnline.Clear

        n = 2
        z = Val(Parse(1))
        For X = n To (z + 1)
            frmMirage.lstOnline.addItem Trim$(Parse(n))
            n = n + 2
        Next X
        Exit Sub
    End If

    ' ::::::::::::::::::::::::
    ' :: Blit Player Damage ::
    ' ::::::::::::::::::::::::
    If casestring = PIn.DrawPlayerDamage Then
        DmgDamage = Val(Parse(1))
        NPCWho = Val(Parse(2))
        DmgTime = GetTickCount
        iii = 0
        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Blit NPC Damage ::
    ' :::::::::::::::::::::
    If casestring = PIn.DrawNPCDamage Then
        NPCDmgDamage = Val(Parse(1))
        NPCDmgTime = GetTickCount
        ii = 0
        Exit Sub
    End If

    ' :::::::::::::::::::::::::::::::::::::
    ' :: Retrieve the player's inventory ::
    ' :::::::::::::::::::::::::::::::::::::
    If casestring = PIn.PrepareTrade Then
        frmPlayerTrade.Items1.Clear
        frmPlayerTrade.Items2.Clear
        For I = 1 To MAX_PLAYER_TRADES
            Trading(I).InvNum = 0
            Trading(I).InvName = vbNullString
            Trading2(I).InvNum = 0
            Trading2(I).InvName = vbNullString
            frmPlayerTrade.Items1.addItem I & ": <Nothing>"
            frmPlayerTrade.Items2.addItem I & ": <Nothing>"
        Next I

        frmPlayerTrade.Items1.ListIndex = 0

        Call UpdateTradeInventory
        frmPlayerTrade.Show vbModeless, frmMirage
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::::
    ' :: Stop trading with player ::
    ' ::::::::::::::::::::::::::::::
    If casestring = PIn.QuitTrade Then
        For I = 1 To MAX_PLAYER_TRADES
            Trading(I).InvNum = 0
            Trading(I).InvName = vbNullString
            Trading2(I).InvNum = 0
            Trading2(I).InvName = vbNullString
        Next I

        frmPlayerTrade.Visible = False
        Exit Sub
    End If

    ' ::::::::::::::::::
    ' :: Disable Time ::
    ' ::::::::::::::::::
    If casestring = PIn.TimeEnabled Then
        If Parse(1) = "True" Then
            frmMirage.lblGameTime.Caption = vbNullString
            frmMirage.lblGameClock.Caption = vbNullString
            frmMirage.lblGameTime.Visible = False
            frmMirage.lblGameClock.Visible = False
            frmMirage.tmrGameClock.Enabled = False
        Else
            frmMirage.lblGameTime.Caption = "It is now:"
            frmMirage.lblGameClock.Caption = vbNullString
            frmMirage.lblGameTime.Visible = True
            frmMirage.lblGameClock.Visible = True
            frmMirage.tmrGameClock.Enabled = True
        End If
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::::
    ' :: Stop trading with player ::
    ' ::::::::::::::::::::::::::::::
    If casestring = PIn.UpdateTradeItem Then
        n = Val(Parse(1))

        Trading2(n).InvNum = Val(Parse(2))
        Trading2(n).InvName = Parse(3)
        Trading2(n).InvAmt = Val(Parse(4))

        If STR(Trading2(n).InvNum) <= 0 Then
            frmPlayerTrade.Items2.List(n - 1) = n & ": <Nothing>"
        Else
            If Trading2(n).InvAmt = 0 Then
                frmPlayerTrade.Items2.List(n - 1) = n & ": " & Trim$(Trading2(n).InvName)
            Else
                frmPlayerTrade.Items2.List(n - 1) = n & ": " & Trim$(Trading2(n).InvName & " [" & Trading2(n).InvAmt & "]")
            End If
        End If
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::::
    ' :: Stop trading with player ::
    ' ::::::::::::::::::::::::::::::
    If casestring = PIn.Trading Then
        n = Val(Parse(1))
        If n = 0 Then
            frmPlayerTrade.Command2.ForeColor = &H0&
        End If
        If n = 1 Then
            frmPlayerTrade.Command2.ForeColor = &HFF00&
        End If
        Exit Sub
    End If

    ' :::::::::::::::::::::::::
    ' :: Chat System Packets ::
    ' :::::::::::::::::::::::::
    If casestring = PIn.PrepareChat Then
        frmPlayerChat.txtChat.Text = vbNullString
        frmPlayerChat.txtSay.Text = vbNullString
        frmPlayerChat.Label1.Caption = "Chatting With: " & Trim$(Player(Val(Parse(1))).Name)

        frmPlayerChat.Show vbModeless, frmMirage
        Exit Sub
    End If

    If casestring = PIn.QuitChat Then
        frmPlayerChat.txtChat.Text = vbNullString
        frmPlayerChat.txtSay.Text = vbNullString
        frmPlayerChat.Visible = False

        Exit Sub
    End If

    If casestring = PIn.SendChat Then
        Dim s As String

        s = vbNewLine & GetPlayerName(Val(Parse(2))) & "> " & Trim$(Parse(1))
        frmPlayerChat.txtChat.SelStart = Len(frmPlayerChat.txtChat.Text)
        frmPlayerChat.txtChat.SelColor = QBColor(BROWN)
        frmPlayerChat.txtChat.SelText = s
        frmPlayerChat.txtChat.SelStart = Len(frmPlayerChat.txtChat.Text) - 1
        Exit Sub
    End If
' :::::::::::::::::::::::::::::
' :: END Chat System Packets ::
' :::::::::::::::::::::::::::::

    ' :::::::::::::::::::::::
    ' :: Play Sound Packet ::
    ' :::::::::::::::::::::::
    If casestring = PIn.Sound Then
        s = LCase$(Parse(1))
        Select Case Trim$(s)
            Case "attack"
                Call PlaySound("sword.wav")
            Case "critical"
                Call PlaySound("critical.wav")
            Case "miss"
                Call PlaySound("miss.wav")
            Case "key"
                Call PlaySound("key.wav")
            Case "magic"
                Call PlaySound("magic" & Val(Parse(2)) & ".wav")
            Case "warp"
                If FileExists("SFX\warp.wav") Then
                    Call PlaySound("warp.wav")
                End If
            Case "pain"
                Call PlaySound("pain.wav")
            Case "soundattribute"
                Call PlaySound(Trim$(Parse(2)))
        End Select
        Exit Sub
    End If

    ' :::::::::::::::::::::::::::::::::::::::
    ' :: Sprite Change Confirmation Packet ::
    ' :::::::::::::::::::::::::::::::::::::::
    If casestring = PIn.SpriteChange Then
        If Val(Parse(1)) = 1 Then
            I = MsgBox("Would you like to buy this sprite?", 4, "Buying Sprite")
            If I = 6 Then
                Call SendData(POut.BuySprite & END_CHAR)
            End If
        Else
            Call SendData(POut.BuySprite & END_CHAR)
        End If
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::::::::::
    ' :: Change Player Direction Packet ::
    ' ::::::::::::::::::::::::::::::::::::
    If casestring = PIn.ChangeDirection Then
        Player(Val(Parse(2))).Dir = Val(Parse(1))
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::::
    ' :: Flash Movie Event Packet ::
    ' ::::::::::::::::::::::::::::::
    If casestring = PIn.FlashEvent Then
        If LCase$(Mid$(Trim$(Parse(1)), 1, 7)) = "http://" Then
            WriteINI "CONFIG", "Music", 0, App.Path & "\config.ini"
            WriteINI "CONFIG", "Sound", 0, App.Path & "\config.ini"
            Call StopBGM
            Call StopSound
            frmFlash.Flash.LoadMovie 0, Trim$(Parse(1))
            frmFlash.Flash.Play
            frmFlash.tmrPlayFlash.Enabled = True
            frmFlash.Show vbModeless, frmMirage
        ElseIf FileExists("Flashs\" & Trim$(Parse(1))) = True Then
            WriteINI "CONFIG", "Music", 0, App.Path & "\config.ini"
            WriteINI "CONFIG", "Sound", 0, App.Path & "\config.ini"
            Call StopBGM
            Call StopSound
            frmFlash.Flash.LoadMovie 0, App.Path & "\Flashs\" & Trim$(Parse(1))
            frmFlash.Flash.Play
            frmFlash.tmrPlayFlash.Enabled = True
            frmFlash.Show vbModeless, frmMirage
        End If
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::
    ' :: Emoticon editor packet ::
    ' ::::::::::::::::::::::::::::
    If (casestring = PIn.EmoticonEditor) Then
        InEmoticonEditor = True

        frmIndex.Show
        frmIndex.lstIndex.Clear

        For I = 0 To MAX_EMOTICONS
            frmIndex.lstIndex.addItem I & ": " & Trim$(Emoticons(I).Command)
        Next I

        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::
    ' :: Element editor packet ::
    ' ::::::::::::::::::::::::::::
    If (casestring = PIn.ElementEditor) Then
        InElementEditor = True

        frmIndex.Show
        frmIndex.lstIndex.Clear

        For I = 0 To MAX_ELEMENTS
            frmIndex.lstIndex.addItem I & ": " & Trim$(Element(I).Name)
        Next I

        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If

    If (casestring = PIn.EditElement) Then
        n = Val(Parse(1))

        Element(n).Name = Parse(2)
        Element(n).Strong = Val(Parse(3))
        Element(n).Weak = Val(Parse(4))

        Call ElementEditorInit
        Exit Sub
    End If

    If (casestring = PIn.UpdateElement) Then
        n = Val(Parse(1))

        Element(n).Name = Parse(2)
        Element(n).Strong = Val(Parse(3))
        Element(n).Weak = Val(Parse(4))
        Exit Sub
    End If

    If (casestring = PIn.EditEmoticon) Then
        n = Val(Parse(1))

        Emoticons(n).Command = Parse(2)
        Emoticons(n).Pic = Val(Parse(3))

        Call EmoticonEditorInit
        Exit Sub
    End If

    If (casestring = PIn.UpdateEmoticon) Then
        n = Val(Parse(1))

        Emoticons(n).Command = Parse(2)
        Emoticons(n).Pic = Val(Parse(3))
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::
    ' :: Arrow editor packet ::
    ' ::::::::::::::::::::::::::::
    If (casestring = PIn.ArrowEditor) Then
        InArrowEditor = True

        frmIndex.Show
        frmIndex.lstIndex.Clear

        For I = 1 To MAX_ARROWS
            frmIndex.lstIndex.addItem I & ": " & Trim$(Arrows(I).Name)
        Next I

        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If

    If (casestring = PIn.UpdateArrow) Then
        n = Val(Parse(1))

        Arrows(n).Name = Parse(2)
        Arrows(n).Pic = Val(Parse(3))
        Arrows(n).Range = Val(Parse(4))
        Arrows(n).Amount = Val(Parse(5))
        Exit Sub
    End If

    If (casestring = PIn.EditArrow) Then
        n = Val(Parse(1))

        Arrows(n).Name = Parse(2)

        Call ArrowEditorInit
        Exit Sub
    End If

    If (casestring = PIn.UpdateArrow) Then
        n = Val(Parse(1))

        Arrows(n).Name = Parse(2)
        Arrows(n).Pic = Val(Parse(3))
        Arrows(n).Range = Val(Parse(4))
        Arrows(n).Amount = Val(Parse(5))
        Exit Sub
    End If

    If (casestring = PIn.HookShot) Then
        n = Val(Parse(1))
        I = Val(Parse(3))

        Player(n).HookShotAnim = Arrows(Val(Parse(2))).Pic
        Player(n).HookShotTime = GetTickCount
        Player(n).HookShotToX = Val(Parse(4))
        Player(n).HookShotToY = Val(Parse(5))
        Player(n).HookShotX = GetPlayerX(n)
        Player(n).HookShotY = GetPlayerY(n)
        Player(n).HookShotSucces = Val(Parse(6))
        Player(n).HookShotDir = Val(Parse(3))

        Call PlaySound("grapple.wav")
        Call PlaySound("grapple-fire.wav")

        If I = DIR_DOWN Then
            Player(n).HookShotX = GetPlayerX(n)
            Player(n).HookShotY = GetPlayerY(n) + 1
            If Player(n).HookShotX - 1 > MAX_MAPY Then
                Player(n).HookShotX = 0
                Player(n).HookShotY = 0
                Exit Sub
            End If
        End If
        If I = DIR_UP Then
            Player(n).HookShotX = GetPlayerX(n)
            Player(n).HookShotY = GetPlayerY(n) - 1
            If Player(n).HookShotY + 1 < 0 Then
                Player(n).HookShotX = 0
                Player(n).HookShotY = 0
                Exit Sub
            End If
        End If
        If I = DIR_RIGHT Then
            Player(n).HookShotX = GetPlayerX(n) + 1
            Player(n).HookShotY = GetPlayerY(n)
            If Player(n).HookShotX - 1 > MAX_MAPX Then
                Player(n).HookShotX = 0
                Player(n).HookShotY = 0
                Exit Sub
            End If
        End If
        If I = DIR_LEFT Then
            Player(n).HookShotX = GetPlayerX(n) - 1
            Player(n).HookShotY = GetPlayerY(n)
            If Player(n).HookShotX + 1 < 0 Then
                Player(n).Arrow(X).Arrow = 0
                Exit Sub
            End If
        End If
        Exit Sub
    End If

    If (casestring = PIn.CheckArrows) Then
        n = Val(Parse(1))
        z = Val(Parse(2))
        I = Val(Parse(3))

        For X = 1 To MAX_PLAYER_ARROWS
            If Player(n).Arrow(X).Arrow = 0 Then
                Player(n).Arrow(X).Arrow = 1
                Player(n).Arrow(X).ArrowNum = z
                Player(n).Arrow(X).ArrowAnim = Arrows(z).Pic
                Player(n).Arrow(X).ArrowTime = GetTickCount
                Player(n).Arrow(X).ArrowVarX = 0
                Player(n).Arrow(X).ArrowVarY = 0
                Player(n).Arrow(X).ArrowY = GetPlayerY(n)
                Player(n).Arrow(X).ArrowX = GetPlayerX(n)
                Player(n).Arrow(X).ArrowAmount = p

                If I = DIR_DOWN Then
                    Player(n).Arrow(X).ArrowY = GetPlayerY(n) + 1
                    Player(n).Arrow(X).ArrowPosition = 0
                    If Player(n).Arrow(X).ArrowY - 1 > MAX_MAPY Then
                        Player(n).Arrow(X).Arrow = 0
                        Exit Sub
                    End If
                End If
                If I = DIR_UP Then
                    Player(n).Arrow(X).ArrowY = GetPlayerY(n) - 1
                    Player(n).Arrow(X).ArrowPosition = 1
                    If Player(n).Arrow(X).ArrowY + 1 < 0 Then
                        Player(n).Arrow(X).Arrow = 0
                        Exit Sub
                    End If
                End If
                If I = DIR_RIGHT Then
                    Player(n).Arrow(X).ArrowX = GetPlayerX(n) + 1
                    Player(n).Arrow(X).ArrowPosition = 2
                    If Player(n).Arrow(X).ArrowX - 1 > MAX_MAPX Then
                        Player(n).Arrow(X).Arrow = 0
                        Exit Sub
                    End If
                End If
                If I = DIR_LEFT Then
                    Player(n).Arrow(X).ArrowX = GetPlayerX(n) - 1
                    Player(n).Arrow(X).ArrowPosition = 3
                    If Player(n).Arrow(X).ArrowX + 1 < 0 Then
                        Player(n).Arrow(X).Arrow = 0
                        Exit Sub
                    End If
                End If
                Exit For
            End If
        Next X
        Exit Sub
    End If

    If (casestring = PIn.CheckSprite) Then
        n = Val(Parse(1))

        Player(n).Sprite = Val(Parse(2))
        Exit Sub
    End If

    If (casestring = PIn.MapReport) Then
        n = 1

        frmMapReport.lstIndex.Clear
        For I = 1 To MAX_MAPS
            frmMapReport.lstIndex.addItem I & ": " & Trim$(Parse(n))
            n = n + 1
        Next I

        frmMapReport.Show vbModeless, frmMirage
        Exit Sub
    End If

    ' :::::::::::::::::
    ' :: Time packet ::
    ' :::::::::::::::::
    If (casestring = PIn.GameTime) Then
        GameTime = Val(Parse(1))
        If GameTime = TIME_DAY Then
            Call AddText("Day has dawned in this realm.", WHITE)
        Else
            Call AddText("Night has fallen upon the weary eyed nightowls.", WHITE)
        End If
        Exit Sub
    End If

    ' :::::::::::::::::::::::
    ' :: Spell anim packet ::
    ' :::::::::::::::::::::::
    If (casestring = PIn.SpellAnimation) Then
        Dim SpellNum As Long
        SpellNum = Val(Parse(1))

        Spell(SpellNum).SpellAnim = Val(Parse(2))
        Spell(SpellNum).SpellTime = Val(Parse(3))
        Spell(SpellNum).SpellDone = Val(Parse(4))
        Spell(SpellNum).Big = Val(Parse(9))

        Player(Val(Parse(5))).SpellNum = SpellNum

        For I = 1 To MAX_SPELL_ANIM
            If Player(Val(Parse(5))).SpellAnim(I).CastedSpell = NO Then
                Player(Val(Parse(5))).SpellAnim(I).SpellDone = 0
                Player(Val(Parse(5))).SpellAnim(I).SpellVar = 0
                Player(Val(Parse(5))).SpellAnim(I).SpellTime = GetTickCount
                Player(Val(Parse(5))).SpellAnim(I).TargetType = Val(Parse(6))
                Player(Val(Parse(5))).SpellAnim(I).Target = Val(Parse(7))
                Player(Val(Parse(5))).SpellAnim(I).CastedSpell = YES
                Exit For
            End If
        Next I
        Exit Sub
    End If

    ' :::::::::::::::::::::::
    ' :: Script Spell anim packet ::
    ' :::::::::::::::::::::::
    If (casestring = PIn.ScriptSpellAnimation) Then
        Spell(Val(Parse(1))).SpellAnim = Val(Parse(2))
        Spell(Val(Parse(1))).SpellTime = Val(Parse(3))
        Spell(Val(Parse(1))).SpellDone = Val(Parse(4))
        Spell(Val(Parse(1))).Big = Val(Parse(7))


        For I = 1 To MAX_SCRIPTSPELLS
            If ScriptSpell(I).CastedSpell = NO Then
                ScriptSpell(I).SpellNum = Val(Parse(1))
                ScriptSpell(I).SpellDone = 0
                ScriptSpell(I).SpellVar = 0
                ScriptSpell(I).SpellTime = GetTickCount
                ScriptSpell(I).X = Val(Parse(5))
                ScriptSpell(I).y = Val(Parse(6))
                ScriptSpell(I).CastedSpell = YES
                Exit For
            End If
        Next I
        Exit Sub
    End If

    If (casestring = PIn.CheckEmoticons) Then
        n = Val(Parse(1))

        Player(n).EmoticonNum = Val(Parse(2))
        Player(n).EmoticonTime = GetTickCount
        Player(n).EmoticonVar = 0
        Exit Sub
    End If


    If casestring = PIn.LevelUp Then
        Player(Val(Parse(1))).LevelUpT = GetTickCount
        Player(Val(Parse(1))).LevelUp = 1
        Exit Sub
    End If

    If casestring = PIn.DamageDisplay Then
        For I = 1 To MAX_BLT_LINE
            If Val(Parse(1)) = 0 Then
                If BattlePMsg(I).Index <= 0 Then
                    BattlePMsg(I).Index = 1
                    BattlePMsg(I).Msg = Parse(2)
                    BattlePMsg(I).Color = Val(Parse(3))
                    BattlePMsg(I).Time = GetTickCount
                    BattlePMsg(I).Done = 1
                    BattlePMsg(I).y = 0
                    Exit Sub
                Else
                    BattlePMsg(I).y = BattlePMsg(I).y - 15
                End If
            Else
                If BattleMMsg(I).Index <= 0 Then
                    BattleMMsg(I).Index = 1
                    BattleMMsg(I).Msg = Parse(2)
                    BattleMMsg(I).Color = Val(Parse(3))
                    BattleMMsg(I).Time = GetTickCount
                    BattleMMsg(I).Done = 1
                    BattleMMsg(I).y = 0
                    Exit Sub
                Else
                    BattleMMsg(I).y = BattleMMsg(I).y - 15
                End If
            End If
        Next I

        z = 1
        If Val(Parse(1)) = 0 Then
            For I = 1 To MAX_BLT_LINE
                If I < MAX_BLT_LINE Then
                    If BattlePMsg(I).y < BattlePMsg(I + 1).y Then
                        z = I
                    End If
                Else
                    If BattlePMsg(I).y < BattlePMsg(1).y Then
                        z = I
                    End If
                End If
            Next I

            BattlePMsg(z).Index = 1
            BattlePMsg(z).Msg = Parse(2)
            BattlePMsg(z).Color = Val(Parse(3))
            BattlePMsg(z).Time = GetTickCount
            BattlePMsg(z).Done = 1
            BattlePMsg(z).y = 0
        Else
            For I = 1 To MAX_BLT_LINE
                If I < MAX_BLT_LINE Then
                    If BattleMMsg(I).y < BattleMMsg(I + 1).y Then
                        z = I
                    End If
                Else
                    If BattleMMsg(I).y < BattleMMsg(1).y Then
                        z = I
                    End If
                End If
            Next I

            BattleMMsg(z).Index = 1
            BattleMMsg(z).Msg = Parse(2)
            BattleMMsg(z).Color = Val(Parse(3))
            BattleMMsg(z).Time = GetTickCount
            BattleMMsg(z).Done = 1
            BattleMMsg(z).y = 0
        End If
        Exit Sub
    End If

    If casestring = PIn.ItemBreak Then
        ItemDur(Val(Parse(1))).item = Val(Parse(2))
        ItemDur(Val(Parse(1))).Dur = Val(Parse(3))
        ItemDur(Val(Parse(1))).Done = 1
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::::::::::::::
    ' :: Index player worn equipment packet ::
    ' ::::::::::::::::::::::::::::::::::::::::
    If casestring = PIn.ItemWorn Then
        Player(Val(Parse(1))).Armor = Val(Parse(2))
        Player(Val(Parse(1))).Weapon = Val(Parse(3))
        Player(Val(Parse(1))).Helmet = Val(Parse(4))
        Player(Val(Parse(1))).Shield = Val(Parse(5))
        Player(Val(Parse(1))).legs = Val(Parse(6))
        Player(Val(Parse(1))).Ring = Val(Parse(7))
        Player(Val(Parse(1))).Necklace = Val(Parse(8))
        Exit Sub
    End If

    If casestring = PIn.ScriptTile Then
        frmScript.lblScript.Caption = Parse(1)
        Exit Sub
    End If

    ' ::::::::::::::::::::::
    ' :: Set player speed ::
    ' ::::::::::::::::::::::
    If casestring = PIn.SetSpeed Then
        SetSpeed Parse(1), Val(Parse(2))
        Exit Sub
    End If

    ' ::::::::::::::::::
    ' :: Custom Menu  ::
    ' ::::::::::::::::::
    If (casestring = PIn.ShowCustomMenu) Then
        ' Error handling
        If Not FileExists(Parse(2)) Then
            Call MsgBox(Parse(2) & " not found. Menu loading aborted. Please contact a GM to fix this problem.", vbExclamation)
            Exit Sub
        End If

        CUSTOM_TITLE = Parse(1)
        CUSTOM_IS_CLOSABLE = Val(Parse(3))

        frmCustom1.picBackground.Top = 0
        frmCustom1.picBackground.Left = 0
        frmCustom1.picBackground = LoadPicture(App.Path & Parse(2))
        frmCustom1.Height = PixelsToTwips(24 + frmCustom1.picBackground.Height, 1)
        frmCustom1.Width = PixelsToTwips(6 + frmCustom1.picBackground.Width, 0)
        frmCustom1.Visible = True

        Exit Sub
    End If

    If (casestring = PIn.CloseCustomMenu) Then

        CUSTOM_TITLE = "CLOSED"
        Unload frmCustom1

        Exit Sub
    End If

    If (casestring = PIn.LoadPictureCustomMenu) Then

        CustomIndex = Parse(1)
        strfilename = Parse(2)
        CustomX = Val(Parse(3))
        CustomY = Val(Parse(4))
        
        If CustomIndex > frmCustom1.picCustom.UBound Then
            Load frmCustom1.picCustom(CustomIndex)
        End If

        If strfilename = vbNullString Then
            strfilename = "MEGAUBERBLANKNESSOFUNHOLYPOWER" 'smooth :\    -Pickle
        End If

        If FileExists(strfilename) = True Then
            frmCustom1.picCustom(CustomIndex) = LoadPicture(App.Path & strfilename)
            frmCustom1.picCustom(CustomIndex).Top = CustomY
            frmCustom1.picCustom(CustomIndex).Left = CustomX
            frmCustom1.picCustom(CustomIndex).Visible = True
        Else
            frmCustom1.picCustom(CustomIndex).Picture = LoadPicture()
            frmCustom1.picCustom(CustomIndex).Visible = False
        End If

        Exit Sub
    End If

    If (casestring = PIn.LoadLabelCustomMenu) Then

        CustomIndex = Parse(1)
        strfilename = Parse(2)
        CustomX = Val(Parse(3))
        CustomY = Val(Parse(4))
        customsize = Val(Parse(5))
        customcolour = Val(Parse(6))
        
        If CustomIndex > frmCustom1.BtnCustom.UBound Then
            Load frmCustom1.BtnCustom(CustomIndex)
        End If

        frmCustom1.BtnCustom(CustomIndex).Caption = strfilename
        frmCustom1.BtnCustom(CustomIndex).Top = CustomY
        frmCustom1.BtnCustom(CustomIndex).Left = CustomX
        frmCustom1.BtnCustom(CustomIndex).Font.Bold = True
        frmCustom1.BtnCustom(CustomIndex).Font.Size = customsize
        frmCustom1.BtnCustom(CustomIndex).ForeColor = QBColor(customcolour)
        frmCustom1.BtnCustom(CustomIndex).Visible = True
        frmCustom1.BtnCustom(CustomIndex).Alignment = Parse(7)

        If Parse(8) <= 0 Or Parse(9) <= 0 Then
            frmCustom1.BtnCustom(CustomIndex).AutoSize = True
        Else
            frmCustom1.BtnCustom(CustomIndex).AutoSize = False
            frmCustom1.BtnCustom(CustomIndex).Width = Parse(8)
            frmCustom1.BtnCustom(CustomIndex).Height = Parse(9)
        End If

        Exit Sub
    End If

    If (casestring = PIn.LoadTextboxCustomMenu) Then

        CustomIndex = Parse(1)
        strfilename = Parse(2)
        CustomX = Val(Parse(3))
        CustomY = Val(Parse(4))
        customtext = Parse(5)
        
        If CustomIndex > frmCustom1.txtCustom.UBound Then
            Load frmCustom1.txtCustom(CustomIndex)
            Load frmCustom1.txtcustomOK(CustomIndex)
        End If

        frmCustom1.txtCustom(CustomIndex).Text = customtext
        frmCustom1.txtCustom(CustomIndex).Top = CustomY
        frmCustom1.txtCustom(CustomIndex).Left = strfilename
        frmCustom1.txtCustom(CustomIndex).Width = CustomX - 32
        frmCustom1.txtcustomOK(CustomIndex).Top = CustomY
        frmCustom1.txtcustomOK(CustomIndex).Left = frmCustom1.txtCustom(CustomIndex).Left + frmCustom1.txtCustom(CustomIndex).Width
        frmCustom1.txtcustomOK(CustomIndex).Visible = True
        frmCustom1.txtCustom(CustomIndex).Visible = True

        Exit Sub
    End If

    If (casestring = PIn.LoadInternetWindow) Then
        customtext = Parse(1)
        ' DEBUG STRING
        ' Call AddText(customtext, 15)
        ShellExecute 1, "open", Trim(customtext), vbNullString, vbNullString, 1
        Exit Sub
    End If

    If (casestring = PIn.ReturnCustomBoxMessage) Then
        customsize = Parse(1)

        packet = POut.CustomBoxReturnMessage & SEP_CHAR & frmCustom1.txtCustom(customsize).Text & END_CHAR
        Call SendData(packet)

        Exit Sub
    End If

    Call AddText("Received invalid packet: " & Parse(0), BRIGHTRED)
End Sub

Public Sub Packet_MaxInfo(ByRef Parse() As String)
    Dim I As Long

    ' Set the global configuration values.
    GAME_NAME = Trim$(Parse(1))
    MAX_PLAYERS = CLng(Parse(2))
    MAX_ITEMS = CLng(Parse(3))
    MAX_NPCS = CLng(Parse(4))
    MAX_SHOPS = CLng(Parse(5))
    MAX_SPELLS = CLng(Parse(6))
    MAX_MAPS = CLng(Parse(7))
    MAX_MAP_ITEMS = CLng(Parse(8))
    MAX_MAPX = CLng(Parse(9))
    MAX_MAPY = CLng(Parse(10))
    MAX_EMOTICONS = CLng(Parse(11))
    MAX_ELEMENTS = CLng(Parse(12))
    PaperDoll = CLng(Parse(13))
    SpriteSize = CLng(Parse(14))
    MAX_SCRIPTSPELLS = CLng(Parse(15))
    CUSTOM_PLAYERS = CLng(Parse(16))
    DISPLAY_LEVEL = CLng(Parse(17))
    MAX_PARTY_MEMBERS = CLng(Parse(18))
    STAT1 = Trim$(Parse(19))
    STAT2 = Trim$(Parse(20))
    STAT3 = Trim$(Parse(21))
    STAT4 = Trim$(Parse(22))

    ' ReDim all of the arrays with our new MAX values.
    ReDim Map(0 To MAX_MAPS) As MapRec
    ReDim Player(1 To MAX_PLAYERS) As PlayerRec
    ReDim item(1 To MAX_ITEMS) As ItemRec
    ReDim Npc(1 To MAX_NPCS) As NpcRec
    ReDim MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
    ReDim Shop(1 To MAX_SHOPS) As ShopRec
    ReDim Spell(1 To MAX_SPELLS) As SpellRec
    ReDim Element(0 To MAX_ELEMENTS) As ElementRec
    ReDim Bubble(1 To MAX_PLAYERS) As ChatBubble
    ReDim ScriptBubble(1 To MAX_BUBBLES) As ScriptBubble
    ReDim SaveMapItem(1 To MAX_MAP_ITEMS) As MapItemRec
    ReDim ScriptSpell(1 To MAX_SCRIPTSPELLS) As ScriptSpellAnimRec

    For I = 1 To MAX_MAPS
        ReDim Map(I).Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
        ReDim Map(I).Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
    Next I
        
    ReDim TempTile(0 To MAX_MAPX, 0 To MAX_MAPY) As TempTileRec
    ReDim Emoticons(0 To MAX_EMOTICONS) As EmoRec
    ReDim MapReport(1 To MAX_MAPS) As MapRec

    MAX_SPELL_ANIM = MAX_MAPX * MAX_MAPY

    MAX_BLT_LINE = 6
    ReDim BattlePMsg(1 To MAX_BLT_LINE) As BattleMsgRec
    ReDim BattleMMsg(1 To MAX_BLT_LINE) As BattleMsgRec

    For I = 1 To MAX_PLAYERS
        ReDim Player(I).SpellAnim(1 To MAX_SPELL_ANIM) As SpellAnimRec
    Next I

    ' Reset the emoticon values.
    For I = 0 To MAX_EMOTICONS
        Emoticons(I).Pic = 0
        Emoticons(I).Command = vbNullString
    Next I

    ' Reset the temporary tile values.
    Call ClearTempTile

    ' Reset all player values.
    For I = 1 To MAX_PLAYERS
        Call ClearPlayer(I)
    Next I

    ' Create the backbuffer - We have to do it this way since Eclipse relies
    ' on the MAX_INFO packet for the MAX_MAPX and MAX_MAPY variables.
    Call BackBuffer_Create

    ' Load the swear system.
    Call SwearFilter_Load(App.Path & "\SwearFilter.ini")

    ' Set the application title.
    frmMirage.Caption = Trim$(GAME_NAME)
    App.Title = Trim$(GAME_NAME)

    ' We received and setup all of the required data.
    AllDataReceived = True
End Sub

Public Sub Packet_ClearParty()
    Dim I As Long

    ' Clear all of the party members.
    For I = 1 To MAX_PARTY_MEMBERS
        Player(MyIndex).Party.Member(I) = 0
    Next I
End Sub

Public Sub Packet_UpdateNpcHP(ByVal MapNpcNum As Long, ByVal CurrentHP As Long, ByVal MaxHP As Long)
    ' Set the current NPC HP.
    MapNpc(MapNpcNum).HP = CurrentHP

    ' Set the maximum NPC HP.
    MapNpc(MapNpcNum).MaxHP = MaxHP
End Sub

Public Sub Packet_AlertMessage(ByVal Message As String)
    Dim frm As Form

    ' Hide any active windows if visible.
    For Each frm In Forms
        If frm.Visible Then
            frm.Visible = False
        End If
    Next

    ' Display the main menu form.
    frmMainMenu.Visible = True

    ' Display the message to the user.
    Call MsgBox(Message, vbOKOnly, GAME_NAME)
End Sub

Public Sub Packet_PlainMessage(ByVal Message As String, ByVal FormID As Long)
    ' Hide the status form.
    frmSendGetData.Visible = False

    ' Display the requested form.
    If FormID = 0 Then frmMainMenu.Show
    If FormID = 1 Then frmNewAccount.Show
    If FormID = 2 Then frmDeleteAccount.Show
    If FormID = 3 Then frmLogin.Show
    If FormID = 4 Then frmNewChar.Show
    If FormID = 5 Then frmChars.Show

    ' Display the message to the user.
    Call MsgBox(Message, vbOKOnly, GAME_NAME)
End Sub

Public Sub Packet_CharacterList(ByRef Parse() As String)
    Dim Name As String
    Dim Class As String
    Dim Level As Long
    Dim LoopID As Long
    Dim Count As Long

    ' Hide the status form.
    frmSendGetData.Visible = False

    ' Show the character form.
    frmChars.Visible = True

    ' Clear the character list.
    frmChars.lstChars.Clear

    ' Start the index as the first packet.
    Count = 1

    ' Loop through all of the characters.
    For LoopID = 1 To MAX_CHARS

        ' Get the character data from the packet.
        Name = Parse(Count)
        Class = Parse(Count + 1)
        Level = CLng(Parse(Count + 2))

        ' Display the character information to the user.
        If Trim$(Name) = vbNullString Then
            frmChars.lstChars.addItem "Free Character Slot"
        Else
            frmChars.lstChars.addItem Name & " a level " & Level & " " & Class
        End If

        ' Start the index at the next character.
        Count = Count + 3
    Next LoopID

    ' Set the first available item.
    frmChars.lstChars.ListIndex = 0
End Sub

Public Sub Packet_LoginOK(ByVal Index As Long)
    ' This is the index associated with the client.
    MyIndex = Index

    ' Hide the character form.
    frmChars.Visible = False

    ' Show the status form.
    frmSendGetData.Visible = True

    ' ReDim the party member array. Needs to be moved. [Mellowz]
    ReDim Player(MyIndex).Party.Member(1 To MAX_PARTY_MEMBERS)

    ' Inform the user we're waiting for information.
    Call SetStatus("Receiving game data...")
End Sub

Public Sub Packet_News(ByVal Title As String, ByVal Description As String, ByVal RED As Long, ByVal GREEN As Long, ByVal BLUE As Long)
    ' Write the news title and description to the local news file.
    Call WriteINI("DATA", "News", Title, App.Path & "\News.ini")
    Call WriteINI("DATA", "Desc", Description, App.Path & "\News.ini")

    ' Write the news colors to the local news file.
    Call WriteINI("COLOR", "Red", CStr(RED), App.Path & "\News.ini")
    Call WriteINI("COLOR", "Green", CStr(GREEN), App.Path & "\News.ini")
    Call WriteINI("COLOR", "Blue", CStr(BLUE), App.Path & "\News.ini")

    ' Parse the news and update the GUI.
    Call ParseNews
End Sub

Public Sub Packet_NewCharacterClasses(ByRef Parse() As String)
    Dim LoopID As Long
    Dim Count As Long

    ' Start the index at the first packet.
    Count = 1

    ' Get the maximum amount of classes.
    Max_Classes = CLng(Parse(1))

    ' Get the toggle if we're using classes or not.
    ClassesOn = CLng(Parse(2))

    ' ReDim the class array based on the maximum amount of classes.
    ReDim Class(0 To Max_Classes) As ClassRec

    ' Start the index at the third packet.
    Count = 3

    ' Loop through all of the classes in the packet.
    For LoopID = 0 To Max_Classes

        ' Get the class name.
        Class(LoopID).Name = Parse(Count)

        ' Get the class vitals.
        Class(LoopID).HP = CLng(Parse(Count + 1))
        Class(LoopID).MP = CLng(Parse(Count + 2))
        Class(LoopID).SP = CLng(Parse(Count + 3))

        ' Get the class status points.
        Class(LoopID).STR = CLng(Parse(Count + 4))
        Class(LoopID).DEF = CLng(Parse(Count + 5))
        Class(LoopID).Speed = CLng(Parse(Count + 6))
        Class(LoopID).MAGI = CLng(Parse(Count + 7))

        ' Get the class gender sprites.
        Class(LoopID).MaleSprite = CLng(Parse(Count + 8))
        Class(LoopID).FemaleSprite = CLng(Parse(Count + 9))

        ' Get the class usable state.
        Class(LoopID).Locked = CLng(Parse(Count + 10))

        ' Get the class description.
        Class(LoopID).desc = Parse(Count + 11)

        ' Start the index at the next class.
        Count = Count + 12
    Next LoopID

    ' Hide the status form.
    frmSendGetData.Visible = False

    ' Show the new character form.
    frmNewChar.Visible = True

    ' Clear the class combo box.
    frmNewChar.cmbClass.Clear

    ' Add the class names to the combo box.
    For LoopID = 0 To Max_Classes
        If Class(LoopID).Locked = 0 Then
            frmNewChar.cmbClass.addItem Trim$(Class(LoopID).Name)
        End If
    Next LoopID

    ' Select the top-most class name.
    frmNewChar.cmbClass.ListIndex = 0

    ' Check if classes are enabled, and show the combo box.
    If ClassesOn = 0 Then
        frmNewChar.cmbClass.Visible = False
        frmNewChar.lblClassDesc.Visible = False
    Else
        frmNewChar.cmbClass.Visible = True
        frmNewChar.lblClassDesc.Visible = True
    End If

    ' Display the class vitals to the user.
    frmNewChar.lblHP.Caption = CStr(Class(0).HP)
    frmNewChar.lblMP.Caption = CStr(Class(0).MP)
    frmNewChar.lblSP.Caption = CStr(Class(0).SP)

    ' Display the class status points to the user.
    frmNewChar.lblSTR.Caption = CStr(Class(0).STR)
    frmNewChar.lblDEF.Caption = CStr(Class(0).DEF)
    frmNewChar.lblSPEED.Caption = CStr(Class(0).Speed)
    frmNewChar.lblMAGI.Caption = CStr(Class(0).MAGI)

    ' Display the class description to the user.
    frmNewChar.lblClassDesc.Caption = Class(0).desc
End Sub

Public Sub Packet_ClassData(ByRef Parse() As String)
    Dim LoopID As Long
    Dim Count As Long

    ' Get the maximum amount of classes.
    Max_Classes = CLng(Parse(1))

    ' ReDim the class array based on the maximum amount of classes.
    ReDim Class(0 To Max_Classes) As ClassRec

    ' Start the index at the second packet.
    Count = 2

    For LoopID = 0 To Max_Classes

        ' Get the class name.
        Class(LoopID).Name = Parse(Count)

        ' Get the class vitals.
        Class(LoopID).HP = CLng(Parse(Count + 1))
        Class(LoopID).MP = CLng(Parse(Count + 2))
        Class(LoopID).SP = CLng(Parse(Count + 3))

        ' Get the class status points.
        Class(LoopID).STR = CLng(Parse(Count + 4))
        Class(LoopID).DEF = CLng(Parse(Count + 5))
        Class(LoopID).Speed = CLng(Parse(Count + 6))
        Class(LoopID).MAGI = CLng(Parse(Count + 7))

        ' Get the class usable state.
        Class(LoopID).Locked = CLng(Parse(Count + 8))

        ' Get the class description.
        Class(LoopID).desc = Parse(Count + 9)

        ' Start the index at the next class.
        Count = Count + 10
    Next LoopID
End Sub

Public Sub Packet_GameClock(ByVal Second As Long, ByVal Minute As Long, ByVal Hour As Long, ByVal GSpeed As Long)
    ' Set the game clock.
    Seconds = Second
    Minutes = Minute
    Hours = Hour

    ' Set the game speed.
    Gamespeed = GSpeed

    ' Display the game clock.
    frmMirage.lblGameTime.Caption = "It is now:"
    frmMirage.lblGameTime.Visible = True
End Sub

Public Sub Packet_InGame()
    Call GameInit
    Call GameLoop
End Sub

Public Sub Packet_PlayerInventory(ByRef Parse() As String)
    Dim Index As Long
    Dim InvSlot As Long
    Dim Count As Long

    ' Get the player's index from the packet.
    Index = CLng(Parse(1))

    ' Start the index at the second packet.
    Count = 2

    ' Loop through all of the inventory slots.
    For InvSlot = 1 To MAX_INV

        ' Create the item client-side.
        Call SetPlayerInvItemNum(Index, InvSlot, Val(Parse(Count)))
        Call SetPlayerInvItemValue(Index, InvSlot, Val(Parse(Count + 1)))
        Call SetPlayerInvItemDur(Index, InvSlot, Val(Parse(Count + 2)))

        ' Start the index at the next item.
        Count = Count + 3
    Next InvSlot

    ' Check if the index is this clients index.
    If Index = MyIndex Then

        ' Draw the item.
        For InvSlot = 1 To MAX_INV
            Call Inv_BltItem(InvSlot)
        Next InvSlot

        ' Draw the equipment slots.
        Call UpdateVisInv
    End If
End Sub

Public Sub Packet_PlayerInventoryUpdate(ByVal InvSlot As Long, ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemValue As Long, ByVal ItemDur As Long)
    ' Create the item client-side.
    Call SetPlayerInvItemNum(Index, InvSlot, ItemNum)
    Call SetPlayerInvItemValue(Index, InvSlot, ItemValue)
    Call SetPlayerInvItemDur(Index, InvSlot, ItemDur)

    ' Draw the item if it's this client.
    If Index = MyIndex Then

        ' Draw the item.
        Call Inv_BltItem(InvSlot)

        ' Draw the equipment slots.
        Call UpdateVisInv
    End If
End Sub

Public Sub Packet_PlayerBank(ByRef Parse() As String)
    Dim LoopID As Long
    Dim Count As Long

    ' Start the index at the first packet.
    Count = 1

    ' Loop through all of the item slots.
    For LoopID = 1 To MAX_BANK

        ' Create the item.
        Call SetPlayerBankItemNum(MyIndex, LoopID, CLng(Parse(Count)))
        Call SetPlayerBankItemValue(MyIndex, LoopID, CLng(Parse(Count + 1)))
        Call SetPlayerBankItemDur(MyIndex, LoopID, CLng(Parse(Count + 2)))

        ' Start the index at the next item.
        Count = Count + 3
    Next LoopID

    ' Update the bank if it's open.
    If frmBank.Visible Then Call UpdateBank
End Sub

Public Sub Packet_PlayerBankUpdate(ByVal InvSlot As Long, ByVal ItemNum As Long, ByVal ItemValue As Long, ByVal ItemDur As Long)
    ' Create the item client-side.
    Call SetPlayerBankItemNum(MyIndex, InvSlot, ItemNum)
    Call SetPlayerBankItemValue(MyIndex, InvSlot, ItemValue)
    Call SetPlayerBankItemDur(MyIndex, InvSlot, ItemDur)

    ' Update the bank if it's open.
    If frmBank.Visible Then Call UpdateBank
End Sub

Public Sub Packet_OpenBank()
    Dim LoopID As Long

    ' Clear the temporary player inventory.
    frmBank.lstInventory.Clear

    ' Clear the actual player bank.
    frmBank.lstBank.Clear

    ' Build the new temporary inventory.
    For LoopID = 1 To MAX_INV

        ' Check if the inventory slot has an item in it.
        If GetPlayerInvItemNum(MyIndex, LoopID) > 0 Then

            ' Check if the item type is currency or stackable.
            If item(GetPlayerInvItemNum(MyIndex, LoopID)).Type = ITEM_TYPE_CURRENCY Or item(GetPlayerInvItemNum(MyIndex, LoopID)).Stackable = 1 Then
                frmBank.lstInventory.addItem LoopID & "> " & Trim$(item(GetPlayerInvItemNum(MyIndex, LoopID)).Name) & " (" & GetPlayerInvItemValue(MyIndex, LoopID) & ")"
            Else
                ' Check if the inventory item number is currently equipped.
                If GetPlayerWeaponSlot(MyIndex) = LoopID Or GetPlayerArmorSlot(MyIndex) = LoopID Or GetPlayerHelmetSlot(MyIndex) = LoopID Or GetPlayerShieldSlot(MyIndex) = LoopID Or GetPlayerLegsSlot(MyIndex) = LoopID Or GetPlayerRingSlot(MyIndex) = LoopID Or GetPlayerNecklaceSlot(MyIndex) = LoopID Then
                    frmBank.lstInventory.addItem LoopID & "> " & Trim$(item(GetPlayerInvItemNum(MyIndex, LoopID)).Name) & " (Worn)"
                Else
                    frmBank.lstInventory.addItem LoopID & "> " & Trim$(item(GetPlayerInvItemNum(MyIndex, LoopID)).Name)
                End If
            End If
        Else
            ' The inventory item is blank.
            frmBank.lstInventory.addItem LoopID & "> Empty"
        End If
    Next LoopID

    ' Build the new player bank.
    For LoopID = 1 To MAX_BANK

        ' Check if the inventory slot has an item in it.
        If GetPlayerBankItemNum(MyIndex, LoopID) > 0 Then

            ' Check if the item type is currency or stackable.
            If item(GetPlayerBankItemNum(MyIndex, LoopID)).Type = ITEM_TYPE_CURRENCY Or item(GetPlayerBankItemNum(MyIndex, LoopID)).Stackable = 1 Then
                frmBank.lstBank.addItem LoopID & "> " & Trim$(item(GetPlayerBankItemNum(MyIndex, LoopID)).Name) & " (" & GetPlayerBankItemValue(MyIndex, LoopID) & ")"
            Else
                ' Check if the inventory item number is currently equipped.
                If GetPlayerWeaponSlot(MyIndex) = LoopID Or GetPlayerArmorSlot(MyIndex) = LoopID Or GetPlayerHelmetSlot(MyIndex) = LoopID Or GetPlayerShieldSlot(MyIndex) = LoopID Or GetPlayerLegsSlot(MyIndex) = LoopID Or GetPlayerRingSlot(MyIndex) = LoopID Or GetPlayerNecklaceSlot(MyIndex) = LoopID Then
                    frmBank.lstBank.addItem LoopID & "> " & Trim$(item(GetPlayerBankItemNum(MyIndex, LoopID)).Name) & " (worn)"
                Else
                    frmBank.lstBank.addItem LoopID & "> " & Trim$(item(GetPlayerBankItemNum(MyIndex, LoopID)).Name)
                End If
            End If
        Else
            ' The inventory item is blank.
            frmBank.lstBank.addItem LoopID & "> Empty"
        End If
    Next LoopID

    ' Set the selected item to the first one available.
    frmBank.lstBank.ListIndex = 0

    ' Set the selected item to the first one available.
    frmBank.lstInventory.ListIndex = 0

    ' Display the bank form.
    frmBank.Show vbModal
End Sub

Public Sub Packet_BankMessage(ByVal Message As String)
    ' Set the current message on the bank form.
    frmBank.lblMsg.Caption = Message
End Sub

Public Sub Packet_PlayerWornEQ(ByRef Parse() As String)
    Dim Index As Long

    ' Get the player index.
    Index = CLng(Parse(1))

    ' Set the player's equipment values.
    Call SetPlayerArmorSlot(Index, CLng(Parse(2)))
    Call SetPlayerWeaponSlot(Index, CLng(Parse(3)))
    Call SetPlayerHelmetSlot(Index, CLng(Parse(4)))
    Call SetPlayerShieldSlot(Index, CLng(Parse(5)))
    Call SetPlayerLegsSlot(Index, CLng(Parse(6)))
    Call SetPlayerRingSlot(Index, CLng(Parse(7)))
    Call SetPlayerNecklaceSlot(Index, CLng(Parse(8)))

    ' Update the items currently equipped.
    If Index = MyIndex Then Call UpdateVisInv
End Sub

Public Sub Packet_PlayerPoints(ByVal PointNum As Long)
    ' Set the new player points value.
    Call SetPlayerPOINTS(MyIndex, PointNum)

    ' Display the '+' next to each status name.
    If GetPlayerPOINTS(MyIndex) > 0 Then
        frmMirage.AddSTR.Visible = True
        frmMirage.AddDEF.Visible = True
        frmMirage.AddSPD.Visible = True
        frmMirage.AddMAGI.Visible = True
    Else
        frmMirage.AddSTR.Visible = False
        frmMirage.AddDEF.Visible = False
        frmMirage.AddSPD.Visible = False
        frmMirage.AddMAGI.Visible = False
    End If

    ' Display the current amount of points.
    frmMirage.lblPoints.Caption = CStr(PointNum)
End Sub

Public Sub Packet_CustomSprite(ByVal Index As Long, ByVal H As Long, ByVal B As Long, ByVal L As Long)
    ' Update the player sprite.
    Player(Index).head = H
    Player(Index).body = B
    Player(Index).leg = L
End Sub

Public Sub Packet_PlayerHP(ByVal MaxHP As Long, ByVal HP As Long)
    ' Set the new maximum amount of HP.
    Call SetPlayerMaxHP(MyIndex, MaxHP)

    ' Set the new amount of HP.
    Call SetPlayerHP(MyIndex, HP)

    ' Create the client HP bar.
    If GetPlayerMaxHP(MyIndex) > 0 Then
        frmMirage.shpHP.Width = ((GetPlayerHP(MyIndex)) / (GetPlayerMaxHP(MyIndex))) * 150
        frmMirage.lblHP.Caption = GetPlayerHP(MyIndex) & " / " & GetPlayerMaxHP(MyIndex)
    End If
End Sub

Public Sub Packet_PlayerMP(ByVal MaxMP As Long, ByVal MP As Long)
    ' Set the new maximum amount of MP.
    Call SetPlayerMaxMP(MyIndex, MaxMP)

    ' Set the new amount of MP.
    Call SetPlayerMP(MyIndex, MP)

    ' Create the client MP bar.
    If GetPlayerMaxMP(MyIndex) > 0 Then
        frmMirage.shpMP.Width = ((GetPlayerMP(MyIndex)) / (GetPlayerMaxMP(MyIndex))) * 150
        frmMirage.lblMP.Caption = GetPlayerMP(MyIndex) & " / " & GetPlayerMaxMP(MyIndex)
    End If
End Sub

Public Sub Packet_PlayerSP(ByVal MaxSP As Long, ByVal SP As Long)
    ' Set the new maximum amount of SP.
    Call SetPlayerMaxSP(MyIndex, MaxSP)

    ' Set the new amount of SP.
    Call SetPlayerSP(MyIndex, SP)

    ' Create the client SP bar.
    If GetPlayerMaxSP(MyIndex) > 0 Then
        frmMirage.lblSP.Caption = Int(GetPlayerSP(MyIndex) / GetPlayerMaxSP(MyIndex) * 100) & "%"
    End If
End Sub

Public Sub Packet_PlayerEXP(ByVal cEXP As Long, ByVal tExp As Long)
    ' Set the new amount of experience.
    Call SetPlayerExp(MyIndex, cEXP)

    ' Update the experience bar.
    frmMirage.lblEXP.Caption = CStr(cEXP) & " / " & CStr(tExp)
    frmMirage.shpTNL.Width = (cEXP / tExp) * 150
End Sub

Public Sub Packet_SpeechBubble(ByVal SpeechText As String, ByVal SpeechIndex As Long)
    If frmMirage.chkSwearFilter.Value = vbChecked Then
        SpeechText = SwearFilter_Replace(SpeechText)
    End If
    
    Bubble(SpeechIndex).Text = SpeechText
    Bubble(SpeechIndex).Created = GetTickCount()
End Sub

Public Sub Packet_ScriptBubble(ByVal SpeechIndex As Long, ByVal SpeechText As String, ByVal MapNum As Long, ByVal X As Long, ByVal y As Long, ByVal Color As Long)
    If frmMirage.chkSwearFilter.Value = vbChecked Then
        SpeechText = SwearFilter_Replace(SpeechText)
    End If

    ScriptBubble(SpeechIndex).Text = SpeechText
    ScriptBubble(SpeechIndex).Map = MapNum
    ScriptBubble(SpeechIndex).X = X
    ScriptBubble(SpeechIndex).y = y
    ScriptBubble(SpeechIndex).Colour = Color
    ScriptBubble(SpeechIndex).Created = GetTickCount()
End Sub

Public Sub Packet_PlayerStats(ByVal Strength As Long, ByVal Defense As Long, ByVal Speed As Long, Magic As Long, EXPNextLvl As Long, cEXP As Long, ByVal cLevel As Long)
        Dim SubDef As Long
        Dim SubMagi As Long
        Dim SubSpeed As Long
        Dim SubStr As Long

        If GetPlayerWeaponSlot(MyIndex) > 0 Then
            SubStr = SubStr + item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AddSTR
            SubDef = SubDef + item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AddDEF
            SubMagi = SubMagi + item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AddMAGI
            SubSpeed = SubSpeed + item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AddSpeed
        End If

        If GetPlayerArmorSlot(MyIndex) > 0 Then
            SubStr = SubStr + item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).AddSTR
            SubDef = SubDef + item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).AddDEF
            SubMagi = SubMagi + item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).AddMAGI
            SubSpeed = SubSpeed + item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).AddSpeed
        End If

        If GetPlayerShieldSlot(MyIndex) > 0 Then
            SubStr = SubStr + item(GetPlayerInvItemNum(MyIndex, GetPlayerShieldSlot(MyIndex))).AddSTR
            SubDef = SubDef + item(GetPlayerInvItemNum(MyIndex, GetPlayerShieldSlot(MyIndex))).AddDEF
            SubMagi = SubMagi + item(GetPlayerInvItemNum(MyIndex, GetPlayerShieldSlot(MyIndex))).AddMAGI
            SubSpeed = SubSpeed + item(GetPlayerInvItemNum(MyIndex, GetPlayerShieldSlot(MyIndex))).AddSpeed
        End If

        If GetPlayerHelmetSlot(MyIndex) > 0 Then
            SubStr = SubStr + item(GetPlayerInvItemNum(MyIndex, GetPlayerHelmetSlot(MyIndex))).AddSTR
            SubDef = SubDef + item(GetPlayerInvItemNum(MyIndex, GetPlayerHelmetSlot(MyIndex))).AddDEF
            SubMagi = SubMagi + item(GetPlayerInvItemNum(MyIndex, GetPlayerHelmetSlot(MyIndex))).AddMAGI
            SubSpeed = SubSpeed + item(GetPlayerInvItemNum(MyIndex, GetPlayerHelmetSlot(MyIndex))).AddSpeed
        End If

        If GetPlayerLegsSlot(MyIndex) > 0 Then
            SubStr = SubStr + item(GetPlayerInvItemNum(MyIndex, GetPlayerLegsSlot(MyIndex))).AddSTR
            SubDef = SubDef + item(GetPlayerInvItemNum(MyIndex, GetPlayerLegsSlot(MyIndex))).AddDEF
            SubMagi = SubMagi + item(GetPlayerInvItemNum(MyIndex, GetPlayerLegsSlot(MyIndex))).AddMAGI
            SubSpeed = SubSpeed + item(GetPlayerInvItemNum(MyIndex, GetPlayerLegsSlot(MyIndex))).AddSpeed
        End If

        If GetPlayerRingSlot(MyIndex) > 0 Then
            SubStr = SubStr + item(GetPlayerInvItemNum(MyIndex, GetPlayerRingSlot(MyIndex))).AddSTR
            SubDef = SubDef + item(GetPlayerInvItemNum(MyIndex, GetPlayerRingSlot(MyIndex))).AddDEF
            SubMagi = SubMagi + item(GetPlayerInvItemNum(MyIndex, GetPlayerRingSlot(MyIndex))).AddMAGI
            SubSpeed = SubSpeed + item(GetPlayerInvItemNum(MyIndex, GetPlayerRingSlot(MyIndex))).AddSpeed
        End If

        If GetPlayerNecklaceSlot(MyIndex) > 0 Then
            SubStr = SubStr + item(GetPlayerInvItemNum(MyIndex, GetPlayerNecklaceSlot(MyIndex))).AddSTR
            SubDef = SubDef + item(GetPlayerInvItemNum(MyIndex, GetPlayerNecklaceSlot(MyIndex))).AddDEF
            SubMagi = SubMagi + item(GetPlayerInvItemNum(MyIndex, GetPlayerNecklaceSlot(MyIndex))).AddMAGI
            SubSpeed = SubSpeed + item(GetPlayerInvItemNum(MyIndex, GetPlayerNecklaceSlot(MyIndex))).AddSpeed
        End If

        If SubStr > 0 Then
            frmMirage.lblSTR.Caption = CStr(Strength - SubStr) & " (+" & SubStr & ")"
        Else
            frmMirage.lblSTR.Caption = CStr(Strength)
        End If

        If SubDef > 0 Then
            frmMirage.lblDEF.Caption = CStr(Defense - SubDef) & " (+" & SubDef & ")"
        Else
            frmMirage.lblDEF.Caption = CStr(Defense)
        End If

        If SubSpeed > 0 Then
            frmMirage.lblSPEED.Caption = CStr(Speed - SubSpeed) & " (+" & SubSpeed & ")"
        Else
            frmMirage.lblSPEED.Caption = CStr(Speed)
        End If

        If SubMagi > 0 Then
            frmMirage.lblMAGI.Caption = CStr(Magic - SubMagi) & " (+" & SubMagi & ")"
        Else
            frmMirage.lblMAGI.Caption = CStr(Magic)
        End If

        frmMirage.lblEXP.Caption = CStr(cEXP) & " / " & CStr(EXPNextLvl)

        frmMirage.shpTNL.Width = (cEXP / EXPNextLvl) * 150
        frmMirage.lblLevel.Caption = CStr(cLevel)

        Player(MyIndex).Level = cLevel
End Sub
