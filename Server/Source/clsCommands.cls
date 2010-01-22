VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GetVar(file As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found

    szReturn = vbNullString

    sSpaces = Space(5000)

    file = App.Path & "\" & file

    Call GetPrivateProfileString(Header, Var, szReturn, sSpaces, Len(sSpaces), file)

    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

Public Sub PutVar(file As String, Header As String, Var As String, Value As String)
    On Error GoTo PutVar_Error
    Dim fpath As String
    fpath = App.Path & "\" & file
    Call WritePrivateProfileString(Header, Var, Value, fpath)

    On Error GoTo 0
    Exit Sub

PutVar_Error:
End Sub

Public Sub AdminMsg(ByVal Msg As String, ByVal Color As Byte)
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If GetPlayerAccess(I) > 0 Then
                Call SendDataTo(I, POut.AdminMessage & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR)
            End If
        End If
    Next I
End Sub

Public Sub GlobalMsg(ByVal Msg As String, ByVal Color As Byte)
    Call SendDataToAll(POut.GlobalMessage & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR)
End Sub

Public Sub PlayerMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Byte)
    Call SendDataTo(Index, POut.PlayerMessage & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR)
End Sub

Public Sub MapMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal Color As Byte)
    Call SendDataToMap(MapNum, POut.MapMessage & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR)
End Sub

Public Sub AlertMsg(ByVal Index As Long, ByVal Msg As String)
    Call SendDataTo(Index, POut.AlertMessage & SEP_CHAR & Msg & END_CHAR)
    Call CloseSocket(Index)
End Sub

Public Function GetPlayerLogin(ByVal Index As Long) As String
    GetPlayerLogin = Trim$(Player(Index).Login)
End Function

Public Function GetPlayerName(ByVal Index As Long) As String
    GetPlayerName = Trim$(Player(Index).Char(Player(Index).CharNum).Name)
End Function

Public Function GetPlayerGuild(ByVal Index As Long) As String
    GetPlayerGuild = Trim$(Player(Index).Char(Player(Index).CharNum).Guild)
End Function

Public Function GetPlayerGuildAccess(ByVal Index As Long) As Long
    GetPlayerGuildAccess = Player(Index).Char(Player(Index).CharNum).GuildAccess
End Function

Public Sub SetPlayerGuildAccess(ByVal Index As Long, ByVal GuildAccess As Long)
    Player(Index).Char(Player(Index).CharNum).GuildAccess = GuildAccess
End Sub

Public Sub SetPlayerGuild(ByVal Index As Long, ByVal Guild As String)
    Player(Index).Char(Player(Index).CharNum).Guild = Guild
End Sub

Public Function GetPlayerClass(ByVal Index As Long) As Long
    GetPlayerClass = Player(Index).Char(Player(Index).CharNum).Class
End Function

Public Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)
    Player(Index).Char(Player(Index).CharNum).Class = ClassNum
End Sub

Public Function GetPlayerClassName(ByVal Index As Long) As String
    GetPlayerClassName = Trim$(ClassData(GetPlayerClass(Index)).Name)
End Function

Public Function GetPlayerSprite(ByVal Index As Long) As Long
    GetPlayerSprite = Player(Index).Char(Player(Index).CharNum).Sprite
End Function

Public Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)
    Player(Index).Char(Player(Index).CharNum).Sprite = Sprite
End Sub

Public Function GetPlayerLevel(ByVal Index As Long) As Long
    GetPlayerLevel = Player(Index).Char(Player(Index).CharNum).LEVEL
End Function

Public Sub SetPlayerLevel(ByVal Index As Long, ByVal LEVEL As Long)
    Player(Index).Char(Player(Index).CharNum).LEVEL = LEVEL
End Sub

Public Function GetPlayerNextLevel(ByVal Index As Long) As Long
    GetPlayerNextLevel = Experience(GetPlayerLevel(Index))
End Function

Public Function GetPlayerExp(ByVal Index As Long) As Long
    GetPlayerExp = Player(Index).Char(Player(Index).CharNum).Exp
End Function

Public Sub SetPlayerExp(ByVal Index As Long, ByVal Exp As Long)
    Player(Index).Char(Player(Index).CharNum).Exp = Exp
End Sub

Public Function GetPlayerAccess(ByVal Index As Long) As Long
    GetPlayerAccess = Player(Index).Char(Player(Index).CharNum).Access
End Function

Public Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    Player(Index).Char(Player(Index).CharNum).Access = Access
End Sub

Public Function GetPlayerPK(ByVal Index As Long) As Long
    GetPlayerPK = Player(Index).Char(Player(Index).CharNum).PK
End Function

Public Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)
    Player(Index).Char(Player(Index).CharNum).PK = PK
End Sub

Public Function GetPlayerHP(ByVal Index As Long) As Long
    GetPlayerHP = Player(Index).Char(Player(Index).CharNum).HP
End Function

Public Sub SetPlayerHP(ByVal Index As Long, ByVal HP As Long)
    Player(Index).Char(Player(Index).CharNum).HP = HP

    If GetPlayerHP(Index) < 0 Then
        Player(Index).Char(Player(Index).CharNum).HP = 0
    End If

    If GetPlayerHP(Index) > GetPlayerMaxHP(Index) Then
        Player(Index).Char(Player(Index).CharNum).HP = GetPlayerMaxHP(Index)
    End If

    Call SendHP(Index)
End Sub

Public Function GetPlayerMP(ByVal Index As Long) As Long
    GetPlayerMP = Player(Index).Char(Player(Index).CharNum).MP
End Function

Public Sub SetPlayerMP(ByVal Index As Long, ByVal MP As Long)
    Player(Index).Char(Player(Index).CharNum).MP = MP

    If GetPlayerMP(Index) < 0 Then
        Player(Index).Char(Player(Index).CharNum).MP = 0
    End If

    If GetPlayerMP(Index) > GetPlayerMaxMP(Index) Then
        Player(Index).Char(Player(Index).CharNum).MP = GetPlayerMaxMP(Index)
    End If

    Call SendMP(Index)
End Sub

Public Function GetPlayerSP(ByVal Index As Long) As Long
    GetPlayerSP = Player(Index).Char(Player(Index).CharNum).SP
End Function

Public Sub SetPlayerSP(ByVal Index As Long, ByVal SP As Long)
    Player(Index).Char(Player(Index).CharNum).SP = SP

    If GetPlayerSP(Index) < 0 Then
        Player(Index).Char(Player(Index).CharNum).SP = 0
    End If

    If GetPlayerSP(Index) > GetPlayerMaxSP(Index) Then
        Player(Index).Char(Player(Index).CharNum).SP = GetPlayerMaxSP(Index)
    End If

    Call SendSP(Index)
End Sub

Public Function GetPlayerMaxHP(ByVal Index As Long) As Long
    Dim Add As Long

    If GetPlayerWeaponSlot(Index) > 0 Then
        Add = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AddHP
    End If
    If GetPlayerArmorSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).AddHP
    End If
    If GetPlayerShieldSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).AddHP
    End If
    If GetPlayerHelmetSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).AddHP
    End If
    If GetPlayerLegsSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerLegsSlot(Index))).AddHP
    End If
    If GetPlayerRingSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerRingSlot(Index))).AddHP
    End If
    If GetPlayerNecklaceSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerNecklaceSlot(Index))).AddHP
    End If

    GetPlayerMaxHP = (GetPlayerLevel(Index) * AddHP.LEVEL) + (GetPlayerSTR(Index) * AddHP.STR) + (GetPlayerDEF(Index) * AddHP.DEF) + (GetPlayerMAGI(Index) * AddHP.Magi) + (GetPlayerSPEED(Index) * AddHP.Speed) + Add
End Function

Public Function GetPlayerMaxMP(ByVal Index As Long) As Long
    Dim Add As Long

    If GetPlayerWeaponSlot(Index) > 0 Then
        Add = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AddMP
    End If
    If GetPlayerArmorSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).AddMP
    End If
    If GetPlayerShieldSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).AddMP
    End If
    If GetPlayerHelmetSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).AddMP
    End If
    If GetPlayerLegsSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerLegsSlot(Index))).AddMP
    End If
    If GetPlayerRingSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerRingSlot(Index))).AddMP
    End If
    If GetPlayerNecklaceSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerNecklaceSlot(Index))).AddMP
    End If

    GetPlayerMaxMP = (GetPlayerLevel(Index) * AddMP.LEVEL) + (GetPlayerSTR(Index) * AddMP.STR) + (GetPlayerDEF(Index) * AddMP.DEF) + (GetPlayerMAGI(Index) * AddMP.Magi) + (GetPlayerSPEED(Index) * AddMP.Speed) + Add
End Function

Public Function GetPlayerMaxSP(ByVal Index As Long) As Long
    GetPlayerMaxSP = (Player(Index).Char(Player(Index).CharNum).LEVEL + Int(GetPlayerSPEED(Index) / 2) + ClassData(Player(Index).Char(Player(Index).CharNum).Class).Speed) * 2
End Function

Public Function GetClassMaxHP(ByVal ClassNum As Long) As Long
    GetClassMaxHP = (1 + Int(ClassData(ClassNum).STR / 2) + ClassData(ClassNum).STR) * 2
End Function

Public Function GetClassMaxMP(ByVal ClassNum As Long) As Long
    GetClassMaxMP = (1 + Int(ClassData(ClassNum).Magi / 2) + ClassData(ClassNum).Magi) * 2
End Function

Public Function GetClassMaxSP(ByVal ClassNum As Long) As Long
    GetClassMaxSP = (1 + Int(ClassData(ClassNum).Speed / 2) + ClassData(ClassNum).Speed) * 2
End Function

Public Function GetClassSTR(ByVal ClassNum As Long) As Long
    GetClassSTR = ClassData(ClassNum).STR
End Function

Public Function GetClassDEF(ByVal ClassNum As Long) As Long
    GetClassDEF = ClassData(ClassNum).DEF
End Function

Public Function GetClassSPEED(ByVal ClassNum As Long) As Long
    GetClassSPEED = ClassData(ClassNum).Speed
End Function

Public Function GetClassMAGI(ByVal ClassNum As Long) As Long
    GetClassMAGI = ClassData(ClassNum).Magi
End Function

Public Function GetClassSpawnMap(ByVal ClassNum As Long) As Long
    GetClassSpawnMap = ClassData(ClassNum).Map
End Function

Public Function GetClassSpawnX(ByVal ClassNum As Long) As Long
    GetClassSpawnX = ClassData(ClassNum).X
End Function

Public Function GetClassSpawnY(ByVal ClassNum As Long) As Long
    GetClassSpawnY = ClassData(ClassNum).Y
End Function

Public Function GetPlayerSTR(ByVal Index As Long) As Long
    GetPlayerSTR = Player(Index).Char(Player(Index).CharNum).STR
End Function

Public Sub SetPlayerSTR(ByVal Index As Long, ByVal STR As Long)
    Player(Index).Char(Player(Index).CharNum).STR = STR
End Sub

Public Function GetPlayerDEF(ByVal Index As Long) As Long
    GetPlayerDEF = Player(Index).Char(Player(Index).CharNum).DEF
End Function

Public Sub SetPlayerDEF(ByVal Index As Long, ByVal DEF As Long)
    Player(Index).Char(Player(Index).CharNum).DEF = DEF
End Sub

Public Function GetPlayerSPEED(ByVal Index As Long) As Long
    GetPlayerSPEED = Player(Index).Char(Player(Index).CharNum).Speed
End Function

Public Sub SetPlayerSPEED(ByVal Index As Long, ByVal Speed As Long)
    Player(Index).Char(Player(Index).CharNum).Speed = Speed
End Sub

Public Function GetPlayerMAGI(ByVal Index As Long) As Long
    GetPlayerMAGI = Player(Index).Char(Player(Index).CharNum).Magi
End Function

Public Sub SetPlayerMAGI(ByVal Index As Long, ByVal Magi As Long)
    Player(Index).Char(Player(Index).CharNum).Magi = Magi
End Sub

Public Function GetPlayerPOINTS(ByVal Index As Long) As Long
    GetPlayerPOINTS = Player(Index).Char(Player(Index).CharNum).POINTS
End Function

Public Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)
    Player(Index).Char(Player(Index).CharNum).POINTS = POINTS
End Sub

Public Function GetPlayerMap(ByVal Index As Long) As Long
    GetPlayerMap = Player(Index).Char(Player(Index).CharNum).Map
End Function

Public Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)
    Player(Index).Char(Player(Index).CharNum).Map = MapNum
End Sub

Public Function GetPlayerX(ByVal Index As Long) As Long
    GetPlayerX = Player(Index).Char(Player(Index).CharNum).X
End Function

Public Sub SetPlayerX(ByVal Index As Long, ByVal X As Long)
    Player(Index).Char(Player(Index).CharNum).X = X
End Sub

Public Function GetPlayerY(ByVal Index As Long) As Long
    GetPlayerY = Player(Index).Char(Player(Index).CharNum).Y
End Function

Public Sub SetPlayerY(ByVal Index As Long, ByVal Y As Long)
    Player(Index).Char(Player(Index).CharNum).Y = Y
End Sub

Public Function GetPlayerDir(ByVal Index As Long) As Long
    GetPlayerDir = Player(Index).Char(Player(Index).CharNum).Dir
End Function

Public Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)
    Player(Index).Char(Player(Index).CharNum).Dir = Dir
End Sub

Public Function GetPlayerIP(ByVal Index As Long) As String
    GetPlayerIP = GameServer.Sockets.Item(Index).RemoteAddress
End Function

Public Function GetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemNum = Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).num
End Function

Public Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).num = ItemNum
End Sub

Public Function GetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemValue = Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Value
End Function

Public Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Value = ItemValue
End Sub

Public Function GetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemDur = Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Dur
End Function

Public Sub SetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
    Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Dur = ItemDur
End Sub

Public Function GetPlayerSpell(ByVal Index As Long, ByVal SpellSlot As Long) As Long
    GetPlayerSpell = Player(Index).Char(Player(Index).CharNum).Spell(SpellSlot)
End Function

Public Sub SetPlayerSpell(ByVal Index As Long, ByVal SpellSlot As Long, ByVal SpellNum As Long)
    Player(Index).Char(Player(Index).CharNum).Spell(SpellSlot) = SpellNum
End Sub

Public Function GetPlayerArmorSlot(ByVal Index As Long) As Long
    GetPlayerArmorSlot = Player(Index).Char(Player(Index).CharNum).ArmorSlot
End Function

Public Sub SetPlayerArmorSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).ArmorSlot = InvNum
End Sub

Public Function GetPlayerWeaponSlot(ByVal Index As Long) As Long
    GetPlayerWeaponSlot = Player(Index).Char(Player(Index).CharNum).WeaponSlot
End Function

Public Sub SetPlayerWeaponSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).WeaponSlot = InvNum
End Sub

Public Function GetPlayerHelmetSlot(ByVal Index As Long) As Long
    GetPlayerHelmetSlot = Player(Index).Char(Player(Index).CharNum).HelmetSlot
End Function

Public Sub SetPlayerHelmetSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).HelmetSlot = InvNum
End Sub

Public Function GetPlayerShieldSlot(ByVal Index As Long) As Long
    GetPlayerShieldSlot = Player(Index).Char(Player(Index).CharNum).ShieldSlot
End Function

Public Sub SetPlayerShieldSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).ShieldSlot = InvNum
End Sub

Public Function GetPlayerNecklaceSlot(ByVal Index As Long) As Long
    GetPlayerNecklaceSlot = Player(Index).Char(Player(Index).CharNum).NecklaceSlot
End Function

Public Sub SetPlayerNecklaceSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).NecklaceSlot = InvNum
End Sub

Public Function GetPlayerRingSlot(ByVal Index As Long) As Long
    GetPlayerRingSlot = Player(Index).Char(Player(Index).CharNum).RingSlot
End Function

Public Sub SetPlayerRingSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).RingSlot = InvNum
End Sub

Public Function GetPlayerLegsSlot(ByVal Index As Long) As Long
    GetPlayerLegsSlot = Player(Index).Char(Player(Index).CharNum).LegsSlot
End Function

Public Sub SetPlayerLegsSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).LegsSlot = InvNum
End Sub

Public Function GetMapBootMap(ByVal Index As Long) As Long
    GetMapBootMap = Map(GetPlayerMap(Index)).BootMap
End Function

Public Function GetMapBootX(ByVal Index As Long) As Long
    GetMapBootX = Map(GetPlayerMap(Index)).BootX
End Function

Public Function GetMapBootY(ByVal Index As Long) As Long
    GetMapBootY = Map(GetPlayerMap(Index)).BootY
End Function

Public Sub MutePlayer(ByVal Index As Integer)
    Player(Index).Mute = True
End Sub

Public Sub UnMutePlayer(ByVal Index As Integer)
    Player(Index).Mute = False
End Sub

Public Sub MakeDay()
    GameTime = TIME_DAY

    Hours = 7
    Minutes = 0
    Seconds = 0

    Call SendTimeToAll
End Sub

Public Sub MakeNight()
    GameTime = TIME_NIGHT

    Hours = 21
    Minutes = 0
    Seconds = 0

    Call SendTimeToAll
End Sub

Public Function IsScrolling() As Boolean
    If IS_SCROLLING = 1 Then
        IsScrolling = True
    End If
End Function

Public Function GetMaxPlayers() As Long
    GetMaxPlayers = MAX_PLAYERS
End Function

Public Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal ItemDur As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim I As Long

    ' Check for subscript out of range
    If ItemNum < 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If

    If MapNum < 1 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    If MapItemSlot < 1 Or MapItemSlot > MAX_MAP_ITEMS Then
        Exit Sub
    End If

    I = MapItemSlot

    If I <> 0 Then
        If ItemNum >= 0 Then
            If ItemNum <= MAX_ITEMS Then
                MapItem(MapNum, I).num = ItemNum
                MapItem(MapNum, I).Value = ItemVal
        
                If ItemNum <> 0 Then
                    If (Item(ItemNum).Type >= ITEM_TYPE_WEAPON) And (Item(ItemNum).Type <= ITEM_TYPE_NECKLACE) Then
                        MapItem(MapNum, I).Dur = ItemDur
                    Else
                        MapItem(MapNum, I).Dur = 0
                    End If
                Else
                    MapItem(MapNum, I).Dur = 0
                End If
        
                MapItem(MapNum, I).X = X
                MapItem(MapNum, I).Y = Y

                Call SendDataToMap(MapNum, POut.SpawnItem & SEP_CHAR & I & SEP_CHAR & ItemNum & SEP_CHAR & ItemVal & SEP_CHAR & MapItem(MapNum, I).Dur & SEP_CHAR & X & SEP_CHAR & Y & END_CHAR)
            End If
        End If
    End If
End Sub

Public Function IsConnected(ByVal Index As Long) As Boolean
    On Error Resume Next

    If Not GameServer.Sockets.Item(Index).Socket Is Nothing Then
        IsConnected = True
    End If
End Function

Public Function IsPlaying(ByVal Index As Long) As Boolean
    If IsConnected(Index) Then
        If Player(Index).InGame Then
            IsPlaying = True
        End If
    End If
End Function

Public Sub SendInventory(ByVal Index As Long)
    Dim packet As String
    Dim I As Long

    packet = POut.PlayerInventory & SEP_CHAR & Index & SEP_CHAR
    For I = 1 To MAX_INV
        packet = packet & GetPlayerInvItemNum(Index, I) & SEP_CHAR & GetPlayerInvItemValue(Index, I) & SEP_CHAR & GetPlayerInvItemDur(Index, I) & SEP_CHAR
    Next I
    packet = packet & END_CHAR

    Call SendDataToMap(GetPlayerMap(Index), packet)
End Sub

Public Sub SendInventoryUpdate(ByVal Index As Long, ByVal InvSlot As Long)
    Call SendDataToMap(GetPlayerMap(Index), POut.PlayerInventoryUpdate & SEP_CHAR & InvSlot & SEP_CHAR & Index & SEP_CHAR & GetPlayerInvItemNum(Index, InvSlot) & SEP_CHAR & GetPlayerInvItemValue(Index, InvSlot) & SEP_CHAR & GetPlayerInvItemDur(Index, InvSlot) & SEP_CHAR & Index & END_CHAR)
End Sub

Public Sub SendIndexInventoryFromMap(ByVal Index As Long)
    Dim packet As String
    Dim n As Long
    Dim I As Long
    
    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If GetPlayerMap(I) = GetPlayerMap(Index) Then
                packet = POut.PlayerInventory & SEP_CHAR & I & SEP_CHAR
                For n = 1 To MAX_INV
                    packet = packet & (GetPlayerInvItemNum(I, n) & SEP_CHAR & GetPlayerInvItemValue(I, n) & SEP_CHAR & GetPlayerInvItemDur(I, n) & SEP_CHAR)
                Next n
                packet = packet & END_CHAR

                Call SendDataTo(Index, packet)
            End If
        End If
    Next I
End Sub

Public Sub SendWornEquipment(ByVal Index As Long)
    If IsPlaying(Index) Then
        Call SendDataToMap(GetPlayerMap(Index), POut.PlayerWornEQ & SEP_CHAR & Index & SEP_CHAR & GetPlayerArmorSlot(Index) & SEP_CHAR & GetPlayerWeaponSlot(Index) & SEP_CHAR & GetPlayerHelmetSlot(Index) & SEP_CHAR & GetPlayerShieldSlot(Index) & SEP_CHAR & GetPlayerLegsSlot(Index) & SEP_CHAR & GetPlayerRingSlot(Index) & SEP_CHAR & GetPlayerNecklaceSlot(Index) & END_CHAR)
    End If
End Sub

Public Sub SendHP(ByVal Index As Long)
    Call SendDataTo(Index, POut.PlayerHP & SEP_CHAR & GetPlayerMaxHP(Index) & SEP_CHAR & GetPlayerHP(Index) & END_CHAR)
End Sub

Public Sub SendMP(ByVal Index As Long)
    Call SendDataTo(Index, POut.PlayerMP & SEP_CHAR & GetPlayerMaxMP(Index) & SEP_CHAR & GetPlayerMP(Index) & END_CHAR)
End Sub

Public Sub SendSP(ByVal Index As Long)
    Call SendDataTo(Index, POut.PlayerSP & SEP_CHAR & GetPlayerMaxSP(Index) & SEP_CHAR & GetPlayerSP(Index) & END_CHAR)
End Sub

Public Sub SendPTS(ByVal Index As Long)
    Call SendDataTo(Index, POut.PlayerPoints & SEP_CHAR & GetPlayerPOINTS(Index) & END_CHAR)
End Sub

Public Sub SendStats(ByVal Index As Long)
    Call SendDataTo(Index, POut.PlayerStats & SEP_CHAR & GetPlayerSTR(Index) & SEP_CHAR & GetPlayerDEF(Index) & SEP_CHAR & GetPlayerSPEED(Index) & SEP_CHAR & GetPlayerMAGI(Index) & SEP_CHAR & GetPlayerNextLevel(Index) & SEP_CHAR & GetPlayerExp(Index) & SEP_CHAR & GetPlayerLevel(Index) & END_CHAR)
End Sub

Public Sub Flash(ByVal Index As Long, ByVal flashfile As String)
    Call SendDataTo(Index, POut.FlashEvent & SEP_CHAR & flashfile & END_CHAR)
End Sub

Public Sub PlaySound(ByVal Index As Long, ByVal Sound As String)
    Call SendDataTo(Index, POut.Sound & SEP_CHAR & "soundattribute" & SEP_CHAR & Sound & END_CHAR)
End Sub

Public Sub PlaySoundToMap(ByVal Index As Long, ByVal Sound As String)
    Call SendDataToMap(GetPlayerMap(Index), POut.Sound & SEP_CHAR & "soundattribute" & SEP_CHAR & Sound & END_CHAR)
End Sub

Public Sub SendPlayerData(ByVal Index As Long)
    Dim packet As String

    packet = POut.PlayerData & SEP_CHAR
    packet = packet & Index & SEP_CHAR
    packet = packet & GetPlayerName(Index) & SEP_CHAR
    packet = packet & GetPlayerSprite(Index) & SEP_CHAR
    packet = packet & GetPlayerMap(Index) & SEP_CHAR
    packet = packet & GetPlayerX(Index) & SEP_CHAR
    packet = packet & GetPlayerY(Index) & SEP_CHAR
    packet = packet & GetPlayerDir(Index) & SEP_CHAR
    packet = packet & GetPlayerAccess(Index) & SEP_CHAR
    packet = packet & GetPlayerPK(Index) & SEP_CHAR
    packet = packet & GetPlayerGuild(Index) & SEP_CHAR
    packet = packet & GetPlayerGuildAccess(Index) & SEP_CHAR
    packet = packet & GetPlayerClass(Index) & SEP_CHAR
    packet = packet & GetPlayerHead(Index) & SEP_CHAR
    packet = packet & GetPlayerBody(Index) & SEP_CHAR
    packet = packet & GetPlayerleg(Index) & SEP_CHAR
    packet = packet & GetPlayerPaperdoll(Index) & SEP_CHAR
    packet = packet & END_CHAR

    Call SendDataToMap(GetPlayerMap(Index), packet)
End Sub

Public Sub SendDataTo(ByVal Index As Long, ByVal Data As String)
    Dim dbytes() As Byte

    dbytes = StrConv(Data, vbFromUnicode)

    If IsConnected(Index) Then
        GameServer.Sockets.Item(Index).WriteBytes dbytes
        DoEvents
    End If
End Sub

Public Sub SendDataToAll(ByVal Data As String)
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            Call SendDataTo(I, Data)
        End If
    Next I
End Sub

Public Sub SendDataToAllBut(ByVal Index As Long, ByVal Data As String)
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If I <> Index Then
                Call SendDataTo(I, Data)
            End If
        End If
    Next I
End Sub

Public Sub SendDataToMap(ByVal MapNum As Long, ByVal Data As String)
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If GetPlayerMap(I) = MapNum Then
                Call SendDataTo(I, Data)
            End If
        End If
    Next I
End Sub

Public Sub SendDataToMapBut(ByVal Index As Long, ByVal MapNum As Long, ByVal Data As String)
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If I <> Index Then
                If GetPlayerMap(I) = MapNum Then
                    Call SendDataTo(I, Data)
                End If
            End If
        End If
    Next I
End Sub

Public Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    Player(Index).Char(Player(Index).CharNum).Name = Name
End Sub

Public Function GetPlayerCharNum(ByVal Index As Long) As Long
    GetPlayerCharNum = Player(Index).CharNum
End Function

Public Function FindPlayer(ByVal Name As String) As Long
    Dim I As Long

    Name = LCase$(Name)

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If Len(GetPlayerName(I)) >= Len(Name) Then
                If LCase$(GetPlayerName(I)) = Name Then
                    FindPlayer = I
                    Exit Function
                End If
            End If
        End If
    Next I
End Function

Public Sub PlayerWarp(ByVal Index As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim OldMap As Long

    On Error GoTo WarpErr

    If Not IsPlaying(Index) Then Exit Sub

    If MapNum < 1 Or MapNum > MAX_MAPS Then Exit Sub

    OldMap = GetPlayerMap(Index)

    If Not OldMap = MapNum Then
        Call SendLeaveMap(Index, OldMap)
    End If

    Call SetPlayerMap(Index, MapNum)
    Call SetPlayerX(Index, X)
    Call SetPlayerY(Index, Y)

    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
    End If

    PlayersOnMap(MapNum) = YES

    Call SendDataToMap(GetPlayerMap(Index), POut.Sound & SEP_CHAR & "warp" & END_CHAR)

    Player(Index).GettingMap = YES

    Call SendDataTo(Index, POut.CheckForMap & SEP_CHAR & MapNum & SEP_CHAR & Map(MapNum).Revision & END_CHAR)

    Call SendInventory(Index)
    Call SendIndexInventoryFromMap(Index)
    Call SendIndexWornEquipmentFromMap(Index)

    If SCRIPTING = 1 Then
        MyScript.ExecuteStatement "Scripts\Main.txt", "OnMapLoad " & Index & "," & OldMap & "," & MapNum
    End If
    
    Exit Sub

WarpErr:
    Call AddLog("PlayerWarp error for player index " & Index & " on map " & GetPlayerMap(Index) & ".", "logs\ErrorLog.txt")
End Sub

Public Sub JoinWarp(ByVal Index As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim OldMap As Long

    If Not IsPlaying(Index) Then Exit Sub

    If MapNum < 1 Or MapNum > MAX_MAPS Then Exit Sub

    OldMap = GetPlayerMap(Index)

    Call SendLeaveMap(Index, OldMap)

    Call SetPlayerMap(Index, MapNum)
    Call SetPlayerX(Index, X)
    Call SetPlayerY(Index, Y)

    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
    End If

    PlayersOnMap(MapNum) = YES

    Player(Index).GettingMap = YES

    Call SendDataTo(Index, POut.CheckForMap & SEP_CHAR & MapNum & SEP_CHAR & Map(MapNum).Revision & END_CHAR)

    Call SendInventory(Index)
    Call SendIndexWornEquipmentFromMap(Index)
End Sub

Public Sub AddLog(ByVal Text As String, ByVal FileName As String)
    Dim FileID As Long

    If ServerLog Then
        FileName = App.Path & "\" & FileName

        If FileExists(FileName) Then
            FileID = FreeFile

            Open FileName For Output As #FileID
                Print #FileID, Time & ": " & Text
            Close #FileID
        End If
    End If
End Sub

Public Sub HackingAttempt(ByVal Index As Long, ByVal Reason As String)
    If IsPlaying(Index) Then
        Call AdminMsg(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has been booted for (" & Reason & ")", WHITE)
        Call AddLog(GetPlayerName(Index) & " was kicked for '" & Reason & "'.", "Logs\HackAttempt.txt")

        Call AlertMsg(Index, "You have lost your connection with " & GAME_NAME & ".")
    End If
End Sub

Public Sub BattleMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Byte, ByVal Side As Long)
    Call SendDataTo(Index, POut.DamageDisplay & SEP_CHAR & Side & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR)
End Sub

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    Rand = Int((High - Low + 1) * Rnd) + Low
End Function

Public Function GetPlayerBankItemNum(ByVal Index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemNum = Player(Index).Char(Player(Index).CharNum).Bank(BankSlot).num
End Function

Public Sub SetPlayerBankItemNum(ByVal Index As Long, ByVal BankSlot As Long, ByVal ItemNum As Long)
    Player(Index).Char(Player(Index).CharNum).Bank(BankSlot).num = ItemNum
    Call SendBankUpdate(Index, BankSlot)
End Sub

Public Function GetPlayerBankItemValue(ByVal Index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemValue = Player(Index).Char(Player(Index).CharNum).Bank(BankSlot).Value
End Function

Public Sub SetPlayerBankItemValue(ByVal Index As Long, ByVal BankSlot As Long, ByVal ItemValue As Long)
    Player(Index).Char(Player(Index).CharNum).Bank(BankSlot).Value = ItemValue
    Call SendBankUpdate(Index, BankSlot)
End Sub

Public Function GetPlayerBankItemDur(ByVal Index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemDur = Player(Index).Char(Player(Index).CharNum).Bank(BankSlot).Dur
End Function

Public Sub SetPlayerBankItemDur(ByVal Index As Long, ByVal BankSlot As Long, ByVal ItemDur As Long)
    Player(Index).Char(Player(Index).CharNum).Bank(BankSlot).Dur = ItemDur
End Sub

Public Function GetPlayerTarget(ByVal Index As Long) As Long
    If Player(Index).TargetType = TARGET_TYPE_PLAYER Then
        GetPlayerTarget = Player(Index).Target
    Else
        GetPlayerTarget = -1
    End If
End Function

Public Sub SetTimer(ByVal Name As String, ByVal Interval As Long)
    Call modGameLogic.TimerCreate(Name, Interval)
End Sub

Public Function GetTimer(ByVal Name As String) As Long
    GetTimer = modGameLogic.TimerGetTime(Name)
End Function

Public Sub RemoveTimer(ByVal Name As String)
    Call modGameLogic.TimerRemove(Name)
End Sub

Public Sub SetTile(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long, ByVal XSet As Long, ByVal YSet As Long, ByVal tileset As Long, ByVal layer As Long)
    Call ScriptSetTile(MapNum, X, Y, XSet, YSet, tileset, layer)
End Sub

Public Function GetTileX(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long, ByVal layer As Long) As Long
    Select Case layer
        Case 0
            GetTileX = Map(MapNum).Tile(X, Y).Ground - Int(Map(MapNum).Tile(X, Y).Ground / 14) * 14
        Case 1
            GetTileX = Map(MapNum).Tile(X, Y).Mask - Int(Map(MapNum).Tile(X, Y).Mask / 14) * 14
        Case 2
            GetTileX = Map(MapNum).Tile(X, Y).Anim - Int(Map(MapNum).Tile(X, Y).Anim / 14) * 14
        Case 3
            GetTileX = Map(MapNum).Tile(X, Y).Mask2 - Int(Map(MapNum).Tile(X, Y).Mask2 / 14) * 14
        Case 4
            GetTileX = Map(MapNum).Tile(X, Y).M2Anim - Int(Map(MapNum).Tile(X, Y).M2Anim / 14) * 14
        Case 5
            GetTileX = Map(MapNum).Tile(X, Y).Fringe - Int(Map(MapNum).Tile(X, Y).Fringe / 14) * 14
        Case 6
            GetTileX = Map(MapNum).Tile(X, Y).FAnim - Int(Map(MapNum).Tile(X, Y).FAnim / 14) * 14
        Case 7
            GetTileX = Map(MapNum).Tile(X, Y).Fringe2 - Int(Map(MapNum).Tile(X, Y).Fringe2 / 14) * 14
        Case 8
            GetTileX = Map(MapNum).Tile(X, Y).F2Anim - Int(Map(MapNum).Tile(X, Y).F2Anim / 14) * 14
    End Select
End Function

Public Function GetTileY(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long, ByVal layer As Long) As Long
    Select Case layer
        Case 0
            GetTileY = Int(Map(MapNum).Tile(X, Y).Ground / 14)
        Case 1
            GetTileY = Int(Map(MapNum).Tile(X, Y).Mask / 14)
        Case 2
            GetTileY = Int(Map(MapNum).Tile(X, Y).Anim / 14)
        Case 3
            GetTileY = Int(Map(MapNum).Tile(X, Y).Mask2 / 14)
        Case 4
            GetTileY = Int(Map(MapNum).Tile(X, Y).M2Anim / 14)
        Case 5
            GetTileY = Int(Map(MapNum).Tile(X, Y).Fringe / 14)
        Case 6
            GetTileY = Int(Map(MapNum).Tile(X, Y).FAnim / 14)
        Case 7
            GetTileY = Int(Map(MapNum).Tile(X, Y).Fringe2 / 14)
        Case 8
            GetTileY = Int(Map(MapNum).Tile(X, Y).F2Anim / 14)
    End Select
End Function

Public Function GetTileSet(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long, ByVal layer As Long) As Long
    Select Case layer
        Case 0
            GetTileSet = Map(MapNum).Tile(X, Y).GroundSet
        Case 1
            GetTileSet = Map(MapNum).Tile(X, Y).MaskSet
        Case 2
            GetTileSet = Map(MapNum).Tile(X, Y).AnimSet
        Case 3
            GetTileSet = Map(MapNum).Tile(X, Y).Mask2Set
        Case 4
            GetTileSet = Map(MapNum).Tile(X, Y).M2AnimSet
        Case 5
            GetTileSet = Map(MapNum).Tile(X, Y).FringeSet
        Case 6
            GetTileSet = Map(MapNum).Tile(X, Y).FAnimSet
        Case 7
            GetTileSet = Map(MapNum).Tile(X, Y).Fringe2Set
        Case 8
            GetTileSet = Map(MapNum).Tile(X, Y).F2AnimSet
    End Select
End Function

Public Sub SendMap(ByVal MapNum As Long)
    Call SendDataToMap(MapNum, POut.CheckForMap & SEP_CHAR & MapNum & SEP_CHAR & (Map(MapNum).Revision + 1) & END_CHAR)
End Sub

Public Sub SpellAnim(ByVal SpellNum As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
    Call SendDataToMap(MapNum, POut.ScriptSpellAnimation & SEP_CHAR & SpellNum & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & Spell(SpellNum).Big & END_CHAR)
End Sub

Public Sub SetAttribute(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long, ByVal Attrib As Long, ByVal Data1 As Long, ByVal Data2 As Long, ByVal Data3 As Long, ByVal String1 As String, ByVal String2 As String, ByVal String3 As String)
    Call ScriptSetAttribute(MapNum, X, Y, Attrib, Data1, Data2, Data3, String1, String2, String3)
End Sub

Public Function GetAttribute(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long) As Long
    GetAttribute = Map(MapNum).Tile(X, Y).Type
End Function

Public Function GetTileData1(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long) As Long
    GetTileData1 = Map(MapNum).Tile(X, Y).Data1
End Function

Public Function GetTileData2(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long) As Long
    GetTileData2 = Map(MapNum).Tile(X, Y).Data2
End Function

Public Function GetTileData3(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long) As Long
    GetTileData3 = Map(MapNum).Tile(X, Y).Data3
End Function

Public Function GetTileString1(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long) As String
    GetTileString1 = Trim$(Map(MapNum).Tile(X, Y).String1)
End Function

Public Function GetTileString2(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long) As String
    GetTileString2 = Trim$(Map(MapNum).Tile(X, Y).String2)
End Function

Public Function GetTileString3(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long) As String
    GetTileString3 = Trim$(Map(MapNum).Tile(X, Y).String3)
End Function

Public Sub UpdatePaperDoll(ByVal Index As Long)
    Call SendDataToMap(GetPlayerMap(Index), POut.CheckForMap & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & END_CHAR)
End Sub

Public Sub UpdateSprite(ByVal Index As Long)
    Call SendDataToMap(GetPlayerMap(Index), POut.SpriteUpdate & SEP_CHAR & Index & SEP_CHAR & GetPlayerSprite(Index))
End Sub

Public Sub SetMapNpcNumber(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Number As Long)
    MapNPC(MapNum, MapNpcNum).num = Number
End Sub

Public Sub SetMapNpcTarget(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Target As Long)
    MapNPC(MapNum, MapNpcNum).Target = Target
End Sub

Public Sub SetMapName(ByVal MapNum As Long, ByVal MapName As String)
    Map(MapNum).Name = MapName
    Map(MapNum).Revision = Map(MapNum).Revision + 1

    Call SendMap(MapNum)
End Sub

Public Sub SetMapNpcDir(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Direction As Long)
    MapNPC(MapNum, MapNpcNum).Dir = Direction
End Sub

Public Sub SetMapNpcY(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Y As Long)
    MapNPC(MapNum, MapNpcNum).Y = Y
End Sub

Public Sub SetMapNpcX(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal X As Long)
    MapNPC(MapNum, MapNpcNum).X = X
End Sub

Public Sub SetMapNpcHP(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal HitPoints As Long)
    MapNPC(MapNum, MapNpcNum).HP = HitPoints
End Sub

Public Sub SendNPC(ByVal MapNum As Long, ByVal MapNpcNum As Long)
    Call SendDataToMap(MapNum, POut.SpawnNPC & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNPC(MapNum, MapNpcNum).num & SEP_CHAR & MapNPC(MapNum, MapNpcNum).X & SEP_CHAR & MapNPC(MapNum, MapNpcNum).Y & SEP_CHAR & MapNPC(MapNum, MapNpcNum).Dir & SEP_CHAR & NPC(MapNPC(MapNum, MapNpcNum).num).Big & END_CHAR)
End Sub

Public Function GetNpcMaxHP(ByVal NPCnum As Long) As Long
    GetNpcMaxHP = NPC(NPCnum).MAXHP
End Function

Public Function GetMapNpcNumber(ByVal MapNum As Long, ByVal MapNpcNum As Long) As Long
    GetMapNpcNumber = MapNPC(MapNum, MapNpcNum).num
End Function

Public Function GetMapNpcHP(ByVal MapNum As Long, ByVal MapNpcNum As Long) As Long
    GetMapNpcHP = MapNPC(MapNum, MapNpcNum).HP
End Function

Public Function GetNpcName(ByVal NPCnum As Long) As String
    GetNpcName = Trim$(NPC(NPCnum).Name)
End Function

Public Function GetNpcBehavior(ByVal NPCnum As Long) As Long
    GetNpcBehavior = NPC(NPCnum).Behavior
End Function

Public Function GetNpcExp(ByVal NPCnum As Long) As Long
    GetNpcExp = NPC(NPCnum).Exp
End Function

Public Function GetNpcDefense(ByVal NPCnum As Long) As Long
    GetNpcDefense = NPC(NPCnum).DEF
End Function

Public Function GetNpcStrength(ByVal Number As Long) As Long
    GetNpcStrength = NPC(Number).STR
End Function

Public Sub SendIndexWornEquipment(ByVal Index As Long)
    Dim packet As String
    Dim Armor As Long
    Dim Helmet As Long
    Dim Shield As Long
    Dim Weapon As Long
    Dim Legs As Long
    Dim Ring As Long
    Dim Necklace As Long

    If GetPlayerArmorSlot(Index) > 0 Then Armor = GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))
    If GetPlayerHelmetSlot(Index) > 0 Then Helmet = GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))
    If GetPlayerShieldSlot(Index) > 0 Then Shield = GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))
    If GetPlayerWeaponSlot(Index) > 0 Then Weapon = GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))
    If GetPlayerLegsSlot(Index) > 0 Then Legs = GetPlayerInvItemNum(Index, GetPlayerLegsSlot(Index))
    If GetPlayerRingSlot(Index) > 0 Then Ring = GetPlayerInvItemNum(Index, GetPlayerRingSlot(Index))
    If GetPlayerNecklaceSlot(Index) > 0 Then Necklace = GetPlayerInvItemNum(Index, GetPlayerNecklaceSlot(Index))

    Call SendDataToMap(GetPlayerMap(Index), POut.ItemWorn & SEP_CHAR & Index & SEP_CHAR & Armor & SEP_CHAR & Weapon & SEP_CHAR & Helmet & SEP_CHAR & Shield & SEP_CHAR & Legs & SEP_CHAR & Ring & SEP_CHAR & Necklace & END_CHAR)
End Sub

Public Sub SendIndexWornEquipmentFromMap(ByVal Index As Long)
    Dim I As Long
    Dim Armor As Long
    Dim Helmet As Long
    Dim Shield As Long
    Dim Weapon As Long
    Dim Legs As Long
    Dim Ring As Long
    Dim Necklace As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If GetPlayerMap(Index) = GetPlayerMap(I) Then
                Armor = 0
                Helmet = 0
                Shield = 0
                Weapon = 0
                Legs = 0
                Ring = 0
                Necklace = 0

                If GetPlayerArmorSlot(I) > 0 Then Armor = GetPlayerInvItemNum(I, GetPlayerArmorSlot(I))
                If GetPlayerHelmetSlot(I) > 0 Then Helmet = GetPlayerInvItemNum(I, GetPlayerHelmetSlot(I))
                If GetPlayerShieldSlot(I) > 0 Then Shield = GetPlayerInvItemNum(I, GetPlayerShieldSlot(I))
                If GetPlayerWeaponSlot(I) > 0 Then Weapon = GetPlayerInvItemNum(I, GetPlayerWeaponSlot(I))
                If GetPlayerLegsSlot(I) > 0 Then Legs = GetPlayerInvItemNum(I, GetPlayerLegsSlot(I))
                If GetPlayerRingSlot(I) > 0 Then Ring = GetPlayerInvItemNum(I, GetPlayerRingSlot(I))
                If GetPlayerNecklaceSlot(I) > 0 Then Necklace = GetPlayerInvItemNum(I, GetPlayerNecklaceSlot(I))

                Call SendDataTo(Index, POut.ItemWorn & SEP_CHAR & I & SEP_CHAR & Armor & SEP_CHAR & Weapon & SEP_CHAR & Helmet & SEP_CHAR & Shield & SEP_CHAR & Legs & SEP_CHAR & Ring & SEP_CHAR & Necklace & END_CHAR)
            End If
        End If
    Next I
End Sub

Public Function GetPlayersOnMap(ByVal MapNum As Long) As Long
    Dim I As Long
    Dim Count As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If GetPlayerMap(I) = MapNum Then
                Count = Count + 1
            End If
        End If
    Next I

    GetPlayersOnMap = Count
End Function

Public Sub ScriptSpawnNpc(ByVal MapNpcNum As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long, ByVal NPCnum As Long)
    If NPCnum = 0 Then
        Map(MapNum).Revision = Map(MapNum).Revision + 1
        MapNPC(MapNum, MapNpcNum).num = 0
        Map(MapNum).NPC(MapNpcNum) = 0
        MapNPC(MapNum, MapNpcNum).Target = 0
        MapNPC(MapNum, MapNpcNum).HP = 0
        MapNPC(MapNum, MapNpcNum).MP = 0
        MapNPC(MapNum, MapNpcNum).SP = 0
        MapNPC(MapNum, MapNpcNum).Dir = 0
        MapNPC(MapNum, MapNpcNum).X = 0
        MapNPC(MapNum, MapNpcNum).Y = 0

        Call SaveMap(MapNum)
    End If

    Map(MapNum).Revision = Map(MapNum).Revision + 1

    MapNPC(MapNum, MapNpcNum).num = NPCnum
    Map(MapNum).NPC(MapNpcNum) = NPCnum

    MapNPC(MapNum, MapNpcNum).Target = 0

    MapNPC(MapNum, MapNpcNum).HP = GetNpcMaxHP(NPCnum)
    MapNPC(MapNum, MapNpcNum).MP = GetNpcMaxMP(NPCnum)
    MapNPC(MapNum, MapNpcNum).SP = GetNpcMaxSP(NPCnum)

    MapNPC(MapNum, MapNpcNum).Dir = Int(Rnd * 4)

    MapNPC(MapNum, MapNpcNum).X = X
    MapNPC(MapNum, MapNpcNum).Y = Y

    Call SendDataToMap(MapNum, POut.SpawnNPC & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNPC(MapNum, MapNpcNum).num & SEP_CHAR & MapNPC(MapNum, MapNpcNum).X & SEP_CHAR & MapNPC(MapNum, MapNpcNum).Y & SEP_CHAR & MapNPC(MapNum, MapNpcNum).Dir & SEP_CHAR & NPC(MapNPC(MapNum, MapNpcNum).num).Big & END_CHAR)

    Call SaveMap(MapNum)
    
    Call SendDataToMap(MapNum, POut.CheckForMap & SEP_CHAR & MapNum & SEP_CHAR & Map(MapNum).Revision & END_CHAR)
End Sub

Public Sub SpawnItemOnMap(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long, ByVal ItemNum As Long, ByVal Amount As Long, ByVal Durability As Long)
    Dim I As Long

    I = FindOpenMapItemSlot(MapNum)

    If I <> 0 Then
        Call SpawnItemSlot(I, ItemNum, Amount, Durability, MapNum, X, Y)
    Else
        Call SpawnItemSlot(MAX_MAP_ITEMS, ItemNum, Amount, Durability, MapNum, X, Y)
    End If
End Sub

Public Function GetItemName(ByVal ItemNum As Long) As String
    GetItemName = Trim$(Item(ItemNum).Name)
End Function

Public Sub ClearItemSlot(ByVal MapNum As Long, ByVal ItemNum As Long)
    Call SpawnItemSlot(ItemNum, 0, 0, 0, MapNum, MapItem(MapNum, ItemNum).X, MapItem(MapNum, ItemNum).Y)
End Sub

Public Sub SendWhosOnline(ByVal Index As Integer)
    Dim Message As String
    Dim Count As Long
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        If I <> Index Then
            If IsPlaying(I) Then
                Message = Message & GetPlayerName(I) & ", "
                Count = Count + 1
            End If
        End If
    Next I

    If Count = 0 Then
        Message = "There are no other players online."
    Else
        Message = Mid$(Message, 1, Len(Message) - 2)
        Message = "There are " & Count & " other players online: " & Message & "."
    End If

    Call PlayerMsg(Index, Message, WhoColor)
End Sub

Public Sub GoShopping(ByVal Index As Long, ByVal ShopNum As Long)
    Call SendTrade(Index, ShopNum)
End Sub

Public Sub LockPlayer(ByVal Index As Long, ByVal Locked As Long)
    If Locked = 0 Then
        Player(Index).Locked = False
        Call SendDataTo(Index, POut.CheckForMap & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & END_CHAR)
    End If

    If Locked = 1 Then
        Player(Index).Locked = True
        Call SendDataTo(Index, POut.CheckForMap & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & END_CHAR)
    End If
End Sub

Public Function IsPlayerLocked(ByVal Index As Long) As Boolean
    IsPlayerLocked = Player(Index).Locked
End Function

Public Sub SetServerTime(ByVal Time_Hour As Long, ByVal Time_Minute As Long, ByVal Time_Second As Long)
    If Time_Hour <> -1 Then
        If Time_Hour > -1 And Time_Hour < 24 Then
            Hours = Time_Hour
        End If
    End If

    If Time_Minute <> -1 Then
        If Time_Minute > -1 And Time_Minute < 60 Then
            Minutes = Time_Minute
        End If
    End If

    If Time_Second <> -1 Then
        If Time_Second > -1 And Time_Second < 60 Then
            Seconds = Time_Second
        End If
    End If
End Sub

Public Function GetServerTime(ByVal TimeType As Long) As Long
    Select Case TimeType
        Case 0
            GetServerTime = Hours
        Case 1
            GetServerTime = Minutes
        Case 2
            GetServerTime = Seconds
    End Select
End Function

Public Function GetInventoryItemSlot(ByVal Index As Long, ByVal ItemNum As Long, ByVal MinItemValue As Long, ByVal MinItemDur As Long) As Long
    Dim InvSlot As Long

    InvSlot = 1

    Do While InvSlot < MAX_INV + 1
        If GetPlayerInvItemNum(Index, InvSlot) = ItemNum Then
            If GetPlayerInvItemValue(Index, InvSlot) >= MinItemValue Then
                If GetPlayerInvItemDur(Index, InvSlot) >= MinItemDur Then
                    GetInventoryItemSlot = InvSlot
                    Exit Do
                End If
            End If
        End If

        InvSlot = InvSlot + 1
    Loop
End Function

Public Sub GivePlayerItemToSlot(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long, ByVal ItemValue As Long, ByVal ItemDur As Long)
    Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).num = ItemNum
    Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Value = ItemValue
    Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Dur = ItemDur

    Call SendDataToMap(GetPlayerMap(Index), POut.PlayerInventoryUpdate & SEP_CHAR & InvSlot & SEP_CHAR & Index & SEP_CHAR & GetPlayerInvItemNum(Index, InvSlot) & SEP_CHAR & GetPlayerInvItemValue(Index, InvSlot) & SEP_CHAR & GetPlayerInvItemDur(Index, InvSlot) & END_CHAR)
End Sub

Public Function GivePlayerItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemValue As Long, ByVal ItemDur As Long) As Long
    Dim InvSlot As Long

    InvSlot = FindOpenInvSlot(Index, ItemNum)

    If InvSlot <> 0 Then
        Call SetPlayerInvItemNum(Index, InvSlot, ItemNum)

        If Item(GetPlayerInvItemNum(Index, InvSlot)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(Index, InvSlot)).Stackable = 1 Then
            Call SetPlayerInvItemValue(Index, InvSlot, GetPlayerInvItemValue(Index, InvSlot) + ItemValue)
        Else
            Call SetPlayerInvItemValue(Index, InvSlot, 0)
        End If

        Call SetPlayerInvItemDur(Index, InvSlot, ItemDur)
        Call SendInventoryUpdate(Index, InvSlot)

        GivePlayerItem = 1
    End If
End Function

Public Sub TextBubble(ByVal Index As Long, ByVal BubbleIndex As Long, ByVal Message As String, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long, ByVal Color As Long)
    Call SendDataTo(Index, POut.ScriptBubble & SEP_CHAR & BubbleIndex & SEP_CHAR & Message & SEP_CHAR & MapNum & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & Color & END_CHAR)
End Sub

Public Sub MapTextBubble(ByVal BubbleIndex As Long, ByVal Message As String, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long, ByVal Color As Long)
    Call SendDataToMap(MapNum, POut.ScriptBubble & SEP_CHAR & BubbleIndex & SEP_CHAR & Message & SEP_CHAR & MapNum & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & Color & END_CHAR)
End Sub

Public Function GetMapName(ByVal MapNum As Long) As String
    GetMapName = Trim$(Map(MapNum).Name)
End Function

Public Function GetMapUp(ByVal MapNum As Long) As Long
    GetMapUp = Map(MapNum).Up
End Function

Public Function GetMapDown(ByVal MapNum As Long) As Long
    GetMapDown = Map(MapNum).Down
End Function

Public Function GetMapLeft(ByVal MapNum As Long) As Long
    GetMapLeft = Map(MapNum).Left
End Function

Public Function GetMapRight(ByVal MapNum As Long) As Long
    GetMapRight = Map(MapNum).Right
End Function

Public Sub CustomMenuShow(ByVal Index As Long, ByVal Title As String, ByVal FileName As String, ByVal Closable As Long)
    Call SendDataTo(Index, POut.ShowCustomMenu & SEP_CHAR & Title & SEP_CHAR & FileName & SEP_CHAR & Closable & END_CHAR)
End Sub

Public Sub CustomMenuClose(ByVal Index As Long)
    Call SendDataTo(Index, POut.CloseCustomMenu & END_CHAR)
End Sub

Public Sub CustomMenuPicture(ByVal Index As Long, ByVal PicIndex As Long, ByVal FileName As String, ByVal Left As Long, ByVal Top As Long)
    Call SendDataTo(Index, POut.LoadPictureCustomMenu & SEP_CHAR & PicIndex & SEP_CHAR & FileName & SEP_CHAR & Left & SEP_CHAR & Top & END_CHAR)
End Sub

Public Sub CustomMenuLabel(ByVal Index As Long, ByVal PicIndex As Long, ByVal Caption As String, ByVal Left As Long, ByVal Top As Long, ByVal CustomSize As Long, ByVal CustomColour As Long, ByVal Alignment As Long, ByVal Width As Long, ByVal Height As Long)
    Call SendDataTo(Index, POut.LoadLabelCustomMenu & SEP_CHAR & PicIndex & SEP_CHAR & Caption & SEP_CHAR & Left & SEP_CHAR & Top & SEP_CHAR & CustomSize & SEP_CHAR & CustomColour & SEP_CHAR & Alignment & SEP_CHAR & Width & SEP_CHAR & Height & END_CHAR)
End Sub

Public Sub CustomMenuTextBox(ByVal Index As Long, ByVal CustomIndex As Long, ByVal Width As Long, ByVal Left As Long, ByVal Top As Long, ByVal Text As String)
    Call SendDataTo(Index, POut.LoadTextboxCustomMenu & SEP_CHAR & CustomIndex & SEP_CHAR & Left & SEP_CHAR & Width & SEP_CHAR & Top & SEP_CHAR & Text & END_CHAR)
End Sub

Public Function GetPlayerMenuClickTitle(ByVal Index As Long) As String
    GetPlayerMenuClickTitle = Trim$(Player(Index).CustomTitle)
End Function

Public Function GetPlayerMenuClickMsg(ByVal Index As Long) As String
    GetPlayerMenuClickMsg = Trim$(Player(Index).CustomMsg)
End Function

Public Sub InternetLoad(ByVal Index As Long, ByVal Address As String)
    Call SendDataTo(Index, POut.LoadInternetWindow & SEP_CHAR & Address & END_CHAR)
End Sub

Public Sub UpdatePlayerMenuText(ByVal Index As Long, ByVal BoxIndex As Long)
    Call SendDataTo(Index, POut.ReturnCustomBoxMessage & SEP_CHAR & BoxIndex & END_CHAR)
End Sub

Public Function GetPlayerMenuText(ByVal Index As Long) As String
    GetPlayerMenuText = Trim$(Player(Index).CustomMsg)
End Function

Public Sub NpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Direction As Long, ByVal Speed As Long)
    If CanNpcMove(MapNum, MapNpcNum, Direction) Then
        Call NpcMove(MapNum, MapNpcNum, Direction, Speed)
    End If
End Sub

Public Function GetPlayerDirX(Index) As Long
    If GetPlayerDir(Index) = 1 Then
        GetPlayerDirX = GetPlayerX(Index)
    ElseIf GetPlayerDir(Index) = 2 Then
        GetPlayerDirX = GetPlayerX(Index) - 1
    ElseIf GetPlayerDir(Index) = 3 Then
        GetPlayerDirX = GetPlayerX(Index) + 1
    ElseIf GetPlayerDir(Index) = 0 Then
        GetPlayerDirX = GetPlayerX(Index)
    End If
End Function

Public Function GetPlayerDirY(Index) As Long
    If GetPlayerDir(Index) = 1 Then
        GetPlayerDirY = GetPlayerY(Index) + 1
    ElseIf GetPlayerDir(Index) = 2 Then
        GetPlayerDirY = GetPlayerY(Index)
    ElseIf GetPlayerDir(Index) = 3 Then
        GetPlayerDirY = GetPlayerY(Index)
    ElseIf GetPlayerDir(Index) = 0 Then
        GetPlayerDirY = GetPlayerY(Index) - 1
    End If
End Function

Public Sub SetSpeed(ByVal Index As Long, ByVal Movement As String, ByVal Speed As Long)
    If Speed < 1 Or Speed > 32 Then
        Call PlayerMsg(Index, "SetSpeed: Valid range is from 1 to 32. Current speed: " & Speed & ".", BRIGHTRED)
        Exit Sub
    End If

    If Int(Log(Speed) / Log(2)) <> Log(Speed) / Log(2) Then
        Exit Sub
    End If

    Call SendDataTo(Index, POut.SetSpeed & SEP_CHAR & Movement & SEP_CHAR & Speed & END_CHAR)
End Sub

Public Function GetPlayerTargetNpc(ByVal Index As Long) As Long
    If Player(Index).TargetType = TARGET_TYPE_NPC Then
        GetPlayerTargetNpc = Player(Index).TargetNPC
    Else
        GetPlayerTargetNpc = -1
    End If
End Function

Public Function GetNpcX(ByVal MapNum As Long, ByVal MapNpcNum As Long) As Long
    GetNpcX = MapNPC(MapNum, MapNpcNum).X
End Function

Public Function GetNpcY(ByVal MapNum As Long, ByVal MapNpcNum As Long) As Long
    GetNpcY = MapNPC(MapNum, MapNpcNum).Y
End Function

Public Sub SetWeather(ByVal MapNum As Long, ByVal Weather As Long, ByVal Interval As Long)
    Call SendDataToMap(MapNum, POut.MapWeather & SEP_CHAR & MapNum & SEP_CHAR & Weather & SEP_CHAR & Interval & END_CHAR)
End Sub

Public Function GetWeather(ByVal MapNum As Long) As Long
    GetWeather = Map(MapNum).Weather
End Function

Public Sub Image(ByVal Index As Long, ByVal X As Long, ByVal Y As Long, ByVal tileset As Long, ByVal Top As Long, ByVal Height As Long, ByVal Left As Long, ByVal Width As Long)
    Call SendDataTo(Index, POut.Fog & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & tileset & SEP_CHAR & Top & SEP_CHAR & Height & SEP_CHAR & Left & SEP_CHAR & Width & END_CHAR)
End Sub

Public Sub SetPlayerNameColor(ByVal Index As Long, ByVal Color As Long)
    Call SendDataTo(Index, POut.NameColor & SEP_CHAR & Color & END_CHAR)
End Sub

Public Sub LockSpells(ByVal Index As Long, ByVal Locked As Long)
    If Locked = 0 Then
        Player(Index).LockedSpells = False
        Call SendDataTo(Index, POut.CheckForMap & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & END_CHAR)
    End If

    If Locked = 1 Then
        Player(Index).LockedSpells = True
        Call SendDataTo(Index, POut.CheckForMap & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & END_CHAR)
    End If
End Sub

Public Sub LockItems(ByVal Index As Long, ByVal Locked As Long)
    If Locked = 0 Then
        Player(Index).LockedItems = False
        Call SendDataTo(Index, POut.CheckForMap & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & END_CHAR)
    End If

    If Locked = 1 Then
        Player(Index).LockedItems = True
        Call SendDataTo(Index, POut.CheckForMap & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & END_CHAR)
    End If
End Sub

Public Sub LockAttack(ByVal Index As Long, ByVal Locked As Long)
    If Locked = 0 Then
        Player(Index).LockedAttack = False
        Call SendDataTo(Index, POut.CheckForMap & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & END_CHAR)
    End If

    If Locked = 1 Then
        Player(Index).LockedAttack = True
        Call SendDataTo(Index, POut.CheckForMap & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & END_CHAR)
    End If
End Sub

Public Function GetIndexFromMap(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long) As Long
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If GetPlayerMap(I) = MapNum Then
                If GetPlayerX(I) = X Then
                    If GetPlayerY(I) = Y Then
                        GetIndexFromMap = I
                        Exit Function
                    End If
                End If
            End If
        End If
    Next I
End Function

Public Function GetPlayerHead(ByVal Index As Long) As Long
    GetPlayerHead = Player(Index).Char(Player(Index).CharNum).Head
End Function

Public Sub SetPlayerHead(ByVal Index As Long, ByVal Head As Long)
    Player(Index).Char(Player(Index).CharNum).Head = Head
End Sub

Public Function GetPlayerBody(ByVal Index As Long) As Long
    GetPlayerBody = Player(Index).Char(Player(Index).CharNum).Body
End Function

Public Sub SetPlayerBody(ByVal Index As Long, ByVal Body As Long)
    Player(Index).Char(Player(Index).CharNum).Body = Body
End Sub

Public Function GetPlayerleg(ByVal Index As Long) As Long
    GetPlayerleg = Player(Index).Char(Player(Index).CharNum).Leg
End Function

Public Sub SetPlayerLeg(ByVal Index As Long, ByVal Leg As Long)
    Player(Index).Char(Player(Index).CharNum).Leg = Leg
End Sub

Public Sub MovePlayer(ByVal Index As Long, ByVal Direction As Long, ByVal Movement As Long)
    Call PlayerMove(Index, Direction, Movement)
End Sub

Public Sub SavePlayer(ByVal Index As Long)
    Call SavePlayer(Index)
End Sub

Public Function GetPlayerGender(ByVal Index As Long) As Long
    GetPlayerGender = Player(Index).Char(Player(Index).CharNum).Sex
End Function

Public Function GetPlayerPaperdoll(ByVal Index As Long) As Byte
    If Player(Index).InGame Then
        GetPlayerPaperdoll = Player(Index).Char(GetPlayerCharNum(Index)).PAPERDOLL
    End If
End Function

Public Sub SetPlayerPaperdoll(ByVal Index As Long, ByVal Mode As Byte)
    If Mode = 0 Or Mode = 1 Then
        If Player(Index).InGame Then
            Player(Index).Char(GetPlayerCharNum(Index)).PAPERDOLL = Mode
        End If
    End If
End Sub

Public Sub ChangeMapName(ByVal MapNum As Long, ByVal Name As String)
    Map(MapNum).Name = Name

    Call SaveMap(MapNum)
    Call SendMap(MapNum)
End Sub


' Eclipse 2.6 Scripting Commands

Public Sub PlayerMapDropItem(ByVal Index As Long, ByVal InvNum As Long, ByVal Amount As Long)
    Call PlayerMapDropItem(Index, InvNum, Amount)
End Sub

Public Function ScriptGetTickCount() As Long
    ScriptGetTickCount = GetTickCount
End Function


' Eclipse 2.7 Scripting Commands

Public Function GetMaxMapX() As Byte
    GetMaxMapX = MAX_MAPX
End Function

Public Function GetMaxMapY() As Byte
    GetMaxMapY = MAX_MAPY
End Function

Public Sub SetMapMoral(ByVal MapNum As Long, ByVal Moral As Long)
    Map(MapNum).Moral = Moral
End Sub

Public Function GetMapMoral(ByVal MapNum As Long) As Byte
    GetMapMoral = Map(MapNum).Moral
End Function

Public Sub ShopReload(ByVal ShopNum As Long)
    Dim FileID As Integer
    Dim FileName As String
    
    FileName = App.Path & "\Shops\Shop" & CStr(ShopNum) & ".dat"

    FileID = FreeFile

    Open FileName For Binary As #FileID
        Get #FileID, , Shop(ShopNum)
    Close #FileID
End Sub

Public Sub DamageNPC(ByVal Index As Long, ByVal NPCnum As Long, ByVal Damage As Long)
    Call AttackNpc(Index, NPCnum, Damage)
    Call SendDataTo(Index, POut.DrawPlayerDamage & SEP_CHAR & Damage & SEP_CHAR & NPCnum & END_CHAR)
End Sub

Public Sub DamagePlayer(ByVal Attacker As Long, ByVal Victim As Long, ByVal Damage As Long)
    Call AttackPlayer(Attacker, Victim, Damage)
End Sub

Public Sub NPCAttack(ByVal MapNpcNum As Long, ByVal Victim As Long, ByVal Damage As Long)
    Call NpcAttackPlayer(MapNpcNum, Victim, Damage)
End Sub

Public Sub FireArrow(ByVal Index As Long, ByVal ArrowID As Long, ByVal Dir As Long)
    Call SendDataToMap(GetPlayerMap(Index), POut.CheckArrows & SEP_CHAR & Index & SEP_CHAR & ArrowID & SEP_CHAR & Dir & END_CHAR)
End Sub
