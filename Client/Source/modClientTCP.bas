Attribute VB_Name = "modClientTCP"
Option Explicit


Public Type ClientPackOut
    GetClasses As Byte
    NewAccount As Byte
    DeleteAccount As Byte
    AccountLogin As Byte
    GiveMeTheMax As Byte
    AddCharacter As Byte
    DeleteCharacter As Byte
    UseCharacter As Byte
    GuildChangeAccess As Byte
    GuildDisown As Byte
    GuildLeave As Byte
    GuildMake As Byte
    GuildMember As Byte
    GuildTrainee As Byte
    SayMessage As Byte
    EmoteMessage As Byte
    BroadcastMessage As Byte
    GlobalMessage As Byte
    AdminMessage As Byte
    PlayerMessage As Byte
    PlayerMove As Byte
    PlayerDirection As Byte
    UseItem As Byte
    PlayerMouseMove As Byte
    Warp As Byte
    EndShot As Byte
    Attack As Byte
    UseStatPoint As Byte
    PlayerSprite As Byte
    GetStats As Byte
    RequestMap As Byte
    WarpMeTo As Byte
    WarpToMe As Byte
    MapData As Byte
    NeedMap As Byte
    MapGetItem As Byte
    MapDropItem As Byte
    MapRespawn As Byte
    KickPlayer As Byte
    BanList As Byte
    BanListDestroy As Byte
    BanPlayer As Byte
    RequestEditMap As Byte
    RequestEditItem As Byte
    EditItem As Byte
    SaveItem As Byte
    EnableDayNight As Byte
    DayNight As Byte
    RequestEditNPC As Byte
    EditNPC As Byte
    SaveNPC As Byte
    RequestEditShop As Byte
    EditShop As Byte
    SaveShop As Byte
    RequestEditSpell As Byte
    EditSpell As Byte
    SaveSpell As Byte
    ForgetSpell As Byte
    SetAccess As Byte
    WhosOnline As Byte
    OnlineList As Byte
    SetMOTD As Byte
    BuyItem As Byte
    SellItem As Byte
    FixItem As Byte
    Search As Byte
    PlayerChat As Byte
    AcceptChat As Byte
    DeclineChat As Byte
    QuitChat As Byte
    SendChat As Byte
    PrepareTrade As Byte
    AcceptTrade As Byte
    DeclineTrade As Byte
    QuitTrade As Byte
    UpdateTradeInventory As Byte
    SwapItems As Byte
    Spells As Byte
    HotScript As Byte
    ScriptTile As Byte
    SpellCast As Byte
    Refresh As Byte
    BuySprite As Byte
    ClearOwner As Byte
    RequestEditHouse As Byte
    BuyHouse As Byte
    CheckCommands As Byte
    RequestEditArrow As Byte
    EditArrow As Byte
    SaveArrow As Byte
    CheckArrows As Byte
    RequestEditEmoticon As Byte
    RequestEditElement As Byte
    RequestEditQuest As Byte
    EditEmoticon As Byte
    EditElement As Byte
    SaveEmoticon As Byte
    SaveElement As Byte
    CheckEmoticons As Byte
    MapReport As Byte
    GMTime As Byte
    Weather As Byte
    WarpTo As Byte
    LocalWarp As Byte
    ArrowHit As Byte
    BankDeposit As Byte
    BankWithdraw As Byte
    ReloadScripts As Byte
    CustomMenuClick As Byte
    CustomBoxReturnMessage As Byte

    ' New Party Packets
    PartyCreate As Byte
    PartyDisband As Byte
    PartyInvite As Byte
    PartyInviteAccept As Byte
    PartyInviteDecline As Byte
    PartyLeave As Byte
    PartyChangeLeader As Byte
End Type

Private Type ClientPackIn
    MaxInfo As Byte
    ClearParty As Byte
    NPCHP As Byte
    AlertMessage As Byte
    PlainMessage As Byte
    CharacterList As Byte
    LoginOK As Byte
    News As Byte
    NewCharClasses As Byte
    ClassData As Byte
    GameClock As Byte
    InGame As Byte
    PlayerInventory As Byte
    PlayerInventoryUpdate As Byte
    PlayerBank As Byte
    PlayerBankUpdate As Byte
    OpenBank As Byte
    BankMessage As Byte
    PlayerWornEQ As Byte
    PlayerPoints As Byte
    CustomSprite As Byte
    PlayerHP As Byte
    PlayerEXP As Byte
    PlayerMP As Byte
    SpeechBubble As Byte
    ScriptBubble As Byte
    PlayerSP As Byte
    PlayerStats As Byte
    PlayerData As Byte
    MapLeave As Byte
    GameLeave As Byte
    PlayerLevel As Byte
    SpriteUpdate As Byte
    PlayerMove As Byte
    NpcMove As Byte
    PlayerDirection As Byte
    NPCDirection As Byte
    PlayerXY As Byte
    PRemoveMembers As Byte
    PUpdateMembers As Byte
    Attack As Byte
    NPCAttack As Byte
    CheckForMap As Byte
    MapData As Byte
    TileCheck As Byte
    TileCheckAttribute As Byte
    MapItemData As Byte
    MapNPCData As Byte
    MapDone As Byte
    SayMessage As Byte
    BroadcastMessage As Byte
    GlobalMessage As Byte
    PlayerMessage As Byte
    MapMessage As Byte
    AdminMessage As Byte
    SpawnItem As Byte
    ItemEditor As Byte
    UpdateItem As Byte
    EditItem As Byte
    Mouse As Byte
    MapWeather As Byte
    SpawnNPC As Byte
    NPCDead As Byte
    NPCEditor As Byte
    UpdateNPC As Byte
    EditNPC As Byte
    MapKey As Byte
    EditMap As Byte
    ShopEditor As Byte
    UpdateShop As Byte
    EditShop As Byte
    SpellEditor As Byte
    UpdateSpell As Byte
    EditSpell As Byte
    OpenShop As Byte
    Spells As Byte
    Weather As Byte
    NameColor As Byte
    Fog As Byte
    OnlineList As Byte
    DrawPlayerDamage As Byte
    DrawNPCDamage As Byte
    PrepareTrade As Byte
    QuitTrade As Byte
    TimeEnabled As Byte
    UpdateTradeItem As Byte
    Trading As Byte
    PrepareChat As Byte
    QuitChat As Byte
    SendChat As Byte
    Sound As Byte
    SpriteChange As Byte
    HouseBuy As Byte
    ChangeDirection As Byte
    FlashEvent As Byte
    EmoticonEditor As Byte
    ElementEditor As Byte
    EditElement As Byte
    UpdateElement As Byte
    EditEmoticon As Byte
    UpdateEmoticon As Byte
    ArrowEditor As Byte
    UpdateArrow As Byte
    EditArrow As Byte
    HookShot As Byte
    CheckArrows As Byte
    CheckSprite As Byte
    MapReport As Byte
    GameTime As Byte
    SpellAnimation As Byte
    ScriptSpellAnimation As Byte
    CheckEmoticons As Byte
    LevelUp As Byte
    DamageDisplay As Byte
    ItemBreak As Byte
    ItemWorn As Byte
    ScriptTile As Byte
    SetSpeed As Byte
    ShowCustomMenu As Byte
    CloseCustomMenu As Byte
    LoadPictureCustomMenu As Byte
    LoadLabelCustomMenu As Byte
    LoadTextboxCustomMenu As Byte
    LoadInternetWindow As Byte
    ReturnCustomBoxMessage As Byte
End Type

Public PIn As ClientPackIn
Public POut As ClientPackOut

Public Sub TCPInit(Optional ByVal IPAddr As String = vbNullString, Optional ByVal Port As Long = -1)
    ' Define the seperate and enging characters.
    SEP_CHAR = Chr$(0)
    END_CHAR = Chr$(237)

    ' Wipe the player buffer incase we're reloading TCP.
    PlayerBuffer = vbNullString

    ' Load the byte headers.
    Call TCPAssignHeaders

    ' Prepare the socket address.
    If LenB(IPAddr) = 0 Then
        frmMirage.Socket.RemoteHost = ReadINI("IPCONFIG", "IP", App.Path & "\Config.ini")
    Else
        frmMirage.Socket.RemoteHost = IPAddr
    End If

    ' Prepare the socket port.
    If Port = -1 Then
        frmMirage.Socket.RemotePort = CLng(ReadINI("IPCONFIG", "PORT", App.Path & "\Config.ini"))
    Else
        frmMirage.Socket.RemotePort = Port
    End If
End Sub

Public Sub TCPDestroy()
    ' Close the socket.
    frmMirage.Socket.Close
End Sub

Public Sub TCPRestart()
    ' Close the socket.
    frmMirage.Socket.Close

    ' Open the socket.
    frmMirage.Socket.Connect
End Sub

Public Sub TCPAssignHeaders()
    With PIn
        .MaxInfo = 0
        .ClearParty = 1
        .NPCHP = 2
        .AlertMessage = 3
        .PlainMessage = 4
        .CharacterList = 5
        .LoginOK = 6
        .News = 7
        .NewCharClasses = 8
        .ClassData = 9
        .GameClock = 10
        .InGame = 11
        .PlayerInventory = 12
        .PlayerInventoryUpdate = 13
        .PlayerBank = 14
        .PlayerBankUpdate = 15
        .OpenBank = 16
        .BankMessage = 17
        .PlayerWornEQ = 18
        .PlayerPoints = 19
        .CustomSprite = 20
        .PlayerHP = 21
        .PlayerEXP = 22
        .PlayerMP = 23
        .SpeechBubble = 24
        .ScriptBubble = 25
        .PlayerSP = 26
        .PlayerStats = 27
        .PlayerData = 28
        .MapLeave = 29
        .GameLeave = 30
        .PlayerLevel = 31
        .SpriteUpdate = 32
        .PlayerMove = 33
        .NpcMove = 34
        .PlayerDirection = 35
        .NPCDirection = 36
        .PlayerXY = 37
        .PRemoveMembers = 38
        .PUpdateMembers = 39
        .Attack = 40
        .NPCAttack = 41
        .CheckForMap = 42
        .MapData = 43
        .TileCheck = 44
        .TileCheckAttribute = 45
        .MapItemData = 46
        .MapNPCData = 47
        .MapDone = 48
        .SayMessage = 49
        .BroadcastMessage = 50
        .GlobalMessage = 51
        .PlayerMessage = 52
        .MapMessage = 53
        .AdminMessage = 54
        .SpawnItem = 55
        .ItemEditor = 56
        .UpdateItem = 57
        .EditItem = 58
        .Mouse = 59
        .MapWeather = 60
        .SpawnNPC = 61
        .NPCDead = 62
        .NPCEditor = 63
        .UpdateNPC = 64
        .EditNPC = 65
        .MapKey = 66
        .EditMap = 67
        .ShopEditor = 68
        .UpdateShop = 69
        .EditShop = 70
        .SpellEditor = 71
        .UpdateSpell = 72
        .EditSpell = 73
        .OpenShop = 74
        .Spells = 75
        .Weather = 76
        .NameColor = 77
        .Fog = 78
        .OnlineList = 79
        .DrawPlayerDamage = 80
        .DrawNPCDamage = 81
        .PrepareTrade = 82
        .QuitTrade = 83
        .TimeEnabled = 84
        .UpdateTradeItem = 85
        .Trading = 86
        .PrepareChat = 87
        .QuitChat = 88
        .SendChat = 89
        .Sound = 90
        .SpriteChange = 91
        .HouseBuy = 92
        .ChangeDirection = 93
        .FlashEvent = 94
        .EmoticonEditor = 95
        .ElementEditor = 96
        .EditElement = 97
        .UpdateElement = 98
        .EditEmoticon = 99
        .UpdateEmoticon = 100
        .ArrowEditor = 101
        .UpdateArrow = 102
        .EditArrow = 103
        .HookShot = 104
        .CheckArrows = 105
        .CheckSprite = 106
        .MapReport = 107
        .GameTime = 108
        .SpellAnimation = 109
        .ScriptSpellAnimation = 110
        .CheckEmoticons = 111
        .LevelUp = 112
        .DamageDisplay = 113
        .ItemBreak = 114
        .ItemWorn = 115
        .ScriptTile = 116
        .SetSpeed = 117
        .ShowCustomMenu = 118
        .CloseCustomMenu = 119
        .LoadPictureCustomMenu = 120
        .LoadLabelCustomMenu = 121
        .LoadTextboxCustomMenu = 122
        .LoadInternetWindow = 123
        .ReturnCustomBoxMessage = 124
    End With

    With POut
        .GetClasses = 0
        .NewAccount = 1
        .DeleteAccount = 2
        .AccountLogin = 3
        .GiveMeTheMax = 4
        .AddCharacter = 5
        .DeleteCharacter = 6
        .UseCharacter = 7
        .GuildChangeAccess = 8
        .GuildDisown = 9
        .GuildLeave = 10
        .GuildMake = 11
        .GuildMember = 12
        .GuildTrainee = 13
        .SayMessage = 14
        .EmoteMessage = 15
        .BroadcastMessage = 16
        .GlobalMessage = 17
        .AdminMessage = 18
        .PlayerMessage = 19
        .PlayerMove = 20
        .PlayerDirection = 21
        .UseItem = 22
        .PlayerMouseMove = 23
        .Warp = 24
        .EndShot = 25
        .Attack = 26
        .UseStatPoint = 27
        .PlayerSprite = 28
        .GetStats = 29
        .RequestMap = 30
        .WarpMeTo = 31
        .WarpToMe = 32
        .MapData = 33
        .NeedMap = 34
        .MapGetItem = 35
        .MapDropItem = 36
        .MapRespawn = 37
        .KickPlayer = 38
        .BanList = 39
        .BanListDestroy = 40
        .BanPlayer = 41
        .RequestEditMap = 42
        .RequestEditItem = 43
        .EditItem = 44
        .SaveItem = 45
        .EnableDayNight = 46
        .DayNight = 47
        .RequestEditNPC = 48
        .EditNPC = 49
        .SaveNPC = 50
        .RequestEditShop = 51
        .EditShop = 52
        .SaveShop = 53
        .RequestEditSpell = 54
        .EditSpell = 55
        .SaveSpell = 56
        .ForgetSpell = 57
        .SetAccess = 58
        .WhosOnline = 59
        .OnlineList = 60
        .SetMOTD = 61
        .BuyItem = 62
        .SellItem = 63
        .FixItem = 64
        .Search = 65
        .PlayerChat = 66
        .AcceptChat = 67
        .DeclineChat = 68
        .QuitChat = 69
        .SendChat = 70
        .PrepareTrade = 71
        .AcceptTrade = 72
        .DeclineTrade = 73
        .QuitTrade = 74
        .UpdateTradeInventory = 75
        .SwapItems = 76
        .Spells = 77
        .HotScript = 78
        .ScriptTile = 79
        .SpellCast = 80
        .Refresh = 81
        .BuySprite = 82
        .ClearOwner = 83
        .RequestEditHouse = 84
        .BuyHouse = 85
        .CheckCommands = 86
        .RequestEditArrow = 87
        .EditArrow = 88
        .SaveArrow = 89
        .CheckArrows = 90
        .RequestEditEmoticon = 91
        .RequestEditElement = 92
        .RequestEditQuest = 93
        .EditEmoticon = 94
        .EditElement = 95
        .SaveEmoticon = 96
        .SaveElement = 97
        .CheckEmoticons = 98
        .MapReport = 99
        .GMTime = 100
        .Weather = 101
        .WarpTo = 102
        .LocalWarp = 103
        .ArrowHit = 104
        .BankDeposit = 105
        .BankWithdraw = 106
        .ReloadScripts = 107
        .CustomMenuClick = 108
        .CustomBoxReturnMessage = 109

        ' New Party Packets
        .PartyCreate = 110
        .PartyDisband = 111
        .PartyInvite = 112
        .PartyInviteAccept = 113
        .PartyInviteDecline = 114
        .PartyLeave = 115
        .PartyChangeLeader = 116
    End With
End Sub

Public Sub IncomingData(ByVal DataLength As Long)
    Dim Buffer As String
    Dim packet As String
    Dim Start As Long

    ' Get the data from the socket.
    frmMirage.Socket.GetData Buffer, vbString, DataLength

    ' Add the packet to the end of the player buffer.
    PlayerBuffer = PlayerBuffer & Buffer

    ' Get the ending point of the first packet.
    Start = InStr(PlayerBuffer, END_CHAR)

    ' Loop through all of the received packets.
    Do While Start > 0

        ' Get the packet from the player buffer, except the END_CHAR.
        packet = Left$(PlayerBuffer, Start - 1)

        ' Remove the packet from the player buffer.
        PlayerBuffer = Mid$(PlayerBuffer, Start + 1, Len(PlayerBuffer))

        ' Get the ending point for the next packet.
        Start = InStr(PlayerBuffer, END_CHAR)

        ' Send the packet to be handled by the client.
        If LenB(packet) <> 0 Then
            Call HandleData(packet)
        End If
    Loop
End Sub

Public Function ConnectToServer() As Boolean
    ' Check if we're already connected.
    If IsConnected Then
        ConnectToServer = True
        Exit Function
    End If

    ' Restart the socket.
    Call TCPRestart

    ' Check if we're connected.
    If IsConnected Then
        ConnectToServer = True
    End If
End Function

Public Function IsConnected() As Boolean
    ' Check if the socket is connected.
    If frmMirage.Socket.State = sckConnected Then
        IsConnected = True
    End If
End Function

Public Function IsPlaying(ByVal Index As Long) As Boolean
    ' Check if the player is in-game.
    If GetPlayerName(Index) <> vbNullString Then
        IsPlaying = True
    End If
End Function

Public Function IsAlphaNumeric(ByVal TestString As String) As Boolean
    Dim LoopID As Integer
    Dim sChar As String

    ' Loop through the whole string.
    For LoopID = 1 To Len(TestString)

        ' Get the character from the string to process.
        sChar = Mid$(TestString, LoopID, 1)

        ' Check if the character is either A-Z, a-z, or 0-9.
        If Not sChar Like "[0-9A-Za-z]" Then
            Exit Function
        End If
    Next

    IsAlphaNumeric = True
End Function

Public Sub SendData(ByVal Data As String)
    Dim DBytes() As Byte

    ' Convert the data from unicode to bytes.
    DBytes = StrConv(Data, vbFromUnicode)

    ' Check if we're connected.
    If IsConnected Then
        ' Send the data to the server.
        frmMirage.Socket.SendData DBytes
    End If

    ' Let the operating system process the out-going data.
    DoEvents
End Sub

Public Sub SendNewAccount(ByVal Name As String, ByVal Password As String, ByVal Email As String)
    Call SendData(POut.NewAccount & SEP_CHAR & Trim$(Name) & SEP_CHAR & Trim$(Password) & SEP_CHAR & Trim$(Email) & END_CHAR)
End Sub

Public Sub SendDelAccount(ByVal Name As String, ByVal Password As String)
    Call SendData(POut.DeleteAccount & SEP_CHAR & Trim$(Name) & SEP_CHAR & Trim$(Password) & END_CHAR)
End Sub

Public Sub SendLogin(ByVal Name As String, ByVal Password As String)
    Call SendData(POut.AccountLogin & SEP_CHAR & Trim$(Name) & SEP_CHAR & Trim$(Password) & SEP_CHAR & App.Major & SEP_CHAR & App.Minor & SEP_CHAR & App.Revision & SEP_CHAR & SEC_CODE & END_CHAR)
End Sub

Public Sub SendAddChar(ByVal Name As String, ByVal Sex As Long, ByVal ClassNum As Long, ByVal slot As Long, ByVal HeadC As Long, ByVal BodyC As Long, ByVal LegC As Long)
    Call SendData(POut.AddCharacter & SEP_CHAR & Trim$(Name) & SEP_CHAR & Sex & SEP_CHAR & ClassNum & SEP_CHAR & slot & SEP_CHAR & HeadC & SEP_CHAR & BodyC & SEP_CHAR & LegC & END_CHAR)
End Sub

Public Sub SendDelChar(ByVal slot As Long)
    Call SendData(POut.DeleteCharacter & SEP_CHAR & slot & END_CHAR)
End Sub

Public Sub SendGetClasses()
    Call SendData(POut.GetClasses & END_CHAR)
End Sub

Public Sub SendUseChar(ByVal CharSlot As Long)
    Call SendData(POut.UseCharacter & SEP_CHAR & CharSlot & END_CHAR)
End Sub

Public Sub SayMsg(ByVal Text As String)
    Call SendData(POut.SayMessage & SEP_CHAR & Text & END_CHAR)
End Sub

Public Sub GlobalMsg(ByVal Text As String)
    Call SendData(POut.GlobalMessage & SEP_CHAR & Text & END_CHAR)
End Sub

Public Sub BroadcastMsg(ByVal Text As String)
    Call SendData(POut.BroadcastMessage & SEP_CHAR & Text & END_CHAR)
End Sub

Public Sub EmoteMsg(ByVal Text As String)
    Call SendData(POut.EmoteMessage & SEP_CHAR & Text & END_CHAR)
End Sub

Public Sub PlayerMsg(ByVal Text As String, ByVal MsgTo As String)
    Call SendData(POut.PlayerMessage & SEP_CHAR & MsgTo & SEP_CHAR & Text & END_CHAR)
End Sub

Public Sub AdminMsg(ByVal Text As String)
    Call SendData(POut.AdminMessage & SEP_CHAR & Text & END_CHAR)
End Sub

Public Sub SendPlayerMove()
    Call SendData(POut.PlayerMove & SEP_CHAR & GetPlayerDir(MyIndex) & SEP_CHAR & Player(MyIndex).Moving & END_CHAR)
End Sub

Public Sub SendPlayerDir()
    Call SendData(POut.PlayerDirection & SEP_CHAR & GetPlayerDir(MyIndex) & END_CHAR)
End Sub

Public Sub SendPlayerRequestNewMap(ByVal Cancel As Long)
    Call SendData(POut.RequestMap & SEP_CHAR & GetPlayerDir(MyIndex) & SEP_CHAR & Cancel & END_CHAR)
End Sub

Public Sub SendMap()
    Dim packet As String
    Dim X As Byte
    Dim y As Byte

    packet = POut.MapData & SEP_CHAR & GetPlayerMap(MyIndex) & SEP_CHAR & Trim$(Map(GetPlayerMap(MyIndex)).Name) & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Revision & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Moral & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Up & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Down & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Left & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Right & SEP_CHAR & Map(GetPlayerMap(MyIndex)).music & SEP_CHAR & Map(GetPlayerMap(MyIndex)).BootMap & SEP_CHAR & Map(GetPlayerMap(MyIndex)).BootX & SEP_CHAR & Map(GetPlayerMap(MyIndex)).BootY & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Indoors & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Weather & SEP_CHAR

    For y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            With Map(GetPlayerMap(MyIndex)).Tile(X, y)
                packet = packet & (.Ground & SEP_CHAR & .Mask & SEP_CHAR & .Anim & SEP_CHAR & .Mask2 & SEP_CHAR & .M2Anim & SEP_CHAR & .Fringe & SEP_CHAR & .FAnim & SEP_CHAR & .Fringe2 & SEP_CHAR & .F2Anim & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & .String1 & SEP_CHAR & .String2 & SEP_CHAR & .String3 & SEP_CHAR & .light & SEP_CHAR)
                packet = packet & (.GroundSet & SEP_CHAR & .MaskSet & SEP_CHAR & .AnimSet & SEP_CHAR & .Mask2Set & SEP_CHAR & .M2AnimSet & SEP_CHAR & .FringeSet & SEP_CHAR & .FAnimSet & SEP_CHAR & .Fringe2Set & SEP_CHAR & .F2AnimSet & SEP_CHAR)
            End With
        Next X
    Next y

    With Map(GetPlayerMap(MyIndex))
        For X = 1 To MAX_MAP_NPCS
            packet = packet & (.Npc(X) & SEP_CHAR & .SpawnX(X) & SEP_CHAR & .SpawnY(X) & SEP_CHAR)
        Next X
    End With

    packet = packet & Map(GetPlayerMap(MyIndex)).owner & END_CHAR

    Call SendData(packet)
End Sub

Public Sub WarpMeTo(ByVal Name As String)
    Call SendData(POut.WarpMeTo & SEP_CHAR & Name & END_CHAR)
End Sub

Public Sub WarpToMe(ByVal Name As String)
    Call SendData(POut.WarpToMe & SEP_CHAR & Name & END_CHAR)
End Sub

Public Sub WarpTo(ByVal MapNum As Long, ByVal X As Long, ByVal y As Long)
    Call SendData(POut.WarpTo & SEP_CHAR & MapNum & SEP_CHAR & X & SEP_CHAR & y & END_CHAR)
End Sub

Public Sub LocalWarp(ByVal X As Long, ByVal y As Long)
    Call SendData(POut.LocalWarp & SEP_CHAR & X & SEP_CHAR & y & END_CHAR)
End Sub

Public Sub SendSetAccess(ByVal Name As String, ByVal Access As Byte)
    Call SendData(POut.SetAccess & SEP_CHAR & Name & SEP_CHAR & Access & END_CHAR)
End Sub

Public Sub SendKick(ByVal Name As String)
    Call SendData(POut.KickPlayer & SEP_CHAR & Name & END_CHAR)
End Sub

Public Sub SendBan(ByVal Name As String)
    Call SendData(POut.BanPlayer & SEP_CHAR & Name & END_CHAR)
End Sub

Public Sub SendBanList()
    Call SendData(POut.BanList & END_CHAR)
End Sub

Public Sub SendRequestEditItem()
    Call SendData(POut.RequestEditItem & END_CHAR)
End Sub

Public Sub SendSaveItem(ByVal ItemNum As Long)
    Call SendData(POut.SaveItem & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).MagicReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR & Item(ItemNum).AddHP & SEP_CHAR & Item(ItemNum).AddMP & SEP_CHAR & Item(ItemNum).AddSP & SEP_CHAR & Item(ItemNum).AddSTR & SEP_CHAR & Item(ItemNum).AddDEF & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddMAGI & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).desc & SEP_CHAR & Item(ItemNum).AttackSpeed & SEP_CHAR & Item(ItemNum).Price & SEP_CHAR & Item(ItemNum).Stackable & SEP_CHAR & Item(ItemNum).Bound & END_CHAR)
End Sub

Public Sub SendRequestEditEmoticon()
    Call SendData(POut.RequestEditEmoticon & END_CHAR)
End Sub

Public Sub SendRequestEditElement()
    Call SendData(POut.RequestEditElement & END_CHAR)
End Sub

Public Sub SendSaveEmoticon(ByVal EmoNum As Long)
    Call SendData(POut.SaveEmoticon & SEP_CHAR & EmoNum & SEP_CHAR & Trim$(Emoticons(EmoNum).Command) & SEP_CHAR & Emoticons(EmoNum).Pic & END_CHAR)
End Sub

Public Sub SendSaveElement(ByVal ElementNum As Long)
    Call SendData(POut.SaveElement & SEP_CHAR & ElementNum & SEP_CHAR & Trim$(Element(ElementNum).Name) & SEP_CHAR & Element(ElementNum).Strong & SEP_CHAR & Element(ElementNum).Weak & END_CHAR)
End Sub

Public Sub SendRequestEditArrow()
    Call SendData(POut.RequestEditArrow & END_CHAR)
End Sub

Public Sub SendSaveArrow(ByVal ArrowNum As Long)
    Call SendData(POut.SaveArrow & SEP_CHAR & ArrowNum & SEP_CHAR & Trim$(Arrows(ArrowNum).Name) & SEP_CHAR & Arrows(ArrowNum).Pic & SEP_CHAR & Arrows(ArrowNum).Range & SEP_CHAR & Arrows(ArrowNum).Amount & END_CHAR)
End Sub

Public Sub SendRequestEditNPC()
    Call SendData(POut.RequestEditNPC & END_CHAR)
End Sub

Public Sub SendSaveNPC(ByVal NpcNum As Long)
    Dim packet As String
    Dim i As Long

    packet = POut.SaveNPC & SEP_CHAR & NpcNum & SEP_CHAR & Trim$(Npc(NpcNum).Name) & SEP_CHAR & Trim$(Npc(NpcNum).AttackSay) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).SpawnSecs & SEP_CHAR & Npc(NpcNum).Behavior & SEP_CHAR & Npc(NpcNum).Range & SEP_CHAR & Npc(NpcNum).STR & SEP_CHAR & Npc(NpcNum).DEF & SEP_CHAR & Npc(NpcNum).Speed & SEP_CHAR & Npc(NpcNum).MAGI & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHP & SEP_CHAR & Npc(NpcNum).Exp & SEP_CHAR & Npc(NpcNum).SpawnTime & SEP_CHAR & Npc(NpcNum).Element & SEP_CHAR & Npc(NpcNum).SpriteSize

    For i = 1 To MAX_NPC_DROPS
        packet = packet & (SEP_CHAR & Npc(NpcNum).ItemNPC(i).chance & SEP_CHAR & Npc(NpcNum).ItemNPC(i).ItemNum & SEP_CHAR & Npc(NpcNum).ItemNPC(i).ItemValue)
    Next i

    packet = packet & END_CHAR

    Call SendData(packet)
End Sub

Public Sub SendMapRespawn()
    Call SendData(POut.MapRespawn & END_CHAR)
End Sub

Public Sub SendUseItem(ByVal InvNum As Long)
    Call SendData(POut.UseItem & SEP_CHAR & InvNum & END_CHAR)
End Sub

Public Sub SendDropItem(ByVal InvNum As Long, ByVal Amount As Long)
    Call SendData(POut.MapDropItem & SEP_CHAR & InvNum & SEP_CHAR & Amount & END_CHAR)
End Sub

Public Sub SendWhosOnline()
    Call SendData(POut.WhosOnline & END_CHAR)
End Sub

Public Sub SendOnlineList()
    Call SendData(POut.OnlineList & END_CHAR)
End Sub

Public Sub SendMOTDChange(ByVal MOTD As String)
    Call SendData(POut.SetMOTD & SEP_CHAR & MOTD & END_CHAR)
End Sub

Public Sub SendRequestEditShop()
    Call SendData(POut.RequestEditShop & END_CHAR)
End Sub

Public Sub SendSaveShop(ByVal shopNum As Long)
    Dim packet As String
    Dim i As Integer

    packet = POut.SaveShop & SEP_CHAR & shopNum & SEP_CHAR & Trim$(Shop(shopNum).Name) & SEP_CHAR & Shop(shopNum).FixesItems & SEP_CHAR & Shop(shopNum).BuysItems & SEP_CHAR & Shop(shopNum).currencyItem & SEP_CHAR

    For i = 1 To MAX_SHOP_ITEMS
        packet = packet & (Shop(shopNum).ShopItem(i).ItemNum & SEP_CHAR & Shop(shopNum).ShopItem(i).Amount & SEP_CHAR & Shop(shopNum).ShopItem(i).Price & SEP_CHAR)
    Next i

    packet = packet & END_CHAR

    Call SendData(packet)
End Sub

Public Sub SendRequestEditSpell()
    Call SendData(POut.RequestEditSpell & END_CHAR)
End Sub

Public Sub SendReloadScripts()
    Call SendData(POut.ReloadScripts & END_CHAR)
End Sub

Public Sub SendSaveSpell(ByVal SpellNum As Long)
    Call SendData(POut.SaveSpell & SEP_CHAR & SpellNum & SEP_CHAR & Trim$(Spell(SpellNum).Name) & SEP_CHAR & Spell(SpellNum).ClassReq & SEP_CHAR & Spell(SpellNum).LevelReq & SEP_CHAR & Spell(SpellNum).Type & SEP_CHAR & Spell(SpellNum).Data1 & SEP_CHAR & Spell(SpellNum).Data2 & SEP_CHAR & Spell(SpellNum).Data3 & SEP_CHAR & Spell(SpellNum).MPCost & SEP_CHAR & Trim$(Spell(SpellNum).Sound) & SEP_CHAR & Spell(SpellNum).Range & SEP_CHAR & Spell(SpellNum).SpellAnim & SEP_CHAR & Spell(SpellNum).SpellTime & SEP_CHAR & Spell(SpellNum).SpellDone & SEP_CHAR & Spell(SpellNum).AE & SEP_CHAR & Spell(SpellNum).Big & SEP_CHAR & Spell(SpellNum).Element & END_CHAR)
End Sub

Public Sub SendRequestEditMap()
    Call SendData(POut.RequestEditMap & END_CHAR)
End Sub

Public Sub SendTradeRequest(ByVal Name As String)
    Call SendData(POut.PrepareTrade & SEP_CHAR & Name & END_CHAR)
End Sub

Public Sub SendAcceptTrade()
    Call SendData(POut.AcceptTrade & END_CHAR)
End Sub

Public Sub SendDeclineTrade()
    Call SendData(POut.DeclineTrade & END_CHAR)
End Sub

Public Sub SendBanDestroy()
    Call SendData(POut.BanListDestroy & END_CHAR)
End Sub

Public Sub SendSetPlayerSprite(ByVal Name As String, ByVal SpriteNum As Long)
    Call SendData(POut.PlayerSprite & SEP_CHAR & Name & SEP_CHAR & CStr(SpriteNum) & END_CHAR)
End Sub

Public Sub SendHotScript(ByVal Value As Byte)
    Call SendData(POut.HotScript & SEP_CHAR & Value & END_CHAR)
End Sub

Public Sub SendScriptTile(ByVal Text As String)
    Call SendData(POut.ScriptTile & SEP_CHAR & Text & END_CHAR)
End Sub

Public Sub SendPlayerMoveMouse()
    Call SendData(POut.PlayerMouseMove & SEP_CHAR & GetPlayerDir(MyIndex) & END_CHAR)
End Sub

' Yes, SendChangeDir() uses the warp packet to change directions. :( [Mellowz]
Public Sub SendChangeDir()
    Call SendData(POut.Warp & SEP_CHAR & GetPlayerDir(MyIndex) & END_CHAR)
End Sub

Public Sub SendUseStatPoint(ByVal Value As Byte)
    Call SendData(POut.UseStatPoint & SEP_CHAR & Value & END_CHAR)
End Sub

Public Sub SendGuildLeave()
    Call SendData(POut.GuildLeave & END_CHAR)
End Sub

Public Sub SendGuildMember(ByVal Name As String)
    Call SendData(POut.GuildMember & SEP_CHAR & Name & END_CHAR)
End Sub

Public Sub SendRequestSpells()
    Call SendData(POut.Spells & END_CHAR)
End Sub

Public Sub SendForgetSpell(ByVal SpellID As Long)
    If Player(MyIndex).Spell(SpellID) > 0 Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            If MsgBox("Are you sure you want to forget this spell?", vbYesNo, "Forget Spell") = vbYes Then
                Call SendData(POut.ForgetSpell & SEP_CHAR & SpellID & END_CHAR)
                frmMirage.picPlayerSpells.Visible = False
            End If
        End If
    Else
        Call AddText("There is no spell here.", BRIGHTRED)
    End If
End Sub

Public Sub SendRequestMyStats()
    Call SendData(POut.GetStats & SEP_CHAR & GetPlayerName(MyIndex) & END_CHAR)
End Sub

Public Sub SendSetTrainee(ByVal Name As String)
    Call SendData(POut.GuildTrainee & SEP_CHAR & Name & END_CHAR)
End Sub

Public Sub SendGuildDisown(ByVal Name As String)
    Call SendData(POut.GuildDisown & SEP_CHAR & Name & END_CHAR)
End Sub

Public Sub SendChangeGuildAccess(ByVal Name As String, ByVal AccessLvl As Long)
    Call SendData(POut.GuildChangeAccess & SEP_CHAR & Name & SEP_CHAR & AccessLvl & END_CHAR)
End Sub

Public Sub SendPlayerChat(ByVal Name As String)
    Call SendData(POut.PlayerChat & SEP_CHAR & Name & END_CHAR)
End Sub
