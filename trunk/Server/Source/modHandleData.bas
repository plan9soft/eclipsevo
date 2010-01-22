Attribute VB_Name = "modHandleData"
Option Explicit

Public Sub HandleData(ByVal Index As Long, ByVal Data As String)
    Dim Parse() As String

    On Error Resume Next

    Parse = Split(Data, SEP_CHAR)

    Select Case CByte(Parse(0))
        Case PIn.GetClasses '"getclasses"
            Call Packet_GetClasses(Index)

        Case PIn.NewAccount '"newaccount"
            Call Packet_NewAccount(Index, Parse(1), Parse(2), Parse(3))

        Case PIn.DeleteAccount '"delaccount"
            Call Packet_DeleteAccount(Index, Parse(1), Parse(2))

        Case PIn.AccountLogin '"acclogin"
            Call Packet_AccountLogin(Index, Parse(1), Parse(2), CLng(Parse(3)), CLng(Parse(4)), CLng(Parse(5)), Parse(6))

        Case PIn.GiveMeTheMax '"givemethemax"
            Call Packet_GiveMeTheMax(Index)

        Case PIn.AddCharacter '"addchar"
            Call Packet_AddCharacter(Index, Parse(1), CLng(Parse(2)), CLng(Parse(3)), CLng(Parse(4)), CLng(Parse(5)), CLng(Parse(6)), CLng(Parse(7)))

        Case PIn.DeleteCharacter '"delchar"
            Call Packet_DeleteCharacter(Index, CLng(Parse(1)))

        Case PIn.UseCharacter '"usechar"
            Call Packet_UseCharacter(Index, CLng(Parse(1)))

        Case PIn.GuildChangeAccess '"guildchangeaccess"
            Call Packet_GuildChangeAccess(Index, Parse(1), CLng(Parse(2)))

        Case PIn.GuildDisown '"guilddisown"
            Call Packet_GuildDisown(Index, Parse(1))

        Case PIn.GuildLeave '"guildleave"
            Call Packet_GuildLeave(Index)

        Case PIn.GuildMake '"guildmake"
            Call Packet_GuildMake(Index, Parse(1), Parse(2))

        Case PIn.GuildMember '"guildmember"
            Call Packet_GuildMember(Index, Parse(1))

        Case PIn.GuildTrainee '"guildtrainee"
            Call Packet_GuildTrainee(Index, Parse(1))

        Case PIn.SayMessage '"saymsg"
            Call Packet_SayMessage(Index, Parse(1))

        Case PIn.EmoteMessage '"emotemsg"
            Call Packet_EmoteMessage(Index, Parse(1))

        Case PIn.BroadcastMessage '"broadcastmsg"
            Call Packet_BroadcastMessage(Index, Parse(1))

        Case PIn.GlobalMessage '"globalmsg"
            Call Packet_GlobalMessage(Index, Parse(1))

        Case PIn.AdminMessage '"adminmsg"
            Call Packet_AdminMessage(Index, Parse(1))

        Case PIn.PlayerMessage '"playermsg"
            Call Packet_PlayerMessage(Index, Parse(1), Parse(2))

        Case PIn.PlayerMove '"playermove"
            Call Packet_PlayerMove(Index, CLng(Parse(1)), CLng(Parse(2)))

        Case PIn.PlayerDirection '"playerdir"
            Call Packet_PlayerDirection(Index, CLng(Parse(1)))

        Case PIn.UseItem '"useitem"
            Call Packet_UseItem(Index, CLng(Parse(1)))

        Case PIn.PlayerMouseMove '"playermovemouse"
            Call Packet_PlayerMoveMouse(Index, CLng(Parse(1)))

        Case PIn.Warp '"warp"
            Call Packet_Warp(Index, CLng(Parse(1)))

        Case PIn.EndShot '"endshot"
            Call Packet_EndShot(Index, CLng(Parse(1)))

        Case PIn.Attack '"attack"
            Call Packet_Attack(Index)

        Case PIn.UseStatPoint '"usestatpoint"
            Call Packet_UseStatPoint(Index, CLng(Parse(1)))

        Case PIn.PlayerSprite '"setplayersprite"
            Call Packet_SetPlayerSprite(Index, Parse(1), CLng(Parse(2)))

        Case PIn.GetStats '"getstats"
            Call Packet_GetStats(Index, Parse(1))

        Case PIn.RequestMap '"requestnewmap"
            Call Packet_RequestNewMap(Index, CLng(Parse(1)))

        Case PIn.WarpMeTo '"warpmeto"
            Call Packet_WarpMeTo(Index, Parse(1))

        Case PIn.WarpToMe '"warptome"
            Call Packet_WarpToMe(Index, Parse(1))

        Case PIn.MapData '"mapdata"
            Call Packet_MapData(Index, Parse)

        Case PIn.NeedMap '"needmap"
            Call Packet_NeedMap(Index, CLng(Parse(1)))

        Case PIn.MapGetItem '"mapgetitem"
            Call Packet_MapGetItem(Index)

        Case PIn.MapDropItem '"mapdropitem"
            Call Packet_MapDropItem(Index, CLng(Parse(1)), CLng(Parse(2)))

        Case PIn.MapRespawn '"maprespawn"
            Call Packet_MapRespawn(Index)

        Case PIn.KickPlayer '"kickplayer"
            Call Packet_KickPlayer(Index, Parse(1))

        Case PIn.BanList '"banlist"
            Call Packet_BanList(Index)

        Case PIn.BanListDestroy '"bandestroy"
            Call Packet_BanListDestroy(Index)

        Case PIn.BanPlayer '"banplayer"
            Call Packet_BanPlayer(Index, Parse(1))

        Case PIn.RequestEditMap '"requesteditmap"
            Call Packet_RequestEditMap(Index)

        Case PIn.RequestEditItem '"requestedititem"
            Call Packet_RequestEditItem(Index)

        Case PIn.EditItem '"edititem"
            Call Packet_EditItem(Index, CLng(Parse(1)))

        Case PIn.SaveItem '"saveitem"
            Call Packet_SaveItem(Index, Parse)

        Case PIn.EnableDayNight '"enabledaynight"
            Call Packet_EnableDayNight(Index)

        Case PIn.DayNight '"daynight"
            Call Packet_DayNight(Index)

        Case PIn.RequestEditNPC '"requesteditnpc"
            Call Packet_RequestEditNPC(Index)

        Case PIn.EditNPC '"editnpc"
            Call Packet_EditNPC(Index, CLng(Parse(1)))

        Case PIn.SaveNPC '"savenpc"
            Call Packet_SaveNPC(Index, Parse)

        Case PIn.RequestEditShop '"requesteditshop"
            Call Packet_RequestEditShop(Index)

        Case PIn.EditShop '"editshop"
            Call Packet_EditShop(Index, CLng(Parse(1)))

        Case PIn.SaveShop '"saveshop"
            Call Packet_SaveShop(Index, Parse)

        Case PIn.RequestEditSpell '"requesteditspell"
            Call Packet_RequestEditSpell(Index)

        Case PIn.EditSpell '"editspell"
            Call Packet_EditSpell(Index, CLng(Parse(1)))

        Case PIn.SaveSpell '"savespell"
            Call Packet_SaveSpell(Index, Parse)

        Case PIn.ForgetSpell '"forgetspell"
            Call Packet_ForgetSpell(Index, CLng(Parse(1)))

        Case PIn.SetAccess '"setaccess"
            Call Packet_SetAccess(Index, Parse(1), CLng(Parse(2)))

        Case PIn.WhosOnline '"whosonline"
            Call Packet_WhoIsOnline(Index)

        Case PIn.OnlineList '"onlinelist"
            Call Packet_OnlineList(Index)

        Case PIn.SetMOTD '"setmotd"
            Call Packet_SetMOTD(Index, Parse(1))

        Case PIn.BuyItem '"buy"
            Call Packet_BuyItem(Index, CLng(Parse(1)), CLng(Parse(2)))

        Case PIn.SellItem '"sellitem"
            Call Packet_SellItem(Index, CLng(Parse(1)), CLng(Parse(2)), CLng(Parse(3)), CLng(Parse(4)))

        Case PIn.FixItem '"fixitem"
            Call Packet_FixItem(Index, CLng(Parse(1)), CLng(Parse(2)))

        Case PIn.Search '"search"
            Call Packet_Search(Index, CLng(Parse(1)), CLng(Parse(2)))

        Case PIn.PlayerChat '"playerchat"
            Call Packet_PlayerChat(Index, Parse(1))

        Case PIn.AcceptChat '"achat"
            Call Packet_AcceptChat(Index)

        Case PIn.DeclineChat '"dchat"
            Call Packet_DenyChat(Index)

        Case PIn.QuitChat '"qchat"
            Call Packet_QuitChat(Index)

        Case PIn.SendChat '"sendchat"
            Call Packet_SendChat(Index, Parse(1))

        Case PIn.PrepareTrade '"pptrade"
            Call Packet_PrepareTrade(Index, Parse(1))

        Case PIn.AcceptTrade '"atrade"
            Call Packet_AcceptTrade(Index)

        Case PIn.QuitTrade '"qtrade"
            Call Packet_QuitTrade(Index)

        Case PIn.DeclineTrade '"dtrade"
            Call Packet_DenyTrade(Index)

        Case PIn.UpdateTradeInventory '"updatetradeinv"
            Call Packet_UpdateTradeInventory(Index, CLng(Parse(1)), CLng(Parse(2)), Parse(3), CLng(Parse(4)))

        Case PIn.SwapItems '"swapitems"
            Call Packet_SwapItems(Index)

        Case PIn.Spells '"spells"
            Call Packet_Spells(Index)

        Case PIn.HotScript '"hotscript"
            Call Packet_HotScript(Index, CLng(Parse(1)))

        Case PIn.ScriptTile '"scripttile"
            Call Packet_ScriptTile(Index, CLng(Parse(1)))

        Case PIn.SpellCast '"cast"
            Call Packet_Cast(Index, CLng(Parse(1)))

        Case PIn.Refresh '"refresh"
            Call Packet_Refresh(Index)

        Case PIn.BuySprite '"buysprite"
            Call Packet_BuySprite(Index)

        Case PIn.CheckCommands '"checkcommands"
            Call Packet_CheckCommands(Index, Parse(1))

        Case PIn.RequestEditArrow '"requesteditarrow"
            Call Packet_RequestEditArrow(Index)

        Case PIn.EditArrow '"editarrow"
            Call Packet_EditArrow(Index, CLng(Parse(1)))

        Case PIn.SaveArrow '"savearrow"
            Call Packet_SaveArrow(Index, CLng(Parse(1)), Parse(2), CLng(Parse(3)), CLng(Parse(4)), CLng(Parse(5)))

        Case PIn.CheckArrows '"checkarrows"
            Call Packet_CheckArrows(Index, CLng(Parse(1)))

        Case PIn.RequestEditEmoticon '"requesteditemoticon"
            Call Packet_RequestEditEmoticon(Index)

        Case PIn.RequestEditElement '"requesteditelement"
            Call Packet_RequestEditElement(Index)

        Case PIn.EditEmoticon '"editemoticon"
            Call Packet_EditEmoticon(Index, CLng(Parse(1)))

        Case PIn.EditElement '"editelement"
            Call Packet_EditElement(Index, CLng(Parse(1)))

        Case PIn.SaveEmoticon '"saveemoticon"
            Call Packet_SaveEmoticon(Index, CLng(Parse(1)), Parse(2), CLng(Parse(3)))

        Case PIn.SaveElement '"saveelement"
            Call Packet_SaveElement(Index, CLng(Parse(1)), Parse(2), CLng(Parse(3)), CLng(Parse(4)))

        Case PIn.CheckEmoticons '"checkemoticons"
            Call Packet_CheckEmoticon(Index, CLng(Parse(1)))

        Case PIn.MapReport '"mapreport"
            Call Packet_MapReport(Index)

        Case PIn.GMTime '"gmtime"
            Call Packet_GMTime(Index, CLng(Parse(1)))

        Case PIn.Weather '"weather"
            Call Packet_Weather(Index, CLng(Parse(1)))

        Case PIn.WarpTo '"warpto"
            Call Packet_WarpTo(Index, CLng(Parse(1)), CLng(Parse(2)), CLng(Parse(3)))

        Case PIn.LocalWarp '"localwarp"
            Call Packet_LocalWarp(Index, CLng(Parse(1)), CLng(Parse(2)))

        Case PIn.ArrowHit '"arrowhit"
            Call Packet_ArrowHit(Index, CLng(Parse(1)), CLng(Parse(2)), CLng(Parse(3)), CLng(Parse(4)))

        Case PIn.BankDeposit '"bankdeposit"
            Call Packet_BankDeposit(Index, CLng(Parse(1)), CLng(Parse(2)))

        Case PIn.BankWithdraw '"bankwithdraw"
            Call Packet_BankWithdraw(Index, CLng(Parse(1)), CLng(Parse(2)))

        Case PIn.ReloadScripts '"reloadscripts"
            Call Packet_ReloadScripts(Index)

        ' There MIGHT be an exploit here. Check the development forums for the description.
        Case PIn.CustomMenuClick '"custommenuclick"
            Call Packet_CustomMenuClick(Index, CLng(Parse(1)), CLng(Parse(2)), Parse(3), CLng(Parse(4)), Parse(5))

        Case PIn.CustomBoxReturnMessage '"returningcustomboxmsg"
            Call Packet_CustomBoxReturnMsg(Index, CLng(Parse(1)))
        
        Case PIn.PartyCreate
            Call Packet_PartyCreate(Index, Parse(1))

        Case PIn.PartyDisband
            Call Packet_PartyDisband(Index)

        Case PIn.PartyInvite
            Call Packet_PartyInvite(Index, Parse(1))

        Case PIn.PartyInviteAccept
            Call Packet_PartyInviteAccept(Index)

        Case PIn.PartyInviteDecline
            Call Packet_PartyInviteDecline(Index)

        Case PIn.PartyLeave
            Call Packet_PartyLeave(Index)

        Case PIn.PartyChangeLeader
            Call Packet_PartyChangeLeader(Index, Parse(1))

        Case Else
            Call TextAdd(frmServer.txtText(7), "[Packet Error] Received invalid packet: " & Parse(0) & ".", True)
            Call HackingAttempt(Index, "Received invalid packet: " & Parse(0))
    End Select
End Sub

Public Sub Packet_GetClasses(ByVal Index As Long)
    ' Check if the player is in-game.
    If Player(Index).InGame Then Exit Sub

    Call SendNewCharClasses(Index)
End Sub

Public Sub Packet_NewAccount(ByVal Index As Long, ByVal Username As String, ByVal Password As String, ByVal Email As String)
    If Not IsLoggedIn(Index) Then
        If LenB(Username) < 6 Then
            Call PlainMsg(Index, "Your username must be at least three characters in length.", 1)
            Exit Sub
        End If

        If LenB(Password) < 6 Then
            Call PlainMsg(Index, "Your password must be at least three characters in length.", 1)
            Exit Sub
        End If

        If EMAIL_AUTH = 1 Then
            If LenB(Email) = 0 Then
                Call PlainMsg(Index, "Your email address cannot be blank.", 1)
                Exit Sub
            End If
        End If

        If Not IsAlphaNumeric(Username) Then
            Call PlainMsg(Index, "Your username must consist of alpha-numeric characters!", 1)
            Exit Sub
        End If

        If Not IsAlphaNumeric(Password) Then
            Call PlainMsg(Index, "Your password must consist of alpha-numeric characters!", 1)
            Exit Sub
        End If

        If Not AccountExists(Username) Then
            Call AddAccount(Index, Username, Password, Email)
            Call PlainMsg(Index, "Your account has been created!", 0)
        Else
            Call PlainMsg(Index, "Sorry, that account name is already taken!", 1)
        End If
    End If
End Sub

Public Sub Packet_DeleteAccount(ByVal Index As Long, ByVal Username As String, ByVal Password As String)
    Dim I As Long
    
    If Not IsLoggedIn(Index) Then
        If Not IsAlphaNumeric(Username) Then
            Call PlainMsg(Index, "Your username must consist of alpha-numeric characters!", 2)
            Exit Sub
        End If

        If Not IsAlphaNumeric(Password) Then
            Call PlainMsg(Index, "Your password must consist of alpha-numeric characters!", 2)
            Exit Sub
        End If

        If Not AccountExists(Username) Then
            Call PlainMsg(Index, "That account name does not exist.", 2)
            Exit Sub
        End If

        If Not PasswordOK(Username, Password) Then
            Call PlainMsg(Index, "You've entered an incorrect password.", 2)
            Exit Sub
        End If
    
        Call LoadPlayer(Index, Username)
        For I = 1 To MAX_CHARS
            If CharExist(Index, I) Then
                Call DeleteName(Player(Index).Char(I).Name)
            End If
        Next I
        Call ClearPlayer(Index)

        ' Remove the users main player profile.
        Call Kill(App.Path & "\Accounts\" & Username & "_Info.ini")
        Call Kill(App.Path & "\Accounts\" & Username & "\*.*")

        ' Delete the users account directory.
        Call RmDir(App.Path & "\Accounts\" & Username & "\")
    
        Call PlainMsg(Index, "Your account has been deleted.", 0)
    End If
End Sub

Public Sub Packet_AccountLogin(ByVal Index As Long, ByVal Username As String, ByVal Password As String, ByVal Major As Long, ByVal Minor As Long, ByVal Revision As Long, ByVal Code As String)
    If Not IsLoggedIn(Index) Then
        If Major < CLIENT_MAJOR Or Minor < CLIENT_MINOR Or Revision < CLIENT_REVISION Then
            Call PlainMsg(Index, "Version out-dated. Please visit '" & WEB_SITE & "' for more information.", 3)
            Exit Sub
        End If

        If Not IsAlphaNumeric(Username) Then
            Call PlainMsg(Index, "Your username must consist of alpha-numeric characters!", 3)
            Exit Sub
        End If

        If Not IsAlphaNumeric(Password) Then
            Call PlainMsg(Index, "Your password must consist of alpha-numeric characters!", 3)
            Exit Sub
        End If

        If Not AccountExists(Username) Then
            Call PlainMsg(Index, "That account name does not exist.", 3)
            Exit Sub
        End If
    
        If Not PasswordOK(Username, Password) Then
            Call PlainMsg(Index, "You've entered an incorrect password.", 3)
            Exit Sub
        End If
    
        If IsMultiAccounts(Username) Then
            Call PlainMsg(Index, "Multiple account logins is not authorized.", 3)
            Exit Sub
        End If
    
        If frmServer.Closed.Value = vbChecked Then
            Call PlainMsg(Index, "The server is closed at the moment!", 3)
            Exit Sub
        End If
    
        If Code <> SEC_CODE Then
            Call AlertMsg(Index, "The client password does not match the server password.")
            Exit Sub
        End If
    
        Call LoadPlayer(Index, Username)
        Call SendChars(Index)
    
        Call TextAdd(frmServer.txtText(0), GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".", True)
    End If
End Sub

Public Sub Packet_GiveMeTheMax(ByVal Index As Long)
    Dim packet As String

    ' Check if the player is in-game.
    If IsLoggedIn(Index) Then Exit Sub

    Call SendNewsTo(Index)

    packet = POut.MaxInfo & SEP_CHAR
    packet = packet & GAME_NAME & SEP_CHAR
    packet = packet & MAX_PLAYERS & SEP_CHAR
    packet = packet & MAX_ITEMS & SEP_CHAR
    packet = packet & MAX_NPCS & SEP_CHAR
    packet = packet & MAX_SHOPS & SEP_CHAR
    packet = packet & MAX_SPELLS & SEP_CHAR
    packet = packet & MAX_MAPS & SEP_CHAR
    packet = packet & MAX_MAP_ITEMS & SEP_CHAR
    packet = packet & MAX_MAPX & SEP_CHAR
    packet = packet & MAX_MAPY & SEP_CHAR
    packet = packet & MAX_EMOTICONS & SEP_CHAR
    packet = packet & MAX_ELEMENTS & SEP_CHAR
    packet = packet & PAPERDOLL & SEP_CHAR
    packet = packet & SPRITESIZE & SEP_CHAR
    packet = packet & MAX_SCRIPTSPELLS & SEP_CHAR
    packet = packet & CUSTOM_SPRITE & SEP_CHAR
    packet = packet & LEVEL & SEP_CHAR
    packet = packet & MAX_PARTY_MEMBERS & SEP_CHAR
    packet = packet & STAT1 & SEP_CHAR
    packet = packet & STAT2 & SEP_CHAR
    packet = packet & STAT3 & SEP_CHAR
    packet = packet & STAT4 & SEP_CHAR
    packet = packet & END_CHAR

    Call SendDataTo(Index, packet)
End Sub

Public Sub Packet_AddCharacter(ByVal Index As Long, ByVal Name As String, ByVal Sex As Long, ByVal Class As Long, ByVal CharNum As Long, ByVal Head As Long, ByVal Body As Long, ByVal Leg As Long)
    If CharNum < 1 Or CharNum > MAX_CHARS Then
        Call HackingAttempt(Index, "Invalid CharNum")
        Exit Sub
    End If

    ' Check if the player is in-game.
    If Player(Index).InGame Then Exit Sub

    If LenB(Name) < 6 Then
        Call HackingAttempt(Index, "Invalid Name Length")
        Exit Sub
    End If
    
    If Sex <> SEX_MALE And Sex <> SEX_FEMALE Then
        Call HackingAttempt(Index, "Invalid Sex")
        Exit Sub
    End If
    
    If Class < 0 Or Class > MAX_CLASSES Then
        Call HackingAttempt(Index, "Invalid Class")
        Exit Sub
    End If

    If ClassData(Class).Locked Then
        Call PlainMsg(Index, "Sorry, but that class cannot be choosen!", 4)
        Exit Sub
    End If

    If Not IsAlphaNumeric(Name) Then
        Call PlainMsg(Index, "Your character name must consist of alpha-numeric characters!", 4)
        Exit Sub
    End If

    If CharExist(Index, CharNum) Then
        Call PlainMsg(Index, "Character already exists!", 4)
        Exit Sub
    End If
    
    If FindChar(Name) Then
        Call PlainMsg(Index, "Sorry, but that character name is in use!", 4)
        Exit Sub
    End If

    Call AddChar(Index, Name, Sex, Class, CharNum, Head, Body, Leg)

    Call SendChars(Index)

    Call PlainMsg(Index, "Character has been created!", 5)

    If SCRIPTING = 1 Then
        Call MyScript.ExecuteStatement("Scripts\Main.txt", "OnNewChar " & Index & "," & CharNum)
    End If
End Sub

Public Sub Packet_DeleteCharacter(ByVal Index As Long, ByVal CharNum As Long)
    If CharNum < 1 Or CharNum > MAX_CHARS Then
        Call HackingAttempt(Index, "Invalid CharNum")
        Exit Sub
    End If

    ' Check if the player is in-game.
    If Player(Index).InGame Then Exit Sub

    If CharExist(Index, CharNum) Then
        Call DelChar(Index, CharNum)
        Call SendChars(Index)
    
        Call PlainMsg(Index, "Character has been deleted!", 5)
    Else
        Call PlainMsg(Index, "Character does not exist!", 5)
    End If
End Sub

Public Sub Packet_UseCharacter(ByVal Index As Long, ByVal CharNum As Long)
    Dim FileID As Integer

    If CharNum < 1 Or CharNum > MAX_CHARS Then
        Call HackingAttempt(Index, "Invalid CharNum")
        Exit Sub
    End If

    ' The character is already in-game.
    If Player(Index).InGame Then Exit Sub

    If CharExist(Index, CharNum) Then
        Player(Index).CharNum = CharNum
    
        If frmServer.GMOnly.Value = vbChecked Then
            If GetPlayerAccess(Index) = 0 Then
                Call PlainMsg(Index, "The server is only available to GMs at the moment!", 5)
                Exit Sub
            End If
        End If
    
        Call JoinGame(Index)

        Call TextAdd(frmServer.txtText(0), GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & GAME_NAME & ".", True)
        Call UpdateTOP
    
        If Not FindChar(GetPlayerName(Index)) Then
            FileID = FreeFile
            Open App.Path & "\Accounts\CharList.txt" For Append As #FileID
                Print #FileID, GetPlayerName(Index)
            Close #FileID
        End If
    Else
        Call PlainMsg(Index, "Character does not exist!", 5)
    End If
End Sub

Public Sub Packet_GuildChangeAccess(ByVal Index As Long, ByVal Name As String, ByVal Rank As Long)
    Dim NameIndex As Long
    
    If LenB(Name) = 0 Then
        Call PlayerMsg(Index, "You must enter a player name to proceed.", WHITE)
        Exit Sub
    End If

    If Rank < 0 Or Rank > 4 Then
        Call PlayerMsg(Index, "You must provide a valid rank to proceed.", RED)
        Exit Sub
    End If

    NameIndex = FindPlayer(Name)

    If NameIndex = 0 Then
        Call PlayerMsg(Index, Name & " is currently offline.", WHITE)
        Exit Sub
    End If

    If LenB(GetPlayerGuild(Index)) = 0 Then
        Call PlayerMsg(Index, "You are not in a guild.", RED)
        Exit Sub
    End If

    If GetPlayerGuild(NameIndex) <> GetPlayerGuild(Index) Then
        Call PlayerMsg(Index, Name & " is not in your guild.", RED)
        Exit Sub
    End If

    If GetPlayerGuildAccess(Index) < 4 Then
        Call PlayerMsg(Index, "You are not the owner of this guild.", RED)
        Exit Sub
    End If

    Call SetPlayerGuildAccess(NameIndex, Rank)
    Call SendPlayerData(NameIndex)
End Sub

Public Sub Packet_GuildDisown(ByVal Index As Long, ByVal Name As String)
    Dim NameIndex As Long

    NameIndex = FindPlayer(Name)

    If NameIndex = 0 Then
        Call PlayerMsg(Index, Name & " is currently offline.", WHITE)
        Exit Sub
    End If

    If LenB(GetPlayerGuild(NameIndex)) = 0 Then
        Call PlayerMsg(Index, Name & " is not in a guild.", RED)
        Exit Sub
    End If

    If GetPlayerGuild(NameIndex) <> GetPlayerGuild(Index) Then
        Call PlayerMsg(Index, Name & " is not in your guild.", RED)
        Exit Sub
    End If

    If GetPlayerGuildAccess(NameIndex) > GetPlayerGuildAccess(Index) Then
        Call PlayerMsg(Index, Name & " has a higher guild level than you.", RED)
        Exit Sub
    End If

    Call SetPlayerGuild(NameIndex, vbNullString)
    Call SetPlayerGuildAccess(NameIndex, 0)
    Call SendPlayerData(NameIndex)
End Sub

Public Sub Packet_GuildLeave(ByVal Index As Long)
Dim X As Integer, J As Integer, GuildName As String

    If LenB(GetPlayerGuild(Index)) = 0 Then
        Call PlayerMsg(Index, "You are not in a guild.", RED)
        Exit Sub
    End If

    GuildName = GetPlayerGuild(Index)
    For X = 1 To MAX_GUILDS
        If GetPlayerName(Index) = Guild(X).Founder Then
            For J = 1 To MAX_PLAYERS
                If GetPlayerGuild(J) = GuildName Then
                    Call PlayerMsg(J, "Your guild has been disbanded.", 1)
                    Call SetPlayerGuild(Index, vbNullString)
                    Call SetPlayerGuildAccess(Index, 0)
                    Call SendPlayerData(Index)
                End If
            Next J
            Guild(X).Founder = vbNullString
            Guild(X).Name = vbNullString
        Else
            Call SetPlayerGuild(Index, vbNullString)
            Call SetPlayerGuildAccess(Index, 0)
            Call SendPlayerData(Index)
        End If
    Next X
End Sub

Public Sub Packet_GuildMake(ByVal Index As Long, ByVal Name As String, ByVal GuildName As String)
    Dim NameIndex As Long, I As Long

    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    If LenB(GuildName) = 0 Or LenB(GuildName) >= NAME_LENGTH Then
        Call PlayerMsg(Index, "This guild name is either to short or to long.", RED)
        Exit Sub
    End If

    NameIndex = FindPlayer(Name)
    If NameIndex = 0 Then
        Call PlayerMsg(Index, Name & " is currently offline.", WHITE)
        Exit Sub
    End If

    If LenB(GetPlayerGuild(NameIndex)) <> 0 Then
        Call PlayerMsg(Index, Name & " is already in a guild.", RED)
        Exit Sub
    End If

    For I = 1 To MAX_GUILDS
        If Trim(Guild(I).Name) = Trim(GuildName) Then
            Call PlayerMsg(Index, "This guild already exists.", RED)
            Exit Sub
        End If
    Next I
    
    For I = 1 To MAX_GUILDS
        If LenB(Guild(I).Name) = 40 Then
            Guild(I).Name = Trim(GuildName)
            Guild(I).Founder = Name
            Call WriteINI(Val(I), "Name", Trim(GuildName), App.Path & "\Guilds.ini")
            Call WriteINI(Val(I), "Founder", Trim(GuildName), App.Path & "\Guilds.ini")
        
            Call SetPlayerGuild(NameIndex, GuildName)
            Call SetPlayerGuildAccess(NameIndex, 4)
            Call SendPlayerData(NameIndex)
            
            Exit Sub
        End If
    Next I
    Exit Sub
End Sub

Public Sub Packet_GuildMember(ByVal Index As Long, ByVal Name As String)
    Dim NameIndex As Long

    NameIndex = FindPlayer(Name)

    If NameIndex = 0 Then
        Call PlayerMsg(Index, Name & " is currently offline.", WHITE)
        Exit Sub
    End If

    If LenB(GetPlayerGuild(NameIndex)) = 0 Then
        Call PlayerMsg(Index, Name & " is not in a guild.", RED)
        Exit Sub
    End If

    If GetPlayerGuild(NameIndex) <> GetPlayerGuild(Index) Then
        Call PlayerMsg(Index, Name & " is not in your guild.", RED)
        Exit Sub
    End If

    If GetPlayerGuildAccess(NameIndex) > 1 Then
        Call PlayerMsg(Index, Name & " has already been admitted.", WHITE)
        Exit Sub
    End If

    Call SetPlayerGuild(NameIndex, GetPlayerGuild(Index))
    Call SetPlayerGuildAccess(NameIndex, 1)
    Call SendPlayerData(NameIndex)
End Sub

Public Sub Packet_GuildTrainee(ByVal Index As Long, ByVal Name As String)
    Dim NameIndex As Long

    NameIndex = FindPlayer(Name)

    If NameIndex = 0 Then
        Call PlayerMsg(Index, Name & " is currently offline.", WHITE)
        Exit Sub
    End If

    If LenB(GetPlayerGuild(NameIndex)) <> 0 Then
        Call PlayerMsg(Index, Name & " is already in a guild.", RED)
        Exit Sub
    End If

    Call SetPlayerGuild(NameIndex, GetPlayerGuild(Index))
    Call SetPlayerGuildAccess(NameIndex, 0)
    Call SendPlayerData(NameIndex)
End Sub

Public Sub Packet_SayMessage(ByVal Index As Long, ByVal Message As String)
    If frmServer.chkLogMap.Value = vbUnchecked Then
        If GetPlayerAccess(Index) = 0 Then
            Call PlayerMsg(Index, "Map messages have been disabled by the server!", BRIGHTRED)
            Exit Sub
        End If
    End If

    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & ": " & Message, SayColor)
    Call MapMsg2(GetPlayerMap(Index), Message, Index)

    Call TextAdd(frmServer.txtText(3), GetPlayerName(Index) & " On Map " & GetPlayerMap(Index) & ": " & Message, True)
    Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " : " & Message, PLAYER_LOG)
End Sub

Public Sub Packet_EmoteMessage(ByVal Index As Long, ByVal Message As String)
    If frmServer.chkLogEmote.Value = vbUnchecked Then
        If GetPlayerAccess(Index) = 0 Then
            Call PlayerMsg(Index, "Emote messages have been disabled by the server!", BRIGHTRED)
            Exit Sub
        End If
    End If

    Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & ": " & Message, EmoteColor)

    Call TextAdd(frmServer.txtText(6), GetPlayerName(Index) & " " & Message, True)
    Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " " & Message, PLAYER_LOG)
End Sub

Public Sub Packet_BroadcastMessage(ByVal Index As Long, ByVal Message As String)
    If frmServer.chkLogBC.Value = vbUnchecked Then
        If GetPlayerAccess(Index) = 0 Then
            Call PlayerMsg(Index, "Broadcast messages have been disabled by the server!", BRIGHTRED)
            Exit Sub
        End If
    End If

    If Player(Index).Mute Then
        Call PlayerMsg(Index, "You are muted. You cannot broadcast messages.", BRIGHTRED)
        Exit Sub
    End If

    Call GlobalMsg(GetPlayerName(Index) & ": " & Message, BroadcastColor)

    Call TextAdd(frmServer.txtText(0), GetPlayerName(Index) & ": " & Message, True)
    Call TextAdd(frmServer.txtText(1), GetPlayerName(Index) & ": " & Message, True)
    Call AddLog(GetPlayerName(Index) & ": " & Message, PLAYER_LOG)
End Sub

Public Sub Packet_GlobalMessage(ByVal Index As Long, ByVal Message As String)
    If frmServer.chkLogGlobal.Value = vbUnchecked Then
        If GetPlayerAccess(Index) = 0 Then
            Call PlayerMsg(Index, "Global messages have been disabled by the server!", BRIGHTRED)
            Exit Sub
        End If
    End If

    If Player(Index).Mute Then
        Call PlayerMsg(Index, "You are muted. You cannot broadcast messages.", BRIGHTRED)
        Exit Sub
    End If

    If GetPlayerAccess(Index) > 0 Then
        Call GlobalMsg("(Global) " & GetPlayerName(Index) & ": " & Message, GlobalColor)

        Call TextAdd(frmServer.txtText(0), "(Global) " & GetPlayerName(Index) & ": " & Message, True)
        Call TextAdd(frmServer.txtText(2), GetPlayerName(Index) & ": " & Message, True)
        Call AddLog("(Global) " & GetPlayerName(Index) & ": " & Message, ADMIN_LOG)
    End If
End Sub

Public Sub Packet_AdminMessage(ByVal Index As Long, ByVal Message As String)
    If frmServer.chkLogAdmin.Value = vbUnchecked Then
        Call PlayerMsg(Index, "Admin messages have been disabled by the server!", BRIGHTRED)
        Exit Sub
    End If

    If GetPlayerAccess(Index) > 0 Then
        Call AdminMsg("(Admin " & GetPlayerName(Index) & ") " & Message, AdminColor)

        Call TextAdd(frmServer.txtText(5), GetPlayerName(Index) & ": " & Message, True)
        Call AddLog("(Admin " & GetPlayerName(Index) & ") " & Message, ADMIN_LOG)
    End If
End Sub

Public Sub Packet_PlayerMessage(ByVal Index As Long, ByVal Name As String, ByVal Message As String)
    Dim MsgTo As Long
    
    If frmServer.chkLogPM.Value = vbUnchecked Then
        If GetPlayerAccess(Index) = 0 Then
            Call PlayerMsg(Index, "Personal messages have been disabled by the server!", BRIGHTRED)
            Exit Sub
        End If
    End If

    If LenB(Name) = 0 Then
        Call PlayerMsg(Index, "You must select a player name to private message.", BRIGHTRED)
        Exit Sub
    End If

    If LenB(Message) = 0 Then
        Call PlayerMsg(Index, "You must send a message to private message another player.", BRIGHTRED)
        Exit Sub
    End If

    MsgTo = FindPlayer(Name)

    If MsgTo = 0 Then
        Call PlayerMsg(Index, Name & " is currently offline.", WHITE)
        Exit Sub
    End If

    Call PlayerMsg(Index, "You tell " & GetPlayerName(MsgTo) & ", '" & Message & "'", TellColor)
    Call PlayerMsg(MsgTo, GetPlayerName(Index) & " tells you, '" & Message & "'", TellColor)

    Call TextAdd(frmServer.txtText(4), "To " & GetPlayerName(MsgTo) & " From " & GetPlayerName(Index) & ": " & Message, True)
    Call AddLog(GetPlayerName(Index) & " tells " & GetPlayerName(MsgTo) & ", " & Message & "'", PLAYER_LOG)
End Sub

Public Sub Packet_PlayerMove(ByVal Index As Long, ByVal Dir As Long, ByVal Movement As Long)
    If Player(Index).GettingMap = YES Then
        Exit Sub
    End If

    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Call HackingAttempt(Index, "Invalid Direction")
        Exit Sub
    End If

    If Movement <> 1 And Movement <> 2 Then
        Call HackingAttempt(Index, "Invalid Movement")
        Exit Sub
    End If

    If Player(Index).CastedSpell = YES Then
        If GetTickCount > Player(Index).AttackTimer + 1000 Then
            Player(Index).CastedSpell = NO
        Else
            Call SendPlayerXY(Index)
            Exit Sub
        End If
    End If

    If Player(Index).Locked = True Then
        Call SendPlayerXY(Index)
        Exit Sub
    End If

    Call PlayerMove(Index, Dir, Movement)
End Sub

Public Sub Packet_PlayerDirection(ByVal Index As Long, ByVal Dir As Long)
    If Player(Index).GettingMap = YES Then
        Exit Sub
    End If

    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Call HackingAttempt(Index, "Invalid Direction")
        Exit Sub
    End If

    Call SetPlayerDir(Index, Dir)

    Call SendDataToMapBut(Index, GetPlayerMap(Index), POut.PlayerDirection & SEP_CHAR & Index & SEP_CHAR & GetPlayerDir(Index) & END_CHAR)
End Sub

Public Sub Packet_UseItem(ByVal Index As Long, ByVal InvNum As Long)
    Dim CharNum As Long
    Dim SpellID As Long
    Dim MinLvl As Long
    Dim X As Long
    Dim Y As Long

    If InvNum < 1 Or InvNum > MAX_ITEMS Then
        Call HackingAttempt(Index, "Invalid InvNum")
        Exit Sub
    End If

    If Player(Index).LockedItems Then
        Call PlayerMsg(Index, "You currently cannot use any items.", BRIGHTRED)
        Exit Sub
    End If

    CharNum = Player(Index).CharNum

    Dim n As Long

    ' Find out what kind of item it is
    Select Case Item(GetPlayerInvItemNum(Index, InvNum)).Type
        Case ITEM_TYPE_ARMOR
            If InvNum <> GetPlayerArmorSlot(Index) Then
                If ItemIsUsable(Index, InvNum) Then
                    Call SetPlayerArmorSlot(Index, InvNum)
                End If
            Else
                Call SetPlayerArmorSlot(Index, 0)
            End If
            Call SendWornEquipment(Index)
    
        Case ITEM_TYPE_WEAPON
            If InvNum <> GetPlayerWeaponSlot(Index) Then
                If ItemIsUsable(Index, InvNum) Then
                    Call SetPlayerWeaponSlot(Index, InvNum)
                End If
            Else
                Call SetPlayerWeaponSlot(Index, 0)
            End If
            Call SendWornEquipment(Index)
    
        Case ITEM_TYPE_TWO_HAND
            If InvNum <> GetPlayerWeaponSlot(Index) Then
                If ItemIsUsable(Index, InvNum) Then
                    If GetPlayerShieldSlot(Index) <> 0 Then
                        Call SetPlayerShieldSlot(Index, 0)
                    End If

                    Call SetPlayerWeaponSlot(Index, InvNum)
                End If
            Else
                Call SetPlayerWeaponSlot(Index, 0)
            End If
            Call SendWornEquipment(Index)
    
        Case ITEM_TYPE_HELMET
            If InvNum <> GetPlayerHelmetSlot(Index) Then
                If ItemIsUsable(Index, InvNum) Then
                    Call SetPlayerHelmetSlot(Index, InvNum)
                End If
            Else
                Call SetPlayerHelmetSlot(Index, 0)
            End If
            Call SendWornEquipment(Index)
    
        Case ITEM_TYPE_SHIELD
            If InvNum <> GetPlayerShieldSlot(Index) Then
                If ItemIsUsable(Index, InvNum) Then
                    If GetPlayerWeaponSlot(Index) <> 0 Then
                        If Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Type = ITEM_TYPE_TWO_HAND Then
                            Call SetPlayerWeaponSlot(Index, 0)
                        End If
                    End If

                    Call SetPlayerShieldSlot(Index, InvNum)
                End If
            Else
                Call SetPlayerShieldSlot(Index, 0)
            End If
            Call SendWornEquipment(Index)
    
        Case ITEM_TYPE_LEGS
            If InvNum <> GetPlayerLegsSlot(Index) Then
                If ItemIsUsable(Index, InvNum) Then
                    Call SetPlayerLegsSlot(Index, InvNum)
                End If
            Else
                Call SetPlayerLegsSlot(Index, 0)
            End If
            Call SendWornEquipment(Index)
    
        Case ITEM_TYPE_RING
            If InvNum <> GetPlayerRingSlot(Index) Then
                If ItemIsUsable(Index, InvNum) Then
                    Call SetPlayerRingSlot(Index, InvNum)
                End If
            Else
                Call SetPlayerRingSlot(Index, 0)
            End If
            Call SendWornEquipment(Index)
    
        Case ITEM_TYPE_NECKLACE
            If InvNum <> GetPlayerNecklaceSlot(Index) Then
                If ItemIsUsable(Index, InvNum) Then
                    Call SetPlayerNecklaceSlot(Index, InvNum)
                End If
            Else
                Call SetPlayerNecklaceSlot(Index, 0)
            End If
            Call SendWornEquipment(Index)

        Case ITEM_TYPE_POTIONADDHP
            Call SetPlayerHP(Index, GetPlayerHP(Index) + Item(Player(Index).Char(CharNum).Inv(InvNum).num).Data1)
            If Item(GetPlayerInvItemNum(Index, InvNum)).Stackable = 1 Then
                Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 1)
            Else
                Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 0)
            End If
    
        Case ITEM_TYPE_POTIONADDMP
            Call SetPlayerMP(Index, GetPlayerMP(Index) + Item(Player(Index).Char(CharNum).Inv(InvNum).num).Data1)
            If Item(GetPlayerInvItemNum(Index, InvNum)).Stackable = 1 Then
                Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 1)
            Else
                Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 0)
            End If
    
        Case ITEM_TYPE_POTIONADDSP
            Call SetPlayerSP(Index, GetPlayerSP(Index) + Item(Player(Index).Char(CharNum).Inv(InvNum).num).Data1)
            If Item(GetPlayerInvItemNum(Index, InvNum)).Stackable = 1 Then
                Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 1)
            Else
                Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 0)
            End If
    
        Case ITEM_TYPE_POTIONSUBHP
            Call SetPlayerHP(Index, GetPlayerHP(Index) - Item(Player(Index).Char(CharNum).Inv(InvNum).num).Data1)
            If Item(GetPlayerInvItemNum(Index, InvNum)).Stackable = 1 Then
                Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 1)
            Else
                Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 0)
            End If
    
        Case ITEM_TYPE_POTIONSUBMP
            Call SetPlayerMP(Index, GetPlayerMP(Index) - Item(Player(Index).Char(CharNum).Inv(InvNum).num).Data1)
            If Item(GetPlayerInvItemNum(Index, InvNum)).Stackable = 1 Then
                Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 1)
            Else
                Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 0)
            End If
    
        Case ITEM_TYPE_POTIONSUBSP
            Call SetPlayerSP(Index, GetPlayerSP(Index) - Item(Player(Index).Char(CharNum).Inv(InvNum).num).Data1)
            If Item(GetPlayerInvItemNum(Index, InvNum)).Stackable = 1 Then
                Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 1)
            Else
                Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).num, 0)
            End If
    
        Case ITEM_TYPE_KEY
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    If GetPlayerY(Index) > 0 Then
                        X = GetPlayerX(Index)
                        Y = GetPlayerY(Index) - 1
                    Else
                        Exit Sub
                    End If
    
                Case DIR_DOWN
                    If GetPlayerY(Index) < MAX_MAPY Then
                        X = GetPlayerX(Index)
                        Y = GetPlayerY(Index) + 1
                    Else
                        Exit Sub
                    End If
    
                Case DIR_LEFT
                    If GetPlayerX(Index) > 0 Then
                        X = GetPlayerX(Index) - 1
                        Y = GetPlayerY(Index)
                    Else
                        Exit Sub
                    End If
    
                Case DIR_RIGHT
                    If GetPlayerX(Index) < MAX_MAPX Then
                        X = GetPlayerX(Index) + 1
                        Y = GetPlayerY(Index)
                    Else
                        Exit Sub
                    End If
            End Select
    
            ' Check if a key exists.
            If Map(GetPlayerMap(Index)).Tile(X, Y).Type = TILE_TYPE_KEY Then
                ' Check if the key they are using matches the map key.
                If GetPlayerInvItemNum(Index, InvNum) = Map(GetPlayerMap(Index)).Tile(X, Y).Data1 Then
                    TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES
                    TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
    
                    Call SendDataToMap(GetPlayerMap(Index), POut.MapKey & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & END_CHAR)

                    If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1) = vbNullString Then
                        Call MapMsg(GetPlayerMap(Index), "A door has been unlocked!", WHITE)
                    Else
                        Call MapMsg(GetPlayerMap(Index), Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1), WHITE)
                    End If

                    Call SendDataToMap(GetPlayerMap(Index), POut.Sound & SEP_CHAR & "key" & END_CHAR)
    
                    ' Check if we are supposed to take away the item.
                    If Map(GetPlayerMap(Index)).Tile(X, Y).Data2 = 1 Then
                        Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), 0)
                        Call PlayerMsg(Index, "The key disolves.", YELLOW)
                    End If
                End If
            End If
    
            If Map(GetPlayerMap(Index)).Tile(X, Y).Type = TILE_TYPE_DOOR Then
                TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES
                TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
    
                Call SendDataToMap(GetPlayerMap(Index), POut.MapKey & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & END_CHAR)
                Call SendDataToMap(GetPlayerMap(Index), POut.Sound & SEP_CHAR & "key" & END_CHAR)
            End If
    
        Case ITEM_TYPE_SPELL
            SpellID = Item(GetPlayerInvItemNum(Index, InvNum)).Data1
    
            If SpellID > 0 Then
                If Spell(SpellID).ClassReq - 1 = GetPlayerClass(Index) Or Spell(SpellID).ClassReq = 0 Then
                    If Spell(SpellID).LevelReq = 0 And Player(Index).Char(Player(Index).CharNum).Access < 1 Then
                        Call PlayerMsg(Index, "This spell can only be used by admins!", BRIGHTRED)
                        Exit Sub
                    End If

                    MinLvl = GetSpellReqLevel(SpellID)

                    If MinLvl <= GetPlayerLevel(Index) Then
                        MinLvl = FindOpenSpellSlot(Index)
    
                        If MinLvl > 0 Then
                            If Not HasSpell(Index, SpellID) Then
                                Call SetPlayerSpell(Index, MinLvl, SpellID)
                                Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), 0)
                                Call PlayerMsg(Index, "You have learned a new spell!", WHITE)
                            Else
                                Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), 0)
                                Call PlayerMsg(Index, "You have already learned this spell!  The spells crumbles into dust.", BRIGHTRED)
                            End If
                        Else
                            Call PlayerMsg(Index, "You have learned all that you can learn!", BRIGHTRED)
                        End If
                    Else
                        Call PlayerMsg(Index, "You must be level " & MinLvl & " to learn this spell.", WHITE)
                    End If
                Else
                    Call PlayerMsg(Index, "This spell can only be learned by a " & GetClassName(Spell(SpellID).ClassReq - 1) & ".", WHITE)
                End If
            End If
    
        Case ITEM_TYPE_SCRIPTED
            If SCRIPTING = 1 Then
                MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedItem " & Index & "," & Item(Player(Index).Char(CharNum).Inv(InvNum).num).Data1
            End If
    End Select
    
    Call SendStats(Index)
    Call SendHP(Index)
    Call SendMP(Index)
    Call SendSP(Index)

    Call SendIndexWornEquipment(Index)
End Sub

' This packet seems to me like it's incomplete. [Mellowz]
Public Sub Packet_PlayerMoveMouse(ByVal Index As Long, ByVal Dir As Long)
    If Player(Index).GettingMap = YES Then
        Exit Sub
    End If

    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Call HackingAttempt(Index, "Invalid Direction")
        Exit Sub
    End If

    If Player(Index).Locked = True Then
        Call SendPlayerXY(Index)
        Exit Sub
    End If

    If Player(Index).CastedSpell = YES Then
        If GetTickCount > Player(Index).AttackTimer + 1000 Then
            Player(Index).CastedSpell = NO
        Else
            Call SendPlayerXY(Index)
            Exit Sub
        End If
    End If

    If CLng(ReadINI("CONFIG", "mouse", FILE_DATAINI, "0")) = 1 Then
        Call SendDataTo(Index, POut.Mouse & END_CHAR)
    End If
End Sub

Public Sub Packet_Warp(ByVal Index As Long, ByVal Dir As Long)
    Select Case Dir
        Case DIR_UP
            If Map(GetPlayerMap(Index)).Up > 0 Then
                If GetPlayerY(Index) = 0 Then
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Up, GetPlayerX(Index), MAX_MAPY)
                    Exit Sub
                End If
            End If

        Case DIR_DOWN
            If Map(GetPlayerMap(Index)).Down > 0 Then
                If GetPlayerY(Index) = MAX_MAPY Then
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Down, GetPlayerX(Index), 0)
                    Exit Sub
                End If
            End If

        Case DIR_LEFT
            If Map(GetPlayerMap(Index)).Left > 0 Then
                If GetPlayerX(Index) = 0 Then
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Left, MAX_MAPX, GetPlayerY(Index))
                    Exit Sub
                End If
            End If

        Case DIR_RIGHT
            If Map(GetPlayerMap(Index)).Right > 0 Then
                If GetPlayerX(Index) = MAX_MAPX Then
                    Call PlayerWarp(Index, Map(GetPlayerMap(Index)).Right, 0, GetPlayerY(Index))
                    Exit Sub
                End If
            End If
    End Select
End Sub

Public Sub Packet_EndShot(ByVal Index As Long, ByVal Unknown As Long)
    If Unknown = 0 Then
        Call SendDataTo(Index, POut.CheckForMap & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & END_CHAR)
        Player(Index).Locked = False
        Player(Index).HookShotX = 0
        Player(Index).HookShotY = 0
        Exit Sub
    End If

    Call PlayerMsg(Index, "You carefully cross the wire.", 1)

    Player(Index).Locked = False

    Call SetPlayerX(Index, Player(Index).HookShotX)
    Call SetPlayerY(Index, Player(Index).HookShotY)

    Player(Index).HookShotX = 0
    Player(Index).HookShotY = 0

    Call SendPlayerXY(Index)
End Sub

Public Sub Packet_Attack(ByVal Index As Long)
    Dim I As Long
    Dim Damage As Long

    If Player(Index).LockedAttack Then
        Exit Sub
    End If

    If GetPlayerWeaponSlot(Index) > 0 Then
        If Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data3 > 0 Then
            If Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Stackable = 0 Then
                Call SendDataToMap(GetPlayerMap(Index), POut.CheckArrows & SEP_CHAR & Index & SEP_CHAR & Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data3 & SEP_CHAR & GetPlayerDir(Index) & END_CHAR)
            Else
                Call GrapleHook(Index)
            End If

            Exit Sub
        End If
    End If

    ' Try to attack another player.
    For I = 1 To MAX_PLAYERS
        If I <> Index Then
            If CanAttackPlayer(Index, I) Then
            
                Player(Index).Target = I
                Player(Index).TargetType = TARGET_TYPE_PLAYER
            
                If Not CanPlayerBlockHit(I) Then
                    If Not CanPlayerCriticalHit(Index) Then
                        Damage = GetPlayerDamage(Index) - GetPlayerProtection(I)
                        Call SendDataToMap(GetPlayerMap(Index), POut.Sound & SEP_CHAR & "attack" & END_CHAR)
                    Else
                        Damage = GetPlayerDamage(Index) + Int(Rnd * Int(GetPlayerDamage(Index) / 2)) + 1 - GetPlayerProtection(I)

                        Call BattleMsg(Index, "You feel a surge of energy upon swinging!", BRIGHTCYAN, 0)
                        Call BattleMsg(I, GetPlayerName(Index) & " swings with enormous might!", BRIGHTCYAN, 1)

                        Call SendDataToMap(GetPlayerMap(Index), POut.Sound & SEP_CHAR & "critical" & END_CHAR)
                    End If

                    If Damage > 0 Then
                        If SCRIPTING = 1 Then
                            MyScript.ExecuteStatement "Scripts\Main.txt", "OnAttack " & Index & "," & Damage
                        Else
                            Call AttackPlayer(Index, I, Damage)
                        End If
                    Else
                        If SCRIPTING = 1 Then
                            MyScript.ExecuteStatement "Scripts\Main.txt", "OnAttack " & Index & "," & Damage
                        End If
                        Call PlayerMsg(Index, "Your attack does nothing.", BRIGHTRED)
                        Call SendDataToMap(GetPlayerMap(Index), POut.Sound & SEP_CHAR & "miss" & END_CHAR)
                    End If
                Else
                    If SCRIPTING = 1 Then
                        MyScript.ExecuteStatement "Scripts\Main.txt", "OnAttack " & Index & "," & 0
                    End If

                    Call BattleMsg(Index, GetPlayerName(I) & " blocked your hit!", BRIGHTCYAN, 0)
                    Call BattleMsg(I, "You blocked " & GetPlayerName(Index) & "'s hit!", BRIGHTCYAN, 1)

                    Call SendDataToMap(GetPlayerMap(Index), POut.Sound & SEP_CHAR & "miss" & END_CHAR)
                End If

                Exit Sub
            End If
        End If
    Next I

    ' Try to attack an NPC.
    For I = 1 To MAX_MAP_NPCS
        If CanAttackNpc(Index, I) Then
            Player(Index).TargetNPC = I
            Player(Index).TargetType = TARGET_TYPE_NPC

            If Not CanPlayerCriticalHit(Index) Then
                Damage = GetPlayerDamage(Index) - Int(NPC(MapNPC(GetPlayerMap(Index), I).num).DEF / 2)
                Call SendDataToMap(GetPlayerMap(Index), POut.Sound & SEP_CHAR & "attack" & END_CHAR)
            Else
                Damage = GetPlayerDamage(Index) + Int(Rnd * Int(GetPlayerDamage(Index) / 2)) + 1 - Int(NPC(MapNPC(GetPlayerMap(Index), I).num).DEF / 2)
                Call BattleMsg(Index, "You feel a surge of energy upon swinging!", BRIGHTCYAN, 0)
                Call SendDataToMap(GetPlayerMap(Index), POut.Sound & SEP_CHAR & "critical" & END_CHAR)
            End If

            If Damage > 0 Then
                If SCRIPTING = 1 Then
                    MyScript.ExecuteStatement "Scripts\Main.txt", "OnAttack " & Index & "," & Damage
                Else
                    Call AttackNpc(Index, I, Damage)
                    Call SendDataTo(Index, POut.DrawPlayerDamage & SEP_CHAR & Damage & SEP_CHAR & I & END_CHAR)
                End If
            Else
                If SCRIPTING = 1 Then
                    MyScript.ExecuteStatement "Scripts\Main.txt", "OnAttack " & Index & "," & Damage
                End If
                
                Call BattleMsg(Index, "Your attack does nothing.", BRIGHTRED, 0)

                Call SendDataTo(Index, POut.DrawPlayerDamage & SEP_CHAR & Damage & SEP_CHAR & I & END_CHAR)
                Call SendDataToMap(GetPlayerMap(Index), POut.Sound & SEP_CHAR & "miss" & END_CHAR)
            End If

            Exit Sub
        End If
    Next I
End Sub

Public Sub Packet_UseStatPoint(ByVal Index As Long, ByVal PointType As Long)
    If PointType < 0 Or PointType > 3 Then
        Call HackingAttempt(Index, "Invalid Point Type")
        Exit Sub
    End If

    If GetPlayerPOINTS(Index) > 0 Then
        If SCRIPTING = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "UsingStatPoints " & Index & "," & PointType
        Else
            Select Case PointType
                Case 0
                    Call SetPlayerSTR(Index, GetPlayerSTR(Index) + 1)
                    Call BattleMsg(Index, "You have gained more strength!", WHITE, 0)

                Case 1
                    Call SetPlayerDEF(Index, GetPlayerDEF(Index) + 1)
                    Call BattleMsg(Index, "You have gained more defense!", WHITE, 0)

                Case 2
                    Call SetPlayerMAGI(Index, GetPlayerMAGI(Index) + 1)
                    Call BattleMsg(Index, "You have gained more magic!", WHITE, 0)

                Case 3
                    Call SetPlayerSPEED(Index, GetPlayerSPEED(Index) + 1)
                    Call BattleMsg(Index, "You have gained more speed!", WHITE, 0)
            End Select

            Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) - 1)
        End If
    Else
        Call BattleMsg(Index, "You have no stat points to train with!", BRIGHTRED, 0)
    End If

    Call SendHP(Index)
    Call SendMP(Index)
    Call SendSP(Index)

    Player(Index).Char(Player(Index).CharNum).MAXHP = GetPlayerMaxHP(Index)
    Player(Index).Char(Player(Index).CharNum).MAXMP = GetPlayerMaxMP(Index)
    Player(Index).Char(Player(Index).CharNum).MAXSP = GetPlayerMaxSP(Index)

    Call SendStats(Index)

    Call SendDataTo(Index, POut.PlayerPoints & SEP_CHAR & GetPlayerPOINTS(Index) & END_CHAR)
End Sub

Public Sub Packet_GetStats(ByVal Index As Long, ByVal Name As String)
    Dim PlayerID As Long
    Dim BlockChance As Long
    Dim CritChance As Long

    PlayerID = FindPlayer(Name)

    If PlayerID > 0 Then
        Call PlayerMsg(Index, "Account: " & Trim$(Player(PlayerID).Login) & "; Name: " & GetPlayerName(PlayerID), BRIGHTGREEN)

        If GetPlayerAccess(Index) > ADMIN_MONITER Then
            Call PlayerMsg(Index, "Stats for " & GetPlayerName(PlayerID) & ":", BRIGHTGREEN)
            Call PlayerMsg(Index, "Level: " & GetPlayerLevel(PlayerID) & "; EXP: " & GetPlayerExp(PlayerID) & "/" & GetPlayerNextLevel(PlayerID), BRIGHTGREEN)
            Call PlayerMsg(Index, "HP: " & GetPlayerHP(PlayerID) & "/" & GetPlayerMaxHP(PlayerID) & "; MP: " & GetPlayerMP(PlayerID) & "/" & GetPlayerMaxMP(PlayerID) & "; SP: " & GetPlayerSP(PlayerID) & "/" & GetPlayerMaxSP(PlayerID), BRIGHTGREEN)
            Call PlayerMsg(Index, "STR: " & GetPlayerSTR(PlayerID) & "; DEF: " & GetPlayerDEF(PlayerID) & "; MGC: " & GetPlayerMAGI(PlayerID) & "; SPD: " & GetPlayerSPEED(PlayerID), BRIGHTGREEN)
            
            CritChance = Int(GetPlayerSTR(PlayerID) / 2) + Int(GetPlayerLevel(PlayerID) / 2)
            If CritChance < 0 Then
                CritChance = 0
            End If
            If CritChance > 100 Then
                CritChance = 100
            End If

            BlockChance = Int(GetPlayerDEF(PlayerID) / 2) + Int(GetPlayerLevel(PlayerID) / 2)
            If BlockChance < 0 Then
                BlockChance = 0
            End If
            If BlockChance > 100 Then
                BlockChance = 100
            End If

            Call PlayerMsg(Index, "Critical Chance: " & CritChance & "%; Block Chance: " & BlockChance & "%", BRIGHTGREEN)
        End If
    Else
        Call PlayerMsg(Index, Name & " is currently not online.", WHITE)
    End If
End Sub

Public Sub Packet_SetPlayerSprite(ByVal Index As Long, ByVal Name As String, ByVal SpriteID As Long)
    Dim PlayerID As Long

    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    PlayerID = FindPlayer(Name)

    If PlayerID > 0 Then
        Call SetPlayerSprite(PlayerID, SpriteID)
        Call SendPlayerData(PlayerID)
    Else
        Call PlayerMsg(Index, Name & " is currently not online.", WHITE)
    End If
End Sub

Public Sub Packet_RequestNewMap(ByVal Index As Long, ByVal Dir As Long)
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Call HackingAttempt(Index, "Invalid Direction")
        Exit Sub
    End If

    Call PlayerMove(Index, Dir, 1)
End Sub

Public Sub Packet_WarpMeTo(ByVal Index As Long, ByVal Name As String)
    Dim PlayerID As Long

    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    PlayerID = FindPlayer(Name)

    If PlayerID > 0 Then
        Call PlayerWarp(Index, GetPlayerMap(PlayerID), GetPlayerX(PlayerID), GetPlayerY(PlayerID))
    Else
        Call PlayerMsg(Index, Name & " is currently not online.", WHITE)
    End If
End Sub

Public Sub Packet_WarpToMe(ByVal Index As Long, ByVal Name As String)
    Dim PlayerID As Long

    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
    
    PlayerID = FindPlayer(Name)

    If PlayerID > 0 Then
        Call PlayerWarp(PlayerID, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
    Else
        Call PlayerMsg(Index, Name & " is currently not online.", WHITE)
    End If
End Sub


Public Sub Packet_MapData(ByVal Index As Long, ByRef MapData() As String)
    Dim MapIndex As Long
    Dim MapNum As Long
    Dim MapRevision As Long
    Dim X As Long
    Dim Y As Long
    Dim I As Long
    
    ' Check to see if the user is at least a mapper.
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
            
    MapNum = GetPlayerMap(Index)
            
    ' Get revision number before it clears
    MapRevision = Map(MapNum).Revision + 1
            
    MapIndex = 1

    Call ClearMap(MapNum)

    MapNum = CLng(MapData(MapIndex))
    Map(MapNum).Name = MapData(MapIndex + 1)
    Map(MapNum).Revision = MapRevision
    Map(MapNum).Moral = CLng(MapData(MapIndex + 3))
    Map(MapNum).Up = CLng(MapData(MapIndex + 4))
    Map(MapNum).Down = CLng(MapData(MapIndex + 5))
    Map(MapNum).Left = CLng(MapData(MapIndex + 6))
    Map(MapNum).Right = CLng(MapData(MapIndex + 7))
    Map(MapNum).music = MapData(MapIndex + 8)
    Map(MapNum).BootMap = CLng(MapData(MapIndex + 9))
    Map(MapNum).BootX = CLng(MapData(MapIndex + 10))
    Map(MapNum).BootY = CLng(MapData(MapIndex + 11))
    Map(MapNum).Indoors = CLng(MapData(MapIndex + 12))
    Map(MapNum).Weather = CLng(MapData(MapIndex + 13))

    MapIndex = MapIndex + 14

    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            Map(MapNum).Tile(X, Y).Ground = CLng(MapData(MapIndex))
            Map(MapNum).Tile(X, Y).Mask = CLng(MapData(MapIndex + 1))
            Map(MapNum).Tile(X, Y).Anim = CLng(MapData(MapIndex + 2))
            Map(MapNum).Tile(X, Y).Mask2 = CLng(MapData(MapIndex + 3))
            Map(MapNum).Tile(X, Y).M2Anim = CLng(MapData(MapIndex + 4))
            Map(MapNum).Tile(X, Y).Fringe = CLng(MapData(MapIndex + 5))
            Map(MapNum).Tile(X, Y).FAnim = CLng(MapData(MapIndex + 6))
            Map(MapNum).Tile(X, Y).Fringe2 = CLng(MapData(MapIndex + 7))
            Map(MapNum).Tile(X, Y).F2Anim = CLng(MapData(MapIndex + 8))
            Map(MapNum).Tile(X, Y).Type = CLng(MapData(MapIndex + 9))
            Map(MapNum).Tile(X, Y).Data1 = CLng(MapData(MapIndex + 10))
            Map(MapNum).Tile(X, Y).Data2 = CLng(MapData(MapIndex + 11))
            Map(MapNum).Tile(X, Y).Data3 = CLng(MapData(MapIndex + 12))
            Map(MapNum).Tile(X, Y).String1 = MapData(MapIndex + 13)
            Map(MapNum).Tile(X, Y).String2 = MapData(MapIndex + 14)
            Map(MapNum).Tile(X, Y).String3 = MapData(MapIndex + 15)
            Map(MapNum).Tile(X, Y).Light = CLng(MapData(MapIndex + 16))
            Map(MapNum).Tile(X, Y).GroundSet = CLng(MapData(MapIndex + 17))
            Map(MapNum).Tile(X, Y).MaskSet = CLng(MapData(MapIndex + 18))
            Map(MapNum).Tile(X, Y).AnimSet = CLng(MapData(MapIndex + 19))
            Map(MapNum).Tile(X, Y).Mask2Set = CLng(MapData(MapIndex + 20))
            Map(MapNum).Tile(X, Y).M2AnimSet = CLng(MapData(MapIndex + 21))
            Map(MapNum).Tile(X, Y).FringeSet = CLng(MapData(MapIndex + 22))
            Map(MapNum).Tile(X, Y).FAnimSet = CLng(MapData(MapIndex + 23))
            Map(MapNum).Tile(X, Y).Fringe2Set = CLng(MapData(MapIndex + 24))
            Map(MapNum).Tile(X, Y).F2AnimSet = CLng(MapData(MapIndex + 25))

            MapIndex = MapIndex + 26
        Next X
    Next Y

    For X = 1 To MAX_MAP_NPCS
        Map(MapNum).NPC(X) = CLng(MapData(MapIndex))
        Map(MapNum).SpawnX(X) = CLng(MapData(MapIndex + 1))
        Map(MapNum).SpawnY(X) = CLng(MapData(MapIndex + 2))
        MapIndex = MapIndex + 3
        Call ClearMapNpc(X, MapNum)
    Next X

    ' Clear out it all
    For I = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(I, 0, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), I).X, MapItem(GetPlayerMap(Index), I).Y)
        Call ClearMapItem(I, GetPlayerMap(Index))
    Next I

    ' Save the map
    Call SaveMap(MapNum)
            
    ' Mapper is on the map
    PlayersOnMap(MapNum) = YES

    ' Respawn
    Call SpawnMapItems(GetPlayerMap(Index))

    ' Respawn NPCS
    For I = 1 To MAX_MAP_NPCS
        Call SpawnNPC(I, GetPlayerMap(Index))
    Next I

    ' Refresh map for everyone online
    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If GetPlayerMap(I) = MapNum Then
                Call SendDataTo(I, POut.CheckForMap & SEP_CHAR & GetPlayerMap(I) & SEP_CHAR & Map(GetPlayerMap(I)).Revision & END_CHAR)
            End If
        End If
    Next I
End Sub

Public Sub Packet_NeedMap(ByVal Index As Long, ByVal NeedMap As Long)
    Dim I As Long

    If NeedMap = 1 Then
        Call SendMap(Index, GetPlayerMap(Index))
    End If

    Call SendMapItemsTo(Index, GetPlayerMap(Index))
    Call SendMapNpcsTo(Index, GetPlayerMap(Index))
    Call SendJoinMap(Index)
    Call SendDataTo(Index, POut.MapDone & END_CHAR)

    Player(Index).GettingMap = NO

    Call SendPlayerData(Index)

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            Call SendIndexWornEquipment(I)
            Call SendWornEquipment(I)
        End If
    Next I
End Sub

Public Sub Packet_MapGetItem(ByVal Index As Long)
    Call PlayerMapGetItem(Index)
End Sub

Public Sub Packet_MapDropItem(ByVal Index As Long, ByVal InvNum As Long, ByVal Amount As Long)
    If InvNum < 1 Or InvNum > MAX_INV Then
        Call HackingAttempt(Index, "Invalid InvNum")
        Exit Sub
    End If

    ' Check if the item is bound to the character.
    If Item(GetPlayerInvItemNum(Index, InvNum)).Bound = 1 Then
        Call PlayerMsg(Index, "You cannot drop items bound to your character.", WHITE)
        Exit Sub
    End If

    ' Check if the player is in a trade.
    If Player(Index).InTrade Then
        Call PlayerMsg(Index, "You cannot drop items while trading.", WHITE)
        Exit Sub
    End If

    ' Prevent hacking
    If Item(GetPlayerInvItemNum(Index, InvNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(Index, InvNum)).Stackable = 1 Then
        If Amount <= 0 Then
            Call PlayerMsg(Index, "You must at least drop 1 of that item!", BRIGHTRED)
            Exit Sub
        End If

        If Amount > GetPlayerInvItemValue(Index, InvNum) Then
            Call PlayerMsg(Index, "You don't have that much to drop!", BRIGHTRED)
            Exit Sub
        End If
    End If

    ' Prevent hacking
    If Item(GetPlayerInvItemNum(Index, InvNum)).Type <> ITEM_TYPE_CURRENCY Then
        If Item(GetPlayerInvItemNum(Index, InvNum)).Stackable = 1 Then
            If Amount > GetPlayerInvItemValue(Index, InvNum) Then
                Call HackingAttempt(Index, "Item amount modification")
                Exit Sub
            End If
        End If
    End If

    Call PlayerMapDropItem(Index, InvNum, Amount)

    Call SendStats(Index)
    Call SendHP(Index)
    Call SendMP(Index)
    Call SendSP(Index)
End Sub

Public Sub Packet_MapRespawn(ByVal Index As Long)
    Dim I As Long

    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    ' Clear out all of the floor items.
    For I = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(I, 0, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), I).X, MapItem(GetPlayerMap(Index), I).Y)
        Call ClearMapItem(I, GetPlayerMap(Index))
    Next I

    ' Respawn all of the floor items.
    Call SpawnMapItems(GetPlayerMap(Index))

    ' Respawn NPCS
    For I = 1 To MAX_MAP_NPCS
        Call SpawnNPC(I, GetPlayerMap(Index))
    Next I

    Call PlayerMsg(Index, "Map respawned.", BLUE)
End Sub

Public Sub Packet_KickPlayer(ByVal Index As Long, ByVal Name As String)
    Dim PlayerIndex As Long

    If GetPlayerAccess(Index) < 1 Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    PlayerIndex = FindPlayer(Name)

    If PlayerIndex > 0 Then
        If PlayerIndex <> Index Then
            If GetPlayerAccess(PlayerIndex) <= GetPlayerAccess(Index) Then
                Call GlobalMsg(GetPlayerName(PlayerIndex) & " has been kicked from " & GAME_NAME & " by " & GetPlayerName(Index) & "!", WHITE)
                Call AddLog(GetPlayerName(Index) & " has kicked " & GetPlayerName(PlayerIndex) & ".", ADMIN_LOG)
                Call AlertMsg(PlayerIndex, "You have been kicked by " & GetPlayerName(Index) & "!")
            Else
                Call PlayerMsg(Index, "That admin has a higher access then you!", WHITE)
            End If
        Else
            Call PlayerMsg(Index, "You cannot kick yourself!", WHITE)
        End If
    Else
        Call PlayerMsg(Index, "Player is not online.", WHITE)
    End If
End Sub

Public Sub Packet_BanList(ByVal Index As Long)
    Dim FileID As Integer
    Dim PlayerName As String

    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If
            
    If Not FileExists("BanList.txt") Then
        Call PlayerMsg(Index, "The ban list cannot be found!", BRIGHTRED)
        Exit Sub
    End If

    FileID = FreeFile

    Open App.Path & "\BanList.txt" For Input As #FileID
    Do While Not EOF(FileID)
        Line Input #FileID, PlayerName
        Call PlayerMsg(Index, PlayerName, WHITE)
    Loop
    Close #FileID
End Sub

Public Sub Packet_BanListDestroy(ByVal Index As Long)
    If GetPlayerAccess(Index) < ADMIN_CREATOR Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    If FileExists("BanList.txt") Then
        Call Kill(App.Path & "\BanList.txt")
    End If

    Call PlayerMsg(Index, "Ban list destroyed.", WHITE)
End Sub

Public Sub Packet_BanPlayer(ByVal Index As Long, ByVal Name As String)
    Dim PlayerIndex As Long

    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    PlayerIndex = FindPlayer(Name)

    If PlayerIndex > 0 Then
        If PlayerIndex <> Index Then
            If GetPlayerAccess(PlayerIndex) <= GetPlayerAccess(Index) Then
                Call BanIndex(PlayerIndex, Index)
            Else
                Call PlayerMsg(Index, "That admin has a higher access then you!", WHITE)
            End If
        Else
            Call PlayerMsg(Index, "You cannot ban yourself!", WHITE)
        End If
    Else
        Call PlayerMsg(Index, "Player is not online.", WHITE)
    End If
End Sub

Public Sub Packet_RequestEditMap(ByVal Index As Long)
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    Call SendDataTo(Index, POut.EditMap & END_CHAR)
End Sub

Public Sub Packet_RequestEditItem(ByVal Index As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    Call SendDataTo(Index, POut.ItemEditor & END_CHAR)
End Sub

Public Sub Packet_EditItem(ByVal Index As Long, ByVal ItemNum As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    If ItemNum < 0 Or ItemNum > MAX_ITEMS Then
        Call HackingAttempt(Index, "Invalid Item Index")
        Exit Sub
    End If

    Call SendEditItemTo(Index, ItemNum)

    Call AddLog(GetPlayerName(Index) & " editing item #" & ItemNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_SaveItem(ByVal Index As Long, ByRef ItemData() As String)
    Dim ItemNum As Long

    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    ItemNum = CLng(ItemData(1))

    If ItemNum < 0 Or ItemNum > MAX_ITEMS Then
        Call HackingAttempt(Index, "Invalid Item Index")
        Exit Sub
    End If

    Item(ItemNum).Name = ItemData(2)
    Item(ItemNum).Pic = CLng(ItemData(3))
    Item(ItemNum).Type = CLng(ItemData(4))
    Item(ItemNum).Data1 = CLng(ItemData(5))
    Item(ItemNum).Data2 = CLng(ItemData(6))
    Item(ItemNum).Data3 = CLng(ItemData(7))
    Item(ItemNum).StrReq = CLng(ItemData(8))
    Item(ItemNum).DefReq = CLng(ItemData(9))
    Item(ItemNum).SpeedReq = CLng(ItemData(10))
    Item(ItemNum).MagicReq = CLng(ItemData(11))
    Item(ItemNum).ClassReq = CLng(ItemData(12))
    Item(ItemNum).AccessReq = CLng(ItemData(13))

    Item(ItemNum).AddHP = CLng(ItemData(14))
    Item(ItemNum).AddMP = CLng(ItemData(15))
    Item(ItemNum).AddSP = CLng(ItemData(16))
    Item(ItemNum).AddStr = CLng(ItemData(17))
    Item(ItemNum).AddDef = CLng(ItemData(18))
    Item(ItemNum).AddMagi = CLng(ItemData(19))
    Item(ItemNum).AddSpeed = CLng(ItemData(20))
    Item(ItemNum).AddEXP = CLng(ItemData(21))
    Item(ItemNum).Desc = ItemData(22)
    Item(ItemNum).AttackSpeed = CLng(ItemData(23))
    Item(ItemNum).Price = CLng(ItemData(24))
    Item(ItemNum).Stackable = CLng(ItemData(25))
    Item(ItemNum).Bound = CLng(ItemData(26))

    Call SendUpdateItemToAll(ItemNum)
    Call SaveItem(ItemNum)

    Call AddLog(GetPlayerName(Index) & " saved item #" & ItemNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_EnableDayNight(ByVal Index As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    If Not TimeDisable Then
        Gamespeed = 0
        frmServer.GameTimeSpeed.Text = 0
        TimeDisable = True
        frmServer.tmrDayNight.Enabled = False
        frmServer.Command69.Caption = "Enable Time"
    Else
        Gamespeed = 1
        frmServer.GameTimeSpeed.Text = 1
        TimeDisable = False
        frmServer.tmrDayNight.Enabled = True
        frmServer.Command69.Caption = "Disable Time"
    End If
End Sub

Public Sub Packet_DayNight(ByVal Index As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    If Hours > 12 Then
        Hours = Hours - 12
    Else
        Hours = Hours + 12
    End If
End Sub

Public Sub Packet_RequestEditNPC(ByVal Index As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    Call SendDataTo(Index, POut.NPCEditor & END_CHAR)
End Sub

Public Sub Packet_EditNPC(ByVal Index As Long, ByVal NPCnum As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    If NPCnum < 0 Or NPCnum > MAX_NPCS Then
        Call HackingAttempt(Index, "Invalid NPC Index")
        Exit Sub
    End If

    Call SendEditNpcTo(Index, NPCnum)

    Call AddLog(GetPlayerName(Index) & " editing npc #" & NPCnum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_SaveNPC(ByVal Index As Long, ByRef NPCData() As String)
    Dim NPCnum As Long
    Dim NPCIndex As Long
    Dim I As Long

    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    NPCnum = CLng(NPCData(1))

    If NPCnum < 0 Or NPCnum > MAX_NPCS Then
        Call HackingAttempt(Index, "Invalid NPC Index")
        Exit Sub
    End If

    NPC(NPCnum).Name = NPCData(2)
    NPC(NPCnum).AttackSay = NPCData(3)
    NPC(NPCnum).Sprite = CLng(NPCData(4))
    NPC(NPCnum).SpawnSecs = CLng(NPCData(5))
    NPC(NPCnum).Behavior = CLng(NPCData(6))
    NPC(NPCnum).Range = CLng(NPCData(7))
    NPC(NPCnum).STR = CLng(NPCData(8))
    NPC(NPCnum).DEF = CLng(NPCData(9))
    NPC(NPCnum).Speed = CLng(NPCData(10))
    NPC(NPCnum).Magi = CLng(NPCData(11))
    NPC(NPCnum).Big = CLng(NPCData(12))
    NPC(NPCnum).MAXHP = CLng(NPCData(13))
    NPC(NPCnum).Exp = CLng(NPCData(14))
    NPC(NPCnum).SpawnTime = CLng(NPCData(15))
    NPC(NPCnum).Element = CLng(NPCData(16))
    NPC(NPCnum).SPRITESIZE = CLng(NPCData(17))

    NPCIndex = 18

    For I = 1 To MAX_NPC_DROPS
        NPC(NPCnum).ItemNPC(I).Chance = CLng(NPCData(NPCIndex))
        NPC(NPCnum).ItemNPC(I).ItemNum = CLng(NPCData(NPCIndex + 1))
        NPC(NPCnum).ItemNPC(I).ItemValue = CLng(NPCData(NPCIndex + 2))
        NPCIndex = NPCIndex + 3
    Next I

    Call SendUpdateNpcToAll(NPCnum)
    Call SaveNPC(NPCnum)

    Call AddLog(GetPlayerName(Index) & " saved npc #" & NPCnum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_RequestEditShop(ByVal Index As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    Call SendDataTo(Index, POut.ShopEditor & END_CHAR)
End Sub

Public Sub Packet_EditShop(ByVal Index As Long, ByVal ShopNum As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    If ShopNum < 0 Or ShopNum > MAX_SHOPS Then
        Call HackingAttempt(Index, "Invalid Shop Index")
        Exit Sub
    End If

    Call SendEditShopTo(Index, ShopNum)

    Call AddLog(GetPlayerName(Index) & " editing shop #" & ShopNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_SaveShop(ByVal Index As Long, ByRef ShopData() As String)
    Dim ShopNum As Long
    Dim ShopIndex As Long
    Dim I As Long

    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    ShopNum = CLng(ShopData(1))

    If ShopNum < 1 Or ShopNum > MAX_SHOPS Then
        Call HackingAttempt(Index, "Invalid Shop Index")
        Exit Sub
    End If

    Shop(ShopNum).Name = ShopData(2)
    Shop(ShopNum).FixesItems = CLng(ShopData(3))
    Shop(ShopNum).BuysItems = CLng(ShopData(4))
    Shop(ShopNum).CurrencyItem = CLng(ShopData(5))

    ShopIndex = 6

    For I = 1 To MAX_SHOP_ITEMS
        Shop(ShopNum).ShopItem(I).ItemNum = CLng(ShopData(ShopIndex))
        Shop(ShopNum).ShopItem(I).Amount = CLng(ShopData(ShopIndex + 1))
        Shop(ShopNum).ShopItem(I).Price = CLng(ShopData(ShopIndex + 2))
        ShopIndex = ShopIndex + 3
    Next I

    Call SendUpdateShopToAll(ShopNum)
    Call SaveShop(ShopNum)

    Call AddLog(GetPlayerName(Index) & " saving shop #" & ShopNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_RequestEditSpell(ByVal Index As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    Call SendDataTo(Index, POut.SpellEditor & END_CHAR)
End Sub

Public Sub Packet_EditSpell(ByVal Index As Long, ByVal SpellNum As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    If SpellNum < 0 Or SpellNum > MAX_SPELLS Then
        Call HackingAttempt(Index, "Invalid Spell Index")
        Exit Sub
    End If

    Call SendEditSpellTo(Index, SpellNum)

    Call AddLog(GetPlayerName(Index) & " editing spell #" & SpellNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_SaveSpell(ByVal Index As Long, ByRef SpellData() As String)
    Dim SpellNum As Long
    
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    SpellNum = CLng(SpellData(1))

    If SpellNum < 1 Or SpellNum > MAX_SPELLS Then
        Call HackingAttempt(Index, "Invalid Spell Index")
        Exit Sub
    End If

    Spell(SpellNum).Name = SpellData(2)
    Spell(SpellNum).ClassReq = CLng(SpellData(3))
    Spell(SpellNum).LevelReq = CLng(SpellData(4))
    Spell(SpellNum).Type = CLng(SpellData(5))
    Spell(SpellNum).Data1 = CLng(SpellData(6))
    Spell(SpellNum).Data2 = CLng(SpellData(7))
    Spell(SpellNum).Data3 = CLng(SpellData(8))
    Spell(SpellNum).MPCost = CLng(SpellData(9))
    Spell(SpellNum).Sound = CLng(SpellData(10))
    Spell(SpellNum).Range = CLng(SpellData(11))
    Spell(SpellNum).SpellAnim = CLng(SpellData(12))
    Spell(SpellNum).SpellTime = CLng(SpellData(13))
    Spell(SpellNum).SpellDone = CLng(SpellData(14))
    Spell(SpellNum).AE = CLng(SpellData(15))
    Spell(SpellNum).Big = CLng(SpellData(16))
    Spell(SpellNum).Element = CLng(SpellData(17))

    Call SendUpdateSpellToAll(SpellNum)
    Call SaveSpell(SpellNum)

    Call AddLog(GetPlayerName(Index) & " saving spell #" & SpellNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_ForgetSpell(ByVal Index As Long, ByVal SpellNum As Long)
    If SpellNum < 1 Or SpellNum > MAX_PLAYER_SPELLS Then
        Call HackingAttempt(Index, "Invalid Spell Slot")
        Exit Sub
    End If

    With Player(Index).Char(Player(Index).CharNum)
        If .Spell(SpellNum) = 0 Then
            Call PlayerMsg(Index, "No spell here.", RED)
        Else
            Call PlayerMsg(Index, "You have forgotten the spell " & Trim$(Spell(.Spell(SpellNum)).Name) & ".", GREEN)

            .Spell(SpellNum) = 0

            Call SendSpells(Index)
        End If
    End With
End Sub

Public Sub Packet_SetAccess(ByVal Index As Long, ByVal Name As String, ByVal AccessLvl As Long)
    Dim PlayerIndex As Long
    
    If GetPlayerAccess(Index) < ADMIN_CREATOR Then
        Call HackingAttempt(Index, "Invalid Access")
        Exit Sub
    End If
    
    If AccessLvl < 0 Or AccessLvl > 5 Then
        Call PlayerMsg(Index, "You have entered an invalid access level.", BRIGHTRED)
        Exit Sub
    End If

    PlayerIndex = FindPlayer(Name)

    If PlayerIndex > 0 Then
        If GetPlayerName(Index) <> GetPlayerName(PlayerIndex) Then
            If GetPlayerAccess(Index) > GetPlayerAccess(PlayerIndex) Then
                Call SetPlayerAccess(PlayerIndex, AccessLvl)
                Call SendPlayerData(PlayerIndex)
    
                If GetPlayerAccess(PlayerIndex) = 0 Then
                    Call GlobalMsg(GetPlayerName(PlayerIndex) & " has been blessed with administrative access.", BRIGHTBLUE)
                End If
    
                Call AddLog(GetPlayerName(Index) & " has modified " & GetPlayerName(PlayerIndex) & "'s access.", ADMIN_LOG)
            Else
                Call PlayerMsg(Index, "Your access level is lower than " & GetPlayerName(PlayerIndex) & ".", RED)
            End If
        Else
            Call PlayerMsg(Index, "You cant change your access.", RED)
        End If
    Else
        Call PlayerMsg(Index, "Player is not online.", WHITE)
    End If
End Sub

Public Sub Packet_WhoIsOnline(ByVal Index As Long)
    Call SendWhosOnline(Index)
End Sub

Public Sub Packet_OnlineList(ByVal Index As Long)
    Call SendOnlineList
End Sub

Public Sub Packet_SetMOTD(ByVal Index As Long, ByVal MOTD As String)
    If GetPlayerAccess(Index) < ADMIN_MAPPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    Call PutVar(App.Path & "\MOTD.ini", "MOTD", "Msg", MOTD)
            
    If SCRIPTING = 1 Then
        MyScript.ExecuteStatement "Scripts\Main.txt", "ChangeMOTD"
    End If
            
    Call GlobalMsg("MOTD changed to: " & MOTD, BRIGHTCYAN)

    Call AddLog(GetPlayerName(Index) & " changed MOTD to: " & MOTD, ADMIN_LOG)
End Sub

Public Sub Packet_BuyItem(ByVal Index As Long, ByVal ShopIndex As Long, ByVal ItemIndex As Long)
    Dim InvItem As Long

    If ShopIndex < 1 Or ShopIndex > MAX_SHOPS Then
        Call HackingAttempt(Index, "Invalid Shop Index")
        Exit Sub
    End If
    
    If ItemIndex < 1 Or ItemIndex > MAX_SHOP_ITEMS Then
        Call HackingAttempt(Index, "Invalid Shop Item")
        Exit Sub
    End If

    ' Check to see if player's inventory is full.
    InvItem = FindOpenInvSlot(Index, Shop(ShopIndex).ShopItem(ItemIndex).ItemNum)
    If InvItem = 0 Then
        Call PlayerMsg(Index, "Your inventory has reached its maximum capacity!", BRIGHTRED)
        Exit Sub
    End If

    ' Check to see if they have enough currency.
    If HasItem(Index, Shop(ShopIndex).CurrencyItem) >= Shop(ShopIndex).ShopItem(ItemIndex).Price Then
        Call TakeItem(Index, Shop(ShopIndex).CurrencyItem, Shop(ShopIndex).ShopItem(ItemIndex).Price)
        Call GiveItem(Index, Shop(ShopIndex).ShopItem(ItemIndex).ItemNum, Shop(ShopIndex).ShopItem(ItemIndex).Amount)

        Call PlayerMsg(Index, "You bought the item.", YELLOW)
    Else
        Call PlayerMsg(Index, "You cannot afford that!", RED)
    End If
End Sub

Public Sub Packet_SellItem(ByVal Index As Long, ByVal ShopNum As Long, ByVal ItemNum As Long, ByVal ItemSlot As Long, ByVal ItemAmt As Long)
    If ItemIsEquipped(Index, ItemNum) Then
        Call PlayerMsg(Index, "You cannot sell worn items.", RED)
        Exit Sub
    End If

    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
        Call PlayerMsg(Index, "You cannot sell currency.", RED)
        Exit Sub
    End If

    If Item(ItemNum).Stackable = YES Then
        If ItemAmt < 1 Or ItemAmt > GetPlayerInvItemValue(Index, ItemSlot) Then
            Call PlayerMsg(Index, "You don't have enough of that item!", RED)
            Exit Sub
        End If
    End If

    If Item(ItemNum).Price > 0 Then
        Call TakeItem(Index, ItemNum, ItemAmt)
        Call GiveItem(Index, Shop(ShopNum).CurrencyItem, Item(ItemNum).Price * ItemAmt)
        Call PlayerMsg(Index, "The shopkeeper hands you " & Item(ItemNum).Price * ItemAmt & " " & Trim$(Item(Shop(ShopNum).CurrencyItem).Name) & ".", YELLOW)
    Else
        Call PlayerMsg(Index, "This item cannot be sold.", RED)
    End If
End Sub

Public Sub Packet_FixItem(ByVal Index As Long, ByVal ShopNum As Long, ByVal InvNum As Long)
    Dim ItemNum As Long
    Dim DurNeeded As Long
    Dim GoldNeeded As Long
    Dim I As Long

    If Item(GetPlayerInvItemNum(Index, InvNum)).Type < ITEM_TYPE_WEAPON Or Item(GetPlayerInvItemNum(Index, InvNum)).Type > ITEM_TYPE_NECKLACE Then
        Call PlayerMsg(Index, "That item doesn't need to be fixed.", BRIGHTRED)
        Exit Sub
    End If

    If FindOpenInvSlot(Index, GetPlayerInvItemNum(Index, InvNum)) = 0 Then
        Call PlayerMsg(Index, "You have no inventory space left!", BRIGHTRED)
        Exit Sub
    End If

    ItemNum = GetPlayerInvItemNum(Index, InvNum)

    I = Int(Item(GetPlayerInvItemNum(Index, InvNum)).Data2 / 5)
    If I <= 0 Then
        I = 1
    End If

    DurNeeded = Item(ItemNum).Data1 - GetPlayerInvItemDur(Index, InvNum)

    GoldNeeded = Int(DurNeeded * I / 2)
    If GoldNeeded <= 0 Then
        GoldNeeded = 1
    End If

    If DurNeeded = 0 Then
        Call PlayerMsg(Index, "This item is in perfect condition!", WHITE)
        Exit Sub
    End If

    If HasItem(Index, Shop(ShopNum).CurrencyItem) >= I Then
        If HasItem(Index, Shop(ShopNum).CurrencyItem) >= GoldNeeded Then
            Call TakeItem(Index, Shop(ShopNum).CurrencyItem, GoldNeeded)
            Call SetPlayerInvItemDur(Index, InvNum, Item(ItemNum).Data1)

            Call PlayerMsg(Index, "Item has been totally restored for " & GoldNeeded & " gold!", BRIGHTBLUE)
        Else
            DurNeeded = (HasItem(Index, Shop(ShopNum).CurrencyItem) / I)
            GoldNeeded = Int(DurNeeded * I / 2)

            If GoldNeeded <= 0 Then
                GoldNeeded = 1
            End If

            Call TakeItem(Index, Shop(ShopNum).CurrencyItem, GoldNeeded)
            Call SetPlayerInvItemDur(Index, InvNum, GetPlayerInvItemDur(Index, InvNum) + DurNeeded)

            Call PlayerMsg(Index, "Item has been partially fixed for " & GoldNeeded & " gold!", BRIGHTBLUE)
        End If
    Else
        Call PlayerMsg(Index, "You don't have enough gold to fix this item!", BRIGHTRED)
    End If
End Sub

Public Sub Packet_Search(ByVal Index As Long, ByVal X As Long, ByVal Y As Long)
    Dim I As Long

    If X < 0 Or X > MAX_MAPX Then
        Exit Sub
    End If

    If Y < 0 Or Y > MAX_MAPY Then
        Exit Sub
    End If

    ' Check for a player
    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If GetPlayerMap(Index) = GetPlayerMap(I) Then
                If GetPlayerX(I) = X Then
                    If GetPlayerY(I) = Y Then
                        If GetPlayerLevel(I) >= GetPlayerLevel(Index) + 5 Then
                            Call PlayerMsg(Index, "You wouldn't stand a chance.", BRIGHTRED)
                        Else
                            If GetPlayerLevel(I) > GetPlayerLevel(Index) Then
                                Call PlayerMsg(Index, "This one seems to have an advantage over you.", YELLOW)
                            Else
                                If GetPlayerLevel(I) = GetPlayerLevel(Index) Then
                                    Call PlayerMsg(Index, "This would be an even fight.", WHITE)
                                Else
                                    If GetPlayerLevel(Index) >= GetPlayerLevel(I) + 5 Then
                                        Call PlayerMsg(Index, "You could slaughter that player.", BRIGHTBLUE)
                                    Else
                                        If GetPlayerLevel(Index) > GetPlayerLevel(I) Then
                                            Call PlayerMsg(Index, "You would have an advantage over that player.", YELLOW)
                                        End If
                                    End If
                                End If
                            End If
                        End If

                        ' Change the target.
                        Player(Index).Target = I
                        Player(Index).TargetType = TARGET_TYPE_PLAYER

                        Call PlayerMsg(Index, "Your target is now " & GetPlayerName(I) & ".", YELLOW)

                        Exit Sub
                    End If
                End If
            End If

        End If
    Next I

    ' Check for an NPC.
    For I = 1 To MAX_MAP_NPCS
        If MapNPC(GetPlayerMap(Index), I).num > 0 Then
            If MapNPC(GetPlayerMap(Index), I).X = X Then
                If MapNPC(GetPlayerMap(Index), I).Y = Y Then
                    Player(Index).TargetNPC = I
                    Player(Index).TargetType = TARGET_TYPE_NPC

                    Call PlayerMsg(Index, "Your target is now a " & Trim$(NPC(MapNPC(GetPlayerMap(Index), I).num).Name) & ".", YELLOW)

                    Exit Sub
                End If
            End If
        End If
    Next I

    ' Check for an item on the ground.
    For I = 1 To MAX_MAP_ITEMS
        If MapItem(GetPlayerMap(Index), I).num > 0 Then
            If MapItem(GetPlayerMap(Index), I).X = X Then
                If MapItem(GetPlayerMap(Index), I).Y = Y Then
                    Call PlayerMsg(Index, "You see a " & Trim$(Item(MapItem(GetPlayerMap(Index), I).num).Name) & ".", YELLOW)
                    Exit Sub
                End If
            End If
        End If
    Next I

    ' Check for an OnClick tile.
    If Map(GetPlayerMap(Index)).Tile(X, Y).Type = TILE_TYPE_ONCLICK Then
        If SCRIPTING = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "OnClick " & Index & "," & Map(GetPlayerMap(Index)).Tile(X, Y).Data1 & "," & X & "," & Y
        End If
    End If
End Sub

Public Sub Packet_PlayerChat(ByVal Index As Long, ByVal Name As String)
    Dim PlayerIndex As Long

    PlayerIndex = FindPlayer(Name)

    If PlayerIndex = 0 Then
        Call PlayerMsg(Index, Name & " is currently not online.", WHITE)
        Exit Sub
    End If

    If PlayerIndex = Index Then
        Call PlayerMsg(Index, "You cannot chat with yourself.", PINK)
        Exit Sub
    End If

    If Player(Index).InChat = 1 Then
        Call PlayerMsg(Index, "You're already in a chat with another player!", PINK)
        Exit Sub
    End If

    If Player(PlayerIndex).InChat = 1 Then
        Call PlayerMsg(Index, Name & " is already in a chat with another player!", PINK)
        Exit Sub
    End If

    Call PlayerMsg(Index, "Chat request has been sent to " & GetPlayerName(PlayerIndex) & ".", PINK)
    Call PlayerMsg(PlayerIndex, GetPlayerName(Index) & " wants you to chat with them. Type /chat to accept, or /chatdecline to decline.", PINK)

    Player(Index).ChatPlayer = PlayerIndex
    Player(PlayerIndex).ChatPlayer = Index
End Sub

Public Sub Packet_AcceptChat(ByVal Index As Long)
    Dim PlayerIndex As Long

    PlayerIndex = Player(Index).ChatPlayer

    If PlayerIndex = 0 Then
        Call PlayerMsg(Index, "You have not requested to chat with anyone.", PINK)
        Exit Sub
    End If

    If Player(PlayerIndex).ChatPlayer <> Index Then
        Call PlayerMsg(Index, "Chat failed.", PINK)
        Exit Sub
    End If

    Call SendDataTo(Index, POut.PrepareChat & SEP_CHAR & PlayerIndex & END_CHAR)
    Call SendDataTo(PlayerIndex, POut.PrepareChat & SEP_CHAR & Index & END_CHAR)
End Sub

Public Sub Packet_DenyChat(ByVal Index As Long)
    Dim PlayerIndex As Long

    PlayerIndex = Player(Index).ChatPlayer

    If PlayerIndex = 0 Then
        Call PlayerMsg(Index, "You have not requested to chat with anyone.", PINK)
        Exit Sub
    End If

    If Player(PlayerIndex).ChatPlayer <> Index Then
        Call PlayerMsg(Index, "Chat failed.", PINK)
        Exit Sub
    End If

    Call PlayerMsg(Index, "Declined chat request.", PINK)
    Call PlayerMsg(PlayerIndex, GetPlayerName(Index) & " declined your request.", PINK)

    Player(Index).ChatPlayer = 0
    Player(Index).InChat = 0

    Player(PlayerIndex).ChatPlayer = 0
    Player(PlayerIndex).InChat = 0
End Sub

Public Sub Packet_QuitChat(ByVal Index As Long)
    Dim PlayerIndex As Long

    PlayerIndex = Player(Index).ChatPlayer

    If PlayerIndex = 0 Then
        Call PlayerMsg(Index, "You have not requested to chat with anyone.", PINK)
        Exit Sub
    End If

    If Player(PlayerIndex).ChatPlayer <> Index Then
        Call PlayerMsg(Index, "Chat failed.", PINK)
        Exit Sub
    End If

    Call SendDataTo(Index, POut.QuitChat & END_CHAR)
    Call SendDataTo(PlayerIndex, POut.QuitChat & END_CHAR)

    Player(Index).ChatPlayer = 0
    Player(Index).InChat = 0

    Player(PlayerIndex).ChatPlayer = 0
    Player(PlayerIndex).InChat = 0
End Sub

Public Sub Packet_SendChat(ByVal Index As Long, ByVal Message As String)
    Dim PlayerIndex As Long

    PlayerIndex = Player(Index).ChatPlayer

    If PlayerIndex = 0 Then
        Call PlayerMsg(Index, "You have not requested to chat with anyone.", PINK)
        Exit Sub
    End If

    If Player(PlayerIndex).ChatPlayer <> Index Then
        Call PlayerMsg(Index, "Chat failed.", PINK)
        Exit Sub
    End If

    Call SendDataTo(PlayerIndex, POut.SendChat & SEP_CHAR & Message & SEP_CHAR & Index & END_CHAR)
End Sub

Public Sub Packet_PrepareTrade(ByVal Index As Long, ByVal Name As String)
    Dim PlayerIndex As Long

    PlayerIndex = FindPlayer(Name)

    If PlayerIndex = 0 Then
        Call PlayerMsg(Index, Name & " is currently not online.", WHITE)
        Exit Sub
    End If

    If PlayerIndex = Index Then
        Call PlayerMsg(Index, "You cannot trade with yourself!", PINK)
        Exit Sub
    End If

    If GetPlayerMap(Index) <> GetPlayerMap(PlayerIndex) Then
        Call PlayerMsg(Index, "You must be on the same map to trade with " & GetPlayerName(PlayerIndex) & "!", PINK)
        Exit Sub
    End If

    If Player(Index).InTrade Then
        Call PlayerMsg(Index, "You're already in a trade with someone else!", PINK)
        Exit Sub
    End If

    If Player(PlayerIndex).InTrade Then
        Call PlayerMsg(Index, Name & " is already in a trade!", PINK)
        Exit Sub
    End If

    Call PlayerMsg(Index, "Trade request has been sent to " & GetPlayerName(PlayerIndex) & ".", PINK)
    Call PlayerMsg(PlayerIndex, GetPlayerName(Index) & " wants you to trade with them. Type /accept to accept, or /decline to decline.", PINK)

    Player(Index).TradePlayer = PlayerIndex
    Player(PlayerIndex).TradePlayer = Index
End Sub

Public Sub Packet_AcceptTrade(ByVal Index As Long)
    Dim PlayerIndex As Long
    Dim I As Long

    PlayerIndex = Player(Index).TradePlayer

    If PlayerIndex = 0 Then
        Call PlayerMsg(Index, "You have not requested to trade with anyone.", PINK)
        Exit Sub
    End If

    If Player(PlayerIndex).TradePlayer <> Index Then
        Call PlayerMsg(Index, "Trade failed.", PINK)
        Exit Sub
    End If

    If GetPlayerMap(Index) <> GetPlayerMap(PlayerIndex) Then
        Call PlayerMsg(Index, "You must be on the same map to trade with " & GetPlayerName(PlayerIndex) & "!", PINK)
        Exit Sub
    End If

    Call PlayerMsg(Index, "You are trading with " & GetPlayerName(PlayerIndex) & "!", PINK)
    Call PlayerMsg(PlayerIndex, GetPlayerName(Index) & " accepted your trade request!", PINK)

    Call SendDataTo(Index, POut.PrepareTrade & END_CHAR)
    Call SendDataTo(PlayerIndex, POut.PrepareTrade & END_CHAR)

    For I = 1 To MAX_PLAYER_TRADES
        Player(Index).Trading(I).InvNum = 0
        Player(Index).Trading(I).InvName = vbNullString

        Player(PlayerIndex).Trading(I).InvNum = 0
        Player(PlayerIndex).Trading(I).InvName = vbNullString
    Next I

    Player(Index).InTrade = True
    Player(Index).TradeItemMax = 0
    Player(Index).TradeItemMax2 = 0

    Player(PlayerIndex).InTrade = True
    Player(PlayerIndex).TradeItemMax = 0
    Player(PlayerIndex).TradeItemMax2 = 0
End Sub

Public Sub Packet_QuitTrade(ByVal Index As Long)
    Dim PlayerIndex As Long

    PlayerIndex = Player(Index).TradePlayer

    If PlayerIndex = 0 Then
        Call PlayerMsg(Index, "You have not requested to trade with anyone.", PINK)
        Exit Sub
    End If

    Call PlayerMsg(Index, "Stopped trading.", PINK)
    Call PlayerMsg(PlayerIndex, GetPlayerName(Index) & " stopped trading with you!", PINK)

    Player(Index).TradeOk = 0
    Player(Index).TradePlayer = 0
    Player(Index).InTrade = False

    Player(PlayerIndex).TradeOk = 0
    Player(PlayerIndex).TradePlayer = 0
    Player(PlayerIndex).InTrade = False

    Call SendDataTo(Index, POut.QuitTrade & END_CHAR)
    Call SendDataTo(PlayerIndex, POut.QuitTrade & END_CHAR)
End Sub

Public Sub Packet_DenyTrade(ByVal Index As Long)
    Dim PlayerIndex As Long

    PlayerIndex = Player(Index).TradePlayer

    If PlayerIndex = 0 Then
        Call PlayerMsg(Index, "You have not requested to trade with anyone.", PINK)
        Exit Sub
    End If

    Call PlayerMsg(Index, "Declined trade request.", PINK)
    Call PlayerMsg(PlayerIndex, GetPlayerName(Index) & " declined your request.", PINK)

    Player(Index).TradePlayer = 0
    Player(Index).InTrade = False

    Player(PlayerIndex).TradePlayer = 0
    Player(PlayerIndex).InTrade = False
End Sub

Public Sub Packet_UpdateTradeInventory(ByVal Index As Long, ByVal TradeIndex As Long, ByVal InvNum As Long, ByVal ItemName As String, ByVal ItemAmount As Long)
    Player(Index).Trading(TradeIndex).InvNum = InvNum
    Player(Index).Trading(TradeIndex).InvName = Trim$(ItemName)
    Player(Player(Index).TradePlayer).Trading(TradeIndex).InvNum = InvNum
    Player(Player(Index).TradePlayer).Trading(TradeIndex).InvName = Trim$(ItemName)

    If Item(GetPlayerInvItemNum(Index, InvNum)).Stackable = 1 Then
        Player(Index).Trading(TradeIndex).InvAmt = ItemAmount
        Player(Player(Index).TradePlayer).Trading(TradeIndex).InvAmt = ItemAmount
    Else
        Player(Index).Trading(TradeIndex).InvAmt = 0
        Player(Player(Index).TradePlayer).Trading(TradeIndex).InvNum = 0
    End If

    If Player(Index).Trading(TradeIndex).InvNum = 0 Then
        Player(Index).TradeItemMax = Player(Index).TradeItemMax - 1
        Player(Index).TradeOk = 0
        Player(TradeIndex).TradeOk = 0

        Call SendDataTo(Index, POut.Trading & SEP_CHAR & 0 & END_CHAR)
        Call SendDataTo(TradeIndex, POut.Trading & SEP_CHAR & 0 & END_CHAR)
    Else
        Player(Index).TradeItemMax = Player(Index).TradeItemMax + 1
    End If

    Call SendDataTo(Player(Index).TradePlayer, POut.UpdateTradeItem & SEP_CHAR & TradeIndex & SEP_CHAR & Player(Index).Trading(TradeIndex).InvNum & SEP_CHAR & Player(Index).Trading(TradeIndex).InvName & SEP_CHAR & Player(Index).Trading(TradeIndex).InvAmt & END_CHAR)
End Sub

Public Sub Packet_SwapItems(ByVal Index As Long)
    Dim TradeIndex As Long
    Dim I As Long
    Dim X As Long

    TradeIndex = Player(Index).TradePlayer

    If Player(Index).TradeOk = 0 Then
        Player(Index).TradeOk = 1
        Call SendDataTo(TradeIndex, POut.Trading & SEP_CHAR & 1 & END_CHAR)
    ElseIf Player(Index).TradeOk = 1 Then
        Player(Index).TradeOk = 0
        Call SendDataTo(TradeIndex, POut.Trading & SEP_CHAR & 0 & END_CHAR)
    End If

    If Player(Index).TradeOk = 1 Then
        If Player(TradeIndex).TradeOk = 1 Then
            Player(Index).TradeItemMax2 = 0
            Player(TradeIndex).TradeItemMax2 = 0
    
            For I = 1 To MAX_INV
                If Player(Index).TradeItemMax = Player(Index).TradeItemMax2 Then
                    Exit For
                End If

                If GetPlayerInvItemNum(TradeIndex, I) < 1 Then
                    Player(Index).TradeItemMax2 = Player(Index).TradeItemMax2 + 1
                End If
            Next I
    
            For I = 1 To MAX_INV
                If Player(TradeIndex).TradeItemMax = Player(TradeIndex).TradeItemMax2 Then
                    Exit For
                End If

                If GetPlayerInvItemNum(Index, I) < 1 Then
                    Player(TradeIndex).TradeItemMax2 = Player(TradeIndex).TradeItemMax2 + 1
                End If
            Next I
    
            If Player(Index).TradeItemMax2 = Player(Index).TradeItemMax And Player(TradeIndex).TradeItemMax2 = Player(TradeIndex).TradeItemMax Then
                For I = 1 To MAX_PLAYER_TRADES
                    For X = 1 To MAX_INV
                        If GetPlayerInvItemNum(TradeIndex, X) < 1 Then
                            If Player(Index).Trading(I).InvNum > 0 Then
                                Call GiveItem(TradeIndex, GetPlayerInvItemNum(Index, Player(Index).Trading(I).InvNum), Player(Index).Trading(I).InvAmt)
                                Call TakeItem(Index, GetPlayerInvItemNum(Index, Player(Index).Trading(I).InvNum), Player(Index).Trading(I).InvAmt)
                                Exit For
                            End If
                        End If
                    Next X
                Next I
    
                For I = 1 To MAX_PLAYER_TRADES
                    For X = 1 To MAX_INV
                        If GetPlayerInvItemNum(Index, X) < 1 Then
                            If Player(TradeIndex).Trading(I).InvNum > 0 Then
                                Call GiveItem(Index, GetPlayerInvItemNum(TradeIndex, Player(TradeIndex).Trading(I).InvNum), Player(Index).Trading(I).InvAmt)
                                Call TakeItem(TradeIndex, GetPlayerInvItemNum(TradeIndex, Player(TradeIndex).Trading(I).InvNum), Player(Index).Trading(I).InvAmt)
                                Exit For
                            End If
                        End If
                    Next X
                Next I

                Call PlayerMsg(Index, "The trade was successful!", BRIGHTGREEN)
                Call PlayerMsg(TradeIndex, "The trade was successful!", BRIGHTGREEN)

                Call SendInventory(Index)
                Call SendInventory(TradeIndex)
            Else
                If Player(Index).TradeItemMax2 < Player(Index).TradeItemMax Then
                    Call PlayerMsg(Index, "Your inventory is full!", BRIGHTRED)
                    Call PlayerMsg(TradeIndex, GetPlayerName(Index) & "'s inventory is full!", BRIGHTRED)
                End If
                        
                If Player(TradeIndex).TradeItemMax2 < Player(TradeIndex).TradeItemMax Then
                    Call PlayerMsg(TradeIndex, "Your inventory is full!", BRIGHTRED)
                    Call PlayerMsg(Index, GetPlayerName(TradeIndex) & "'s inventory is full!", BRIGHTRED)
                End If
            End If
    
            Player(Index).TradePlayer = 0
            Player(Index).InTrade = False
            Player(Index).TradeOk = 0

            Player(TradeIndex).TradePlayer = 0
            Player(TradeIndex).InTrade = False
            Player(TradeIndex).TradeOk = 0

            Call SendDataTo(Index, POut.QuitTrade & END_CHAR)
            Call SendDataTo(TradeIndex, POut.QuitTrade & END_CHAR)
        End If
    End If
End Sub

Public Sub Packet_Spells(ByVal Index As Long)
    Call SendPlayerSpells(Index)
End Sub

Public Sub Packet_HotScript(ByVal Index As Long, ByVal ScriptID As Long)
    If SCRIPTING = 1 Then
        MyScript.ExecuteStatement "Scripts\Main.txt", "HotScript " & Index & "," & ScriptID
    End If
End Sub

Public Sub Packet_ScriptTile(ByVal Index As Long, ByVal TileNum As Long)
    Call SendDataTo(Index, POut.ScriptTile & SEP_CHAR & GetVar(FILE_TILESINI, "Names", "Tile" & TileNum) & END_CHAR)
End Sub

Public Sub Packet_Cast(ByVal Index As Long, ByVal SpellNum As Long)
    Call CastSpell(Index, SpellNum)
End Sub

Public Sub Packet_Refresh(ByVal Index As Long)
    Call SendDataToMap(GetPlayerMap(Index), POut.PlayerXY & SEP_CHAR & Index & SEP_CHAR & GetPlayerX(Index) & SEP_CHAR & GetPlayerY(Index) & END_CHAR)
End Sub

Public Sub Packet_BuySprite(ByVal Index As Long)
    Dim I As Long

    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type <> TILE_TYPE_SPRITE_CHANGE Then
        Call PlayerMsg(Index, "You need to be on a sprite tile to buy it!", BRIGHTRED)
        Exit Sub
    End If

    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2 = 0 Then
        Call SetPlayerSprite(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1)
        Call SendDataToMap(GetPlayerMap(Index), POut.CheckSprite & SEP_CHAR & Index & SEP_CHAR & GetPlayerSprite(Index) & END_CHAR)
        Exit Sub
    End If

    For I = 1 To MAX_INV
        If GetPlayerInvItemNum(Index, I) = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2 Then
            If Item(GetPlayerInvItemNum(Index, I)).Type = ITEM_TYPE_CURRENCY Then
                If GetPlayerInvItemValue(Index, I) >= Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data3 Then
                    Call SetPlayerInvItemValue(Index, I, GetPlayerInvItemValue(Index, I) - Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data3)

                    If GetPlayerInvItemValue(Index, I) = 0 Then
                        Call SetPlayerInvItemNum(Index, I, 0)
                    End If

                    Call PlayerMsg(Index, "You have bought a new sprite!", BRIGHTGREEN)
                    Call SetPlayerSprite(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1)
                    Call SendDataToMap(GetPlayerMap(Index), POut.CheckSprite & SEP_CHAR & Index & SEP_CHAR & GetPlayerSprite(Index) & END_CHAR)
                    Call SendInventory(Index)
                End If
            Else
                If GetPlayerWeaponSlot(Index) <> I And GetPlayerArmorSlot(Index) <> I And GetPlayerShieldSlot(Index) <> I And GetPlayerHelmetSlot(Index) <> I And GetPlayerLegsSlot(Index) <> I And GetPlayerRingSlot(Index) <> I And GetPlayerNecklaceSlot(Index) <> I Then
                    Call SetPlayerInvItemNum(Index, I, 0)
                    Call PlayerMsg(Index, "You have bought a new sprite!", BRIGHTGREEN)
                    Call SetPlayerSprite(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1)
                    Call SendDataToMap(GetPlayerMap(Index), POut.CheckSprite & SEP_CHAR & Index & SEP_CHAR & GetPlayerSprite(Index) & END_CHAR)
                    Call SendInventory(Index)
                End If
            End If

            If GetPlayerWeaponSlot(Index) <> I And GetPlayerArmorSlot(Index) <> I And GetPlayerShieldSlot(Index) <> I And GetPlayerHelmetSlot(Index) <> I And GetPlayerLegsSlot(Index) <> I And GetPlayerRingSlot(Index) <> I And GetPlayerNecklaceSlot(Index) <> I Then
                Exit Sub
            End If
        End If
    Next I

    Call PlayerMsg(Index, "You don't have enough to buy this sprite!", BRIGHTRED)
End Sub

Public Sub Packet_CheckCommands(ByVal Index As Long, ByVal Command As String)
    If SCRIPTING = 1 Then
        MyScript.ExecuteStatement "Scripts\Main.txt", "Commands " & Index & "," & Trim$(Command)
    Else
        Call PlayerMsg(Index, "That is not a valid command!", BRIGHTRED)
    End If
End Sub

Public Sub Packet_RequestEditArrow(ByVal Index As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    Call SendDataTo(Index, POut.ArrowEditor & END_CHAR)
End Sub

Public Sub Packet_EditArrow(ByVal Index As Long, ByVal ArrowNum As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    If ArrowNum < 0 Or ArrowNum > MAX_ARROWS Then
        Call HackingAttempt(Index, "Invalid Arrow Index")
        Exit Sub
    End If

    Call SendEditArrowTo(Index, ArrowNum)

    Call AddLog(GetPlayerName(Index) & " editing arrow #" & ArrowNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_SaveArrow(ByVal Index As Long, ByVal ArrowNum As Long, ByVal Name As String, ByVal Pic As Long, ByVal Range As Long, ByVal Amount As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    If ArrowNum < 0 Or ArrowNum > MAX_ITEMS Then
        Call HackingAttempt(Index, "Invalid Arrow Index")
        Exit Sub
    End If

    Arrows(ArrowNum).Name = Name
    Arrows(ArrowNum).Pic = Pic
    Arrows(ArrowNum).Range = Range
    Arrows(ArrowNum).Amount = Amount

    Call SendUpdateArrowToAll(ArrowNum)
    Call SaveArrow(ArrowNum)

    Call AddLog(GetPlayerName(Index) & " saved arrow #" & ArrowNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_CheckArrows(ByVal Index As Long, ByVal ArrowNum As Long)
    Call SendDataToMap(GetPlayerMap(Index), POut.CheckArrows & SEP_CHAR & Index & SEP_CHAR & Arrows(ArrowNum).Pic & END_CHAR)
End Sub

Public Sub Packet_RequestEditEmoticon(ByVal Index As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    Call SendDataTo(Index, POut.EmoticonEditor & END_CHAR)
End Sub

Public Sub Packet_RequestEditElement(ByVal Index As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    Call SendDataTo(Index, POut.ElementEditor & END_CHAR)
End Sub

Public Sub Packet_EditEmoticon(ByVal Index As Long, ByVal EmoteNum As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    If EmoteNum < 0 Or EmoteNum > MAX_EMOTICONS Then
        Call HackingAttempt(Index, "Invalid Emoticon Index")
        Exit Sub
    End If

    Call SendEditEmoticonTo(Index, EmoteNum)

    Call AddLog(GetPlayerName(Index) & " editing emoticon #" & EmoteNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_EditElement(ByVal Index As Long, ByVal ElementNum As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    If ElementNum < 0 Or ElementNum > MAX_ELEMENTS Then
        Call HackingAttempt(Index, "Invalid Emoticon Index")
        Exit Sub
    End If

    Call SendEditElementTo(Index, ElementNum)

    Call AddLog(GetPlayerName(Index) & " editing element #" & ElementNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_SaveEmoticon(ByVal Index As Long, ByVal EmoteNum As Long, ByVal Command As String, ByVal Pic As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    If EmoteNum < 0 Or EmoteNum > MAX_EMOTICONS Then
        Call HackingAttempt(Index, "Invalid Emoticon Index")
        Exit Sub
    End If

    Emoticons(EmoteNum).Command = Command
    Emoticons(EmoteNum).Pic = Pic

    Call SendUpdateEmoticonToAll(EmoteNum)
    Call SaveEmoticon(EmoteNum)

    Call AddLog(GetPlayerName(Index) & " saved emoticon #" & EmoteNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_SaveElement(ByVal Index As Long, ByVal ElementNum As Long, ByVal Name As String, ByVal Strong As Long, ByVal Weak As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Admin Cloning")
        Exit Sub
    End If

    If ElementNum < 0 Or ElementNum > MAX_ELEMENTS Then
        Call HackingAttempt(Index, "Invalid Element Index")
        Exit Sub
    End If

    Element(ElementNum).Name = Name
    Element(ElementNum).Strong = Strong
    Element(ElementNum).Weak = Weak

    Call SendUpdateElementToAll(ElementNum)
    Call SaveElement(ElementNum)

    Call AddLog(GetPlayerName(Index) & " saved element #" & ElementNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_CheckEmoticon(ByVal Index As Long, ByVal EmoteNum As Long)
    Call SendDataToMap(GetPlayerMap(Index), POut.CheckEmoticons & SEP_CHAR & Index & SEP_CHAR & Emoticons(EmoteNum).Pic & END_CHAR)
End Sub

Public Sub Packet_MapReport(ByVal Index As Long)
    Dim packet As String
    Dim I As Long

    If GetPlayerAccess(Index) < ADMIN_MONITER Then
        Call HackingAttempt(Index, "Packet Modification")
        Exit Sub
    End If

    packet = POut.MapReport & SEP_CHAR

    For I = 1 To MAX_MAPS
        packet = packet & Map(I).Name & SEP_CHAR
    Next I

    packet = packet & END_CHAR

    Call SendDataTo(Index, packet)
End Sub

Public Sub Packet_GMTime(ByVal Index As Long, ByVal SetTime As Long)
    If GetPlayerAccess(Index) < ADMIN_MONITER Then
        Call HackingAttempt(Index, "Packet Modification")
        Exit Sub
    End If

    ' Check if it's a valid time variable. [Mellowz]
    If SetTime < 0 Or SetTime > 1 Then
        Call PlayerMsg(Index, "Please enter a valid time variable.", BRIGHTRED)
        Exit Sub
    End If

    GameTime = SetTime

    Call SendTimeToAll
End Sub

Public Sub Packet_Weather(ByVal Index As Long, ByVal WeatherNum As Long)
    If GetPlayerAccess(Index) < ADMIN_MONITER Then
        Call HackingAttempt(Index, "Packet Modification")
        Exit Sub
    End If

    ' Check if it's a valid weather variable. [Mellowz]
    If WeatherNum < 0 Or WeatherNum > 3 Then
        Call PlayerMsg(Index, "Please enter a valid weather variable.", BRIGHTRED)
        Exit Sub
    End If

    WeatherType = WeatherNum

    Call SendWeatherToAll
End Sub

Public Sub Packet_WarpTo(ByVal Index As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
    If GetPlayerAccess(Index) < ADMIN_MONITER Then
        Call HackingAttempt(Index, "Packet Modification")
        Exit Sub
    End If
    
    If X < 0 Or X > MAX_MAPX Then
        Call PlayerMsg(Index, "Please enter a valid X coordinate.", BRIGHTRED)
        Exit Sub
    End If

    If Y < 0 Or Y > MAX_MAPY Then
        Call PlayerMsg(Index, "Please enter a valid Y coordinate.", BRIGHTRED)
        Exit Sub
    End If

    Call PlayerWarp(Index, MapNum, X, Y)
End Sub

Public Sub Packet_LocalWarp(ByVal Index As Long, ByVal X As Long, ByVal Y As Long)
    If GetPlayerAccess(Index) < ADMIN_MONITER Then
        Call HackingAttempt(Index, "Packet Modification")
        Exit Sub
    End If
    
    If X < 0 Or X > MAX_MAPX Then
        Call PlayerMsg(Index, "Please enter a valid X coordinate.", BRIGHTRED)
        Exit Sub
    End If

    If Y < 0 Or Y > MAX_MAPY Then
        Call PlayerMsg(Index, "Please enter a valid Y coordinate.", BRIGHTRED)
        Exit Sub
    End If

    Call SetPlayerX(Index, X)
    Call SetPlayerY(Index, Y)

    Call SendPlayerXY(Index)
End Sub

Public Sub Packet_ArrowHit(ByVal Index As Long, ByVal TargetType As Long, ByVal PlayerIndex As Long, ByVal X As Long, ByVal Y As Long)
    Dim Damage As Long
    
    If TargetType = TARGET_TYPE_PLAYER Then
        If PlayerIndex <> Index Then
            If CanAttackPlayerWithArrow(Index, PlayerIndex) Then
                Player(Index).Target = PlayerIndex
                Player(Index).TargetType = TARGET_TYPE_PLAYER
                If Not CanPlayerBlockHit(PlayerIndex) Then
                    If Not CanPlayerCriticalHit(Index) Then
                        Damage = GetPlayerDamage(Index) - GetPlayerProtection(PlayerIndex)
                        Call SendDataToMap(GetPlayerMap(Index), POut.Sound & SEP_CHAR & "attack" & END_CHAR)
                    Else
                        TargetType = GetPlayerDamage(Index)
                        Damage = TargetType + Int(Rnd * Int(TargetType / 2)) + 1 - GetPlayerProtection(PlayerIndex)

                        Call BattleMsg(Index, "You feel a surge of energy upon shooting!", BRIGHTCYAN, 0)
                        Call BattleMsg(PlayerIndex, GetPlayerName(Index) & " shoots With amazing accuracy!", BRIGHTCYAN, 1)

                        Call SendDataToMap(GetPlayerMap(Index), POut.Sound & SEP_CHAR & "critical" & END_CHAR)
                    End If

                    If Damage > 0 Then
                        If SCRIPTING = 1 Then
                            MyScript.ExecuteStatement "Scripts\Main.txt", "OnArrowHit " & Index & "," & Damage
                        Else
                            Call AttackPlayer(Index, PlayerIndex, Damage)
                        End If
                    Else
                        If SCRIPTING = 1 Then
                            MyScript.ExecuteStatement "Scripts\Main.txt", "OnArrowHit " & Index & "," & 0
                        End If
                        Call BattleMsg(Index, "Your attack does nothing.", BRIGHTRED, 0)
                        Call BattleMsg(PlayerIndex, GetPlayerName(Index) & "'s attack did nothing.", BRIGHTRED, 1)

                        Call SendDataToMap(GetPlayerMap(Index), POut.Sound & SEP_CHAR & "miss" & END_CHAR)
                    End If
                Else
                
                    If SCRIPTING = 1 Then
                        MyScript.ExecuteStatement "Scripts\Main.txt", "OnArrowHit " & Index & "," & 0
                    End If
                    Call BattleMsg(Index, GetPlayerName(PlayerIndex) & " blocked your hit!", BRIGHTCYAN, 0)
                    Call BattleMsg(PlayerIndex, "You blocked " & GetPlayerName(Index) & "'s hit!", BRIGHTCYAN, 1)

                    Call SendDataToMap(GetPlayerMap(Index), POut.Sound & SEP_CHAR & "miss" & END_CHAR)
                End If

                Exit Sub
            End If
        End If
    ElseIf TargetType = TARGET_TYPE_NPC Then
        If CanAttackNpcWithArrow(Index, PlayerIndex) Then
        Player(Index).TargetType = TARGET_TYPE_NPC
        Player(Index).TargetNPC = PlayerIndex
            If Not CanPlayerCriticalHit(Index) Then
                Damage = GetPlayerDamage(Index) - Int(NPC(MapNPC(GetPlayerMap(Index), PlayerIndex).num).DEF / 2)
                Call SendDataToMap(GetPlayerMap(Index), POut.Sound & SEP_CHAR & "attack" & END_CHAR)
            Else
                TargetType = GetPlayerDamage(Index)
                Damage = TargetType + Int(Rnd * Int(TargetType / 2)) + 1 - Int(NPC(MapNPC(GetPlayerMap(Index), PlayerIndex).num).DEF / 2)

                Call BattleMsg(Index, "You feel a surge of energy upon shooting!", BRIGHTCYAN, 0)

                Call SendDataToMap(GetPlayerMap(Index), POut.Sound & SEP_CHAR & "critical" & END_CHAR)
            End If

            If Damage > 0 Then
                If SCRIPTING = 1 Then
                    MyScript.ExecuteStatement "Scripts\Main.txt", "OnArrowHit " & Index & "," & Damage
                Else
                    Call AttackNpc(Index, PlayerIndex, Damage)
                    Call SendDataTo(Index, POut.DrawPlayerDamage & SEP_CHAR & Damage & SEP_CHAR & PlayerIndex & END_CHAR)
                End If
            Else
                If SCRIPTING = 1 Then
                    MyScript.ExecuteStatement "Scripts\Main.txt", "OnArrowHit " & Index & "," & Damage
                End If
                Call BattleMsg(Index, "Your attack does nothing.", BRIGHTRED, 0)

                Call SendDataTo(Index, POut.DrawPlayerDamage & SEP_CHAR & Damage & SEP_CHAR & PlayerIndex & END_CHAR)
                Call SendDataToMap(GetPlayerMap(Index), POut.Sound & SEP_CHAR & "miss" & END_CHAR)
            End If

            Exit Sub
        End If
    End If
End Sub

Public Sub Packet_BankDeposit(ByVal Index As Long, ByVal InvNum As Long, ByVal Amount As Long)
    Dim BankSlot As Long
    Dim ItemNum As Long

    ItemNum = GetPlayerInvItemNum(Index, InvNum)

    BankSlot = FindOpenBankSlot(Index, ItemNum)
    If BankSlot = 0 Then
        Call SendDataTo(Index, POut.BankMessage & SEP_CHAR & "Bank full!" & END_CHAR)
        Exit Sub
    End If

    If Amount < 0 Or Amount > GetPlayerInvItemValue(Index, InvNum) Then
        Call SendDataTo(Index, POut.BankMessage & SEP_CHAR & "You can't deposit more than you have!" & END_CHAR)
        Exit Sub
    End If

    If GetPlayerWeaponSlot(Index) = ItemNum Or GetPlayerArmorSlot(Index) = ItemNum Or GetPlayerShieldSlot(Index) = ItemNum Or GetPlayerHelmetSlot(Index) = ItemNum Or GetPlayerLegsSlot(Index) = ItemNum Or GetPlayerRingSlot(Index) = ItemNum Or GetPlayerNecklaceSlot(Index) = ItemNum Then
        Call SendDataTo(Index, POut.BankMessage & SEP_CHAR & "You can't deposit worn equipment!" & END_CHAR)
        Exit Sub
    End If

    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then
        If Amount = 0 Then
            Call SendDataTo(Index, POut.BankMessage & SEP_CHAR & "You must deposit more than 0!" & END_CHAR)
            Exit Sub
        End If
    End If

    Call TakeItem(Index, ItemNum, Amount)
    Call GiveBankItem(Index, ItemNum, Amount, BankSlot)

    Call SendBank(Index)
End Sub

Public Sub Packet_BankWithdraw(ByVal Index As Long, ByVal BankInvNum As Long, ByVal Amount As Long)
    Dim BankItemNum As Long
    Dim BankInvSlot As Long

    BankItemNum = GetPlayerBankItemNum(Index, BankInvNum)

    BankInvSlot = FindOpenInvSlot(Index, BankItemNum)
    If BankInvSlot = 0 Then
        Call SendDataTo(Index, POut.BankMessage & SEP_CHAR & "Inventory full!" & END_CHAR)
        Exit Sub
    End If

    If Amount < 0 Or Amount > GetPlayerBankItemValue(Index, BankInvNum) Then
        Call SendDataTo(Index, POut.BankMessage & SEP_CHAR & "You can't withdraw more than you have!" & END_CHAR)
        Exit Sub
    End If

    If Item(BankItemNum).Type = ITEM_TYPE_CURRENCY Or Item(BankItemNum).Stackable = 1 Then
        If Amount = 0 Then
            Call SendDataTo(Index, POut.BankMessage & SEP_CHAR & "You must withdraw more than 0!" & END_CHAR)
            Exit Sub
        End If
    End If

    Call TakeBankItem(Index, BankItemNum, Amount)
    Call GiveItem(Index, BankItemNum, Amount)

    Call SendBank(Index)
End Sub

Public Sub Packet_ReloadScripts(ByVal Index As Long)
    If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(Index, "Packet Modification")
        Exit Sub
    End If

    Set MyScript = Nothing
    Set clsScriptCommands = Nothing

    Set MyScript = New clsSadScript
    Set clsScriptCommands = New clsCommands

    MyScript.ReadInCode App.Path & "\Scripts\Main.txt", "Scripts\Main.txt", MyScript.SControl
    MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True

    MyScript.ExecuteStatement "Scripts\Main.txt", "OnScriptReload"

    Call TextAdd(frmServer.txtText(0), "Scripts reloaded.", True)
    Call AdminMsg("Scripts reloaded by " & GetPlayerName(Index) & ".", WHITE)
End Sub

Public Sub Packet_CustomMenuClick(ByVal Index As Long, ByVal MenuIndex As Long, ByVal ClickIndex As Long, ByVal CustomTitle As String, ByVal MenuType As Long, ByVal CustomMsg As String)
    Player(Index).CustomTitle = CustomTitle
    Player(Index).CustomMsg = CustomMsg

    If SCRIPTING = 1 Then
        MyScript.ExecuteStatement "Scripts\Main.txt", "menuscripts " & MenuIndex & "," & ClickIndex & "," & MenuType
    End If
End Sub

Public Sub Packet_CustomBoxReturnMsg(ByVal Index As Long, ByVal CustomMsg As String)
    Player(Index).CustomMsg = CustomMsg
End Sub

Public Sub Packet_PartyCreate(ByVal Index As Long, ByVal PartyName As String)
    Dim PartyID As Long

    ' Check if the player is in a party.
    If Player(Index).InParty Then
        Call PlayerMsg(Index, "You are already in a party.", WHITE)
        Exit Sub
    End If

    ' Check if the party name is long enough.
    If LenB(PartyName) = 0 Then
        Call PlayerMsg(Index, "The party name was too short.", WHITE)
        Exit Sub
    End If

    ' Find an available party slot.
    PartyID = Party_FindOpenPartySlot()

    ' If there's no available slot, then we need to increase the size of the array.
    If PartyID = 0 Then
        ReDim Preserve Party(1 To UBound(Party) + 1) As NewPartyRec
        PartyID = UBound(Party)
    End If

    ' Create the party server-side.
    Party(PartyID).InUse = True

    Party(PartyID).PartyName = PartyName
    Party(PartyID).PartyLeader = Index
    
    Party(PartyID).Members(1) = Index

    ' Update the players party index.
    Player(Index).InParty = True
    Player(Index).PartyIndex = PartyID

    ' The party has been created.
    Call PlayerMsg(Index, "The party '" & PartyName & "' has been successfully created.", WHITE)
End Sub

Public Sub Packet_PartyDisband(ByVal Index As Long)
    Dim PartyID As Long
    Dim I As Long

    ' Check if the player is in a party.
    If Not Player(Index).InParty Then
        Call PlayerMsg(Index, "You are not in a party.", WHITE)
        Exit Sub
    End If

    ' Check if the player is the party leader.
    If Not Party(Player(Index).PartyIndex).PartyLeader = Index Then
        Call PlayerMsg(Index, "You are not the party leader.", WHITE)
        Exit Sub
    End If
    
    Call PartyMsg(Player(Index).PartyIndex, "The party has been disband.")

    ' Get the party index.
    PartyID = Player(Index).PartyIndex

    ' Disband the party.
    Party(PartyID).InUse = False

    Party(PartyID).PartyName = vbNullString
    Party(PartyID).PartyLeader = 0

    ' Kick out any other players in the party and reset
    ' their party data so they can join future parties.
    For I = 1 To 8
        If Party(PartyID).Members(I) > 0 Then
            Player(Party(PartyID).Members(I)).InParty = False
            Player(Party(PartyID).Members(I)).PartyIndex = 0
            Party(PartyID).Members(I) = 0
        End If
    Next I
End Sub

Public Sub Packet_PartyInvite(ByVal Index As Long, ByVal PlayerName As String)
    Dim PlayerIndex As Long

    ' Check if the player who invited is in a party.
    If Not Player(Index).InParty Then
        Call PlayerMsg(Index, "You are not in a party.", WHITE)
        Exit Sub
    End If

    ' Check if the player who invited is a party leader.
    If Not Party(Player(Index).PartyIndex).PartyLeader = Index Then
        Call PlayerMsg(Index, "You are not the party leader.", WHITE)
        Exit Sub
    End If

    ' Find the player's index.
    PlayerIndex = FindPlayer(PlayerName)

    ' Check if a player was found.
    If PlayerIndex = 0 Then
        Call PlayerMsg(Index, PlayerName & " is offline or doesn't exist.", WHITE)
        Exit Sub
    End If

    ' Check if they are already in a party.
    If Player(PlayerIndex).InParty Then
        Call PlayerMsg(Index, PlayerName & " is already in a party.", WHITE)
        Exit Sub
    End If

    Player(PlayerIndex).PartyIndex = Player(Index).PartyIndex

    ' Send to the party leader.
    Call PlayerMsg(Index, PlayerName & " has been invited to join the party.", WHITE)

    ' Send to the requested index.
    Call PlayerMsg(PlayerIndex, "Would you like to join " & GetPlayerName(Index) & "'s party?", WHITE)
    Call PlayerMsg(PlayerIndex, "Respond with either '/accept' or '/decline' to trigger the party.", WHITE)
End Sub

Public Sub Packet_PartyInviteAccept(ByVal Index As Long)
    Dim MemberID As Long

    ' Check if there's an invitation.
    If Player(Index).PartyIndex = 0 Then
        Call PlayerMsg(Index, "You haven't been invited to any party.", WHITE)
        Exit Sub
    End If

    ' Check if they are already in a party.
    If Player(Index).InParty Then
        Call PlayerMsg(Index, "You are already in a party.", WHITE)
        Exit Sub
    End If

    ' Get an available member ID.
    MemberID = Party_FindOpenPlayerSlot(Player(Index).PartyIndex)
    
    ' Check if the party is full.
    If MemberID = 0 Then
        Call PlayerMsg(Index, "The party is full.", WHITE)
        Player(Index).PartyIndex = 0
        Exit Sub
    End If

    ' Add them to the party.
    Player(Index).InParty = True

    Party(Player(Index).PartyIndex).Members(MemberID) = Index

    ' Inform everyone currently in the party.
    Call PartyMsg(Player(Index).PartyIndex, GetPlayerName(Index) & " has joined the party.")
End Sub

Public Sub Packet_PartyInviteDecline(ByVal Index As Long)
    ' Check if there's an invitation.
    If Player(Index).PartyIndex = 0 Then
        Call PlayerMsg(Index, "You haven't been invited to any party.", WHITE)
        Exit Sub
    End If

    ' Check if they are already in a party.
    If Player(Index).InParty Then
        Call PlayerMsg(Index, "You are already in a party.", WHITE)
        Exit Sub
    End If

    ' Inform everyone currently in the party.
    Call PartyMsg(Player(Index).PartyIndex, GetPlayerName(Index) & " has declined the party invitation.")

    ' Clear the party data from the player.
    Player(Index).PartyIndex = 0
End Sub

Public Sub Packet_PartyLeave(ByVal Index As Long)
    Dim PlayerSlot As Long

    ' Check if they are in a party.
    If Not Player(Index).InParty Then
        Call PlayerMsg(Index, "You are not in a party.", WHITE)
        Exit Sub
    End If

    ' Check who we're dealing with: the leader or a member.
    If Player(Index).PartyIndex = Party(Player(Index).PartyIndex).PartyLeader Then
        ' The leader.

        ' For now, to debug the party, we just remove the party.
        Call Packet_PartyDisband(Index)

        ' Party_DisbandParty() resets all player and party data.
    Else
        ' The member.

        ' Find the players player slot in the party array.
        PlayerSlot = Party_GetPlayerSlot(Index)
        
        ' We shouldn't have to check if a player is found, because
        ' their index will ALWAYS be found no matter what.

        ' Inform the party.
        Call PartyMsg(Player(Index).PartyIndex, GetPlayerName(Index) & " has left the party.")

        ' Reset the players party data.
        Player(Index).InParty = False
        Player(Index).PartyIndex = 0

        ' Reset the players party data.
        Party(Player(Index).PartyIndex).Members(PlayerSlot) = 0
    End If
End Sub

Public Sub Packet_PartyChangeLeader(ByVal Index As Long, ByVal PlayerName As String)
    Dim PlayerIndex As Long

    ' Check if they are in a party.
    If Not Player(Index).InParty Then
        Call PlayerMsg(Index, "You are not in a party.", WHITE)
        Exit Sub
    End If

    ' Check if they are the party leader.
    If Not Party(Player(Index).PartyIndex).PartyLeader = Index Then
        Call PlayerMsg(Index, "You are not the party leader.", WHITE)
        Exit Sub
    End If

    ' Find the player's index.
    PlayerIndex = FindPlayer(PlayerName)

    ' Check if a player was found.
    If PlayerIndex = 0 Then
        Call PlayerMsg(Index, PlayerName & " is offline or doesn't exist.", WHITE)
        Exit Sub
    End If

    ' Check if they are already in a party.
    If Not Player(PlayerIndex).InParty Then
        Call PlayerMsg(Index, PlayerName & " is not in a party.", WHITE)
        Exit Sub
    End If

    ' Check if the player is in the same party as the party leader.
    If Not Player(Index).PartyIndex = Player(PlayerIndex).PartyIndex Then
        Call PlayerMsg(Index, PlayerName & " is not in your party.", WHITE)
        Exit Sub
    End If

    ' Set the new party leader.
    Party(Player(Index).PartyIndex).PartyLeader = PlayerIndex

    ' Inform the group.
    Call PartyMsg(Player(Index).PartyIndex, GetPlayerName(Index) & " has given " & GetPlayerName(PlayerIndex) & " party leadership.")
End Sub

