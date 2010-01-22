Attribute VB_Name = "modGameLogic"
Option Explicit

Public Function TwipsToPixels(lngTwips As Long, _
        lngDirection As Long) As Long

    ' Handle to device
    Dim lngDC As Long
    Dim lngPixelsPerInch As Long
    Const nTwipsPerInch = 1440
    lngDC = GetDC(0)

    If (lngDirection = 0) Then       'Horizontal
        lngPixelsPerInch = GetDeviceCaps(lngDC, 88)
    Else                            'Vertical
        lngPixelsPerInch = GetDeviceCaps(lngDC, 90)
    End If
    lngDC = ReleaseDC(0, lngDC)
    TwipsToPixels = (lngTwips / nTwipsPerInch) * lngPixelsPerInch

End Function

Public Function PixelsToTwips(lngTwips As Long, _
        lngDirection As Long) As Long

    ' Handle to device
    Dim lngDC As Long
    Dim lngPixelsPerInch As Long
    Const nTwipsPerInch = 1440
    lngDC = GetDC(0)

    If (lngDirection = 0) Then       'Horizontal
        lngPixelsPerInch = GetDeviceCaps(lngDC, 88)
    Else                            'Vertical
        lngPixelsPerInch = GetDeviceCaps(lngDC, 90)
    End If
    lngDC = ReleaseDC(0, lngDC)
    PixelsToTwips = (lngTwips / lngPixelsPerInch) * nTwipsPerInch

End Function

Public Sub SetStatus(ByVal Message As String)
    ' Set the caption to the message.
    frmSendGetData.lblStatus.Caption = Message

    ' Let the OS update to display the message.
    DoEvents
End Sub

Public Sub MenuState(ByVal State As Long)
    ' Display the status form.
    frmSendGetData.Visible = True

    ' Display the message on the status form.
    Call SetStatus("Connecting to Server...")

    ' Find the proper state before proceeding.
    Select Case State
        Case MENU_STATE_NEWACCOUNT
            ' Hide the new account form.
            frmNewAccount.Visible = False

            ' Check if we're connected to the server.
            If ConnectToServer Then
                ' Display the message on the status form.
                Call SetStatus("Connected! Creating Account...")

                ' Check if the email field is visible and create the account.
                If Not frmNewAccount.txtEmail.Visible Then
                    Call SendNewAccount(frmNewAccount.txtName.Text, frmNewAccount.txtPassword.Text, vbNullString)
                Else
                    Call SendNewAccount(frmNewAccount.txtName.Text, frmNewAccount.txtPassword.Text, frmNewAccount.txtEmail.Text)
                End If
            End If

        Case MENU_STATE_DELACCOUNT
            ' Hide the delete account form.
            frmDeleteAccount.Visible = False

            ' Check if we're connected to the server.
            If ConnectToServer Then
                ' Display the message on the status form.
                Call SetStatus("Connected. Deleting Account...")

                ' Delete the requested account.
                Call SendDelAccount(frmDeleteAccount.txtName.Text, frmDeleteAccount.txtPassword.Text)
            End If

        Case MENU_STATE_LOGIN
            ' Hide the login form.
            frmLogin.Visible = False

            ' Check if we're connected to the server.
            If ConnectToServer Then
                ' Display the message on the status form.
                Call SetStatus("Connected. Logging In...")

                ' Login to the requested account.
                Call SendLogin(frmLogin.txtName.Text, frmLogin.txtPassword.Text)
            End If

        Case MENU_STATE_NEWCHAR
            ' Hide the character form.
            frmChars.Visible = False

            ' Check if we're connected to the server.
            If ConnectToServer Then
                ' Display the message on the status form.
                Call SetStatus("Connected. Receiving Classes...")

                ' Check if the sprite size is 32x64.
                If SpriteSize = 1 Then
                    frmNewChar.Picture4.Top = frmNewChar.Picture4.Top - 32
                    frmNewChar.Picture4.Height = 69
                    frmNewChar.Picpic.Height = 65
                End If

                ' Check if we're using custom players: faces, clothes, and legs.
                If CUSTOM_PLAYERS <> 0 Then
                    frmNewChar.hsHead.Visible = True
                    frmNewChar.hsBody.Visible = True
                    frmNewChar.hsLegs.Visible = True
                End If

                ' Request the available classes.
                Call SendGetClasses
            End If

        Case MENU_STATE_ADDCHAR
            ' Hide the new character form.
            frmNewChar.Visible = False

            ' Check if we're connected to the server.
            If ConnectToServer Then
                ' Display the message on the status form.
                Call SetStatus("Connected. Creating Character...")

                ' Check the gender and then create the character.
                If frmNewChar.optMale.Value Then
                    Call SendAddChar(frmNewChar.txtName, 0, frmNewChar.cmbClass.ListIndex, frmChars.lstChars.ListIndex + 1, frmNewChar.hsHead.Value, frmNewChar.hsBody.Value, frmNewChar.hsLegs.Value)
                Else
                    Call SendAddChar(frmNewChar.txtName, 1, frmNewChar.cmbClass.ListIndex, frmChars.lstChars.ListIndex + 1, frmNewChar.hsHead.Value, frmNewChar.hsBody.Value, frmNewChar.hsLegs.Value)
                End If
            End If

        Case MENU_STATE_DELCHAR
            ' Hide the character form.
            frmChars.Visible = False

            ' Check if we're connected to the server.
            If ConnectToServer Then
                ' Display the message on the status form.
                Call SetStatus("Connected. Deleting Character...")

                ' Delete the requested character.
                Call SendDelChar(frmChars.lstChars.ListIndex + 1)
            End If

        Case MENU_STATE_USECHAR
            ' Hide the character form.
            frmChars.Visible = False

            ' Check if we're connected to the server.
            If ConnectToServer Then
                ' Display the message on the status form.
                Call SetStatus("Connected. Entering " & GAME_NAME & "...")

                ' Request to join the game.
                Call SendUseChar(frmChars.lstChars.ListIndex + 1)
            End If
    End Select

    ' Check if we're not connected to the server.
    If Not IsConnected Then
        ' Hide the status form.
        frmSendGetData.Visible = False

        ' Display the main menu form.
        frmMainMenu.Visible = True

        ' Display the message box to the user.
        Call MsgBox("The server is currently offline. Please try to reconnect in a few minutes or visit " & WEBSITE, vbOKOnly, GAME_NAME)
    End If
End Sub

Public Sub GameInit()
    On Error Resume Next

    ' Generate the player's HP bar and check for divison errors.
    If GetPlayerMaxHP(MyIndex) > 0 Then
        frmMirage.shpHP.Width = ((GetPlayerHP(MyIndex)) / (GetPlayerMaxHP(MyIndex))) * 150
        frmMirage.lblHP.Caption = GetPlayerHP(MyIndex) & " / " & GetPlayerMaxHP(MyIndex)
    End If

    ' Generate the player's MP bar and check for division errors.
    If GetPlayerMaxMP(MyIndex) > 0 Then
        frmMirage.shpMP.Width = ((GetPlayerMP(MyIndex)) / (GetPlayerMaxMP(MyIndex))) * 150
        frmMirage.lblMP.Caption = GetPlayerMP(MyIndex) & " / " & GetPlayerMaxMP(MyIndex)
    End If

    ' Load the background of the game client.
    frmMirage.Picture = LoadPicture(App.Path & "\GUI\800x600.jpg")

    ' Load the background for all of the client tabs.
    frmMirage.picCharStatus.Picture = LoadPicture(App.Path & "\GUI\MiniMenu.jpg")
    frmMirage.picEquipment.Picture = LoadPicture(App.Path & "\GUI\MiniMenu.jpg")
    frmMirage.picPlayerSpells.Picture = LoadPicture(App.Path & "\GUI\MiniMenu.jpg")
    frmMirage.picInventory.Picture = LoadPicture(App.Path & "\GUI\MiniMenu.jpg")
    frmMirage.picGuildAdmin.Picture = LoadPicture(App.Path & "\GUI\MiniMenu.jpg")
    frmMirage.picWhosOnline.Picture = LoadPicture(App.Path & "\GUI\MiniMenu.jpg")
    frmMirage.picGuildMember.Picture = LoadPicture(App.Path & "\GUI\MiniMenu.jpg")
    frmMirage.picInventory3.Picture = LoadPicture(App.Path & "\GUI\MiniMenu.jpg")

    ' Unload main menu forms after character logs in.
    Unload frmSendGetData
    Unload frmMainMenu
    Unload frmChars
    Unload frmNewChar
    Unload frmSendGetData

    ' Stop the BGM incase there is menu music playing.
    Call StopBGM

    ' Prepare the in-game client loop.
    InGame = True

    ' Display the game client to the user.
    frmMirage.Visible = True

    ' Set the focus to the game client and game screen.
    frmMirage.SetFocus
    frmMirage.picScreen.SetFocus
End Sub

Sub GameLoop()
    Dim Tick As Long
    Dim TickFPS As Long
    Dim FPS As Long
    Dim X As Long
    Dim y As Long
    Dim I As Long
    Dim z As Long
    Dim sRECT As DXVBLib.RECT
    Dim dRECT As DXVBLib.RECT

    ' This will be re-enabled once Eclipse Evolution 2.7 is released. [Mellowz]
    On Error Resume Next

    ' Used for calculating fps
    TickFPS = GetTickCount
    FPS = 0

    ' *******************************************
    ' * ECLIPSE EVOLUTION MAIN GAME LOOP BEGIN  *
    ' *******************************************
    Do While InGame
        Tick = GetTickCount

        If frmMirage.WindowState = 0 Then

            ' Check if we need to restore surfaces
            If NeedToRestoreSurfaces Then
                DD.RestoreAllSurfaces
                
                Call BackBuffer_Create
                Call InitSurfaces
            End If

            If Not GettingMap Then

                ' Check to make sure they aren't trying to auto do anything
                If GetAsyncKeyState(VK_UP) >= 0 And DirUp Then
                    DirUp = False
                End If
                If GetAsyncKeyState(VK_DOWN) >= 0 And DirDown Then
                    DirDown = False
                End If
                If GetAsyncKeyState(VK_LEFT) >= 0 And DirLeft Then
                    DirLeft = False
                End If
                If GetAsyncKeyState(VK_RIGHT) >= 0 And DirRight Then
                    DirRight = False
                End If
                If GetAsyncKeyState(VK_CONTROL) >= 0 And ControlDown Then
                    ControlDown = False
                End If
                If GetAsyncKeyState(VK_SHIFT) >= 0 And ShiftDown Then
                    ShiftDown = False
                End If
    
                ' Check to make sure we are still connected
                If Not IsConnected Then
                    InGame = False
                    Exit Do
                End If

                NewX = 10
                NewY = 7

                NewPlayerY = Player(MyIndex).y - NewY
                NewPlayerX = Player(MyIndex).X - NewX

                NewX = NewX * PIC_X
                NewY = NewY * PIC_Y

                NewXOffset = Player(MyIndex).xOffset
                NewYOffset = Player(MyIndex).yOffset

                If Player(MyIndex).y - 7 < 1 Then
                    NewY = Player(MyIndex).y * PIC_Y + Player(MyIndex).yOffset
                    NewYOffset = 0
                    NewPlayerY = 0
                    If Player(MyIndex).y = 7 And Player(MyIndex).Dir = DIR_UP Then
                        NewPlayerY = Player(MyIndex).y - 7
                        NewY = 7 * PIC_Y
                        NewYOffset = Player(MyIndex).yOffset
                    End If
                ElseIf Player(MyIndex).y + 9 > MAX_MAPY + 1 Then
                    NewY = (Player(MyIndex).y - (MAX_MAPY - 14)) * PIC_Y + Player(MyIndex).yOffset
                    NewYOffset = 0
                    NewPlayerY = MAX_MAPY - 14
                    If Player(MyIndex).y = MAX_MAPY - 7 And Player(MyIndex).Dir = DIR_DOWN Then
                        NewPlayerY = Player(MyIndex).y - 7
                        NewY = 7 * PIC_Y
                        NewYOffset = Player(MyIndex).yOffset
                    End If
                End If

                If Player(MyIndex).X - 10 < 1 Then
                    NewX = Player(MyIndex).X * PIC_X + Player(MyIndex).xOffset
                    NewXOffset = 0
                    NewPlayerX = 0
                    If Player(MyIndex).X = 10 And Player(MyIndex).Dir = DIR_LEFT Then
                        NewPlayerX = Player(MyIndex).X - 10
                        NewX = 10 * PIC_X
                        NewXOffset = Player(MyIndex).xOffset
                    End If
                ElseIf Player(MyIndex).X + 11 > MAX_MAPX + 1 Then
                    NewX = (Player(MyIndex).X - (MAX_MAPX - 19)) * PIC_X + Player(MyIndex).xOffset
                    NewXOffset = 0
                    NewPlayerX = MAX_MAPX - 19
                    If Player(MyIndex).X = MAX_MAPX - 9 And Player(MyIndex).Dir = DIR_RIGHT Then
                        NewPlayerX = Player(MyIndex).X - 10
                        NewX = 10 * PIC_X
                        NewXOffset = Player(MyIndex).xOffset
                    End If
                End If

                ScreenX = GetScreenLeft(MyIndex)
                ScreenY = GetScreenTop(MyIndex)
                ScreenX2 = GetScreenRight(MyIndex)
                ScreenY2 = GetScreenBottom(MyIndex)

                If ScreenX < 0 Then
                    ScreenX = 0
                    ScreenX2 = 20
                ElseIf ScreenX2 > MAX_MAPX Then
                    ScreenX2 = MAX_MAPX
                    ScreenX = MAX_MAPX - 20
                End If
            
                If ScreenY < 0 Then
                    ScreenY = 0
                    ScreenY2 = 15
                ElseIf ScreenY2 > MAX_MAPY Then
                    ScreenY2 = MAX_MAPY
                    ScreenY = MAX_MAPY - 15
                End If

                sx = 32
                If MAX_MAPX = 19 Then
                    NewX = Player(MyIndex).X * PIC_X + Player(MyIndex).xOffset
                    NewXOffset = 0
                    NewPlayerX = 0
                    NewY = Player(MyIndex).y * PIC_Y + Player(MyIndex).yOffset
                    NewYOffset = 0
                    NewPlayerY = 0
                    ScreenX = 0
                    ScreenY = 0
                    ScreenX2 = MAX_MAPX
                    ScreenY2 = MAX_MAPY
                    sx = 0
                End If

                ' Blit out tiles layers ground/anim1/anim2
                For y = ScreenY To ScreenY2
                    For X = ScreenX To ScreenX2
                        Call BltTile(X, y)
                    Next X
                Next y

                If ScreenMode = 0 Then
                
                    ' Blit out the items
                    For I = 1 To MAX_MAP_ITEMS
                        If MapItem(I).num > 0 Then
                            Call BltItem(I)
                        End If
                    Next I
                    
                    ' Blit out NPC hp bars
                    If frmMirage.chkNpcBar.Value = vbChecked Then
                        For I = 1 To MAX_MAP_NPCS
                            Call BltNpcBars(I)
                        Next I
                    End If
                    
                     ' Blit players bar
                    If frmMirage.chkPlayerBar.Value = vbChecked Then
                        For I = 1 To MAX_PLAYERS
                            If IsPlaying(I) Then
                                If GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                                    Call BltPlayerBars(I)
                                End If
                            End If
                        Next I
                    End If

                    ' Blit out the sprite change attribute
                    If Right$(Trim$(Map(GetPlayerMap(MyIndex)).Name), 1) = "*" Then
                        For y = ScreenY To ScreenY2
                            For X = ScreenX To ScreenX2
                                Call BltSpriteChange(X, y)
                            Next X
                        Next y
                    End If

                    ' Blit out grapple
                    For I = 1 To MAX_PLAYERS
                        If IsPlaying(I) Then
                            If GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                                Call Bltgrapple(I)
                            End If
                        End If
                    Next I

                    ' Blit out players and arrows
                    For I = 1 To MAX_PLAYERS
                        If IsPlaying(I) Then
                            If GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                                Call BltPlayer(I)
                                Call BltArrow(I)
                            End If
                        End If
                    Next I
                    
                    ' Blit out the npc base
                    For I = 1 To MAX_MAP_NPCS
                        If MapNpc(I).num > 0 Then
                            Call BltNpcBody(I)
                        End If
                    Next I

                    ' Blit out the npc tops
                    For I = 1 To MAX_MAP_NPCS
                        If MapNpc(I).num > 0 Then
                            Call BltNpcTop(I)
                        End If
                    Next I

                    ' Blit out players top
                    If SpriteSize >= 1 Then
                        For I = 1 To MAX_PLAYERS
                            If IsPlaying(I) Then
                                If GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                                    Call BltPlayerTop(I)
                                End If
                            End If
                        Next I
                    End If

                    ' Blt out the spells
                    For I = 1 To MAX_PLAYERS
                        If IsPlaying(I) Then
                            If GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                                Call BltSpell(I)
                            End If
                        End If
                    Next I

                    ' Blt out the scripted spells
                    For I = 1 To MAX_SCRIPTSPELLS
                        If ScriptSpell(I).SpellNum > 0 Then
                            If ScriptSpell(I).SpellNum <= MAX_SPELLS Then
                                If ScriptSpell(I).CastedSpell = YES Then
                                    Call BltScriptSpell(I)
                                End If
                            End If
                        End If
                    Next I
                    
                    ' Draw 'level up!' text
                    For I = 1 To MAX_PLAYERS
                        If IsPlaying(I) Then
                            If GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                                Call BltLevelUp(I)
                            End If
                        End If
                    Next I

                End If

                ' Blit out tile layer fringe
                For y = ScreenY To ScreenY2
                    For X = ScreenX To ScreenX2
                        Call BltFringeTile(X, y)
                    Next X
                Next y

                ' Check for roof tiles
                For y = ScreenY To ScreenY2
                    For X = ScreenX To ScreenX2
                        If Not IsTileRoof(X, y) Then
                            Call BltFringe2Tile(X, y)
                        End If
                    Next X
                Next y
                
                ' Blit out emoticons
                For I = 1 To MAX_PLAYERS
                    If IsPlaying(I) Then
                        If GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                            Call BltEmoticons(I)
                        End If
                    End If
                Next I

                ' Draw night (for normal players).
                If GameTime = TIME_NIGHT Then
                    If Map(GetPlayerMap(MyIndex)).Indoors = 0 Then
                        If Not InEditor Then
                            Call Night
                        End If
                    End If
                End If
            
                ' Draw night (for administrators).
                If InEditor Then
                    If NightMode = 1 Then
                        Call Night
                    End If
                End If
            
                ' Draw weather (for all players)
                If Map(GetPlayerMap(MyIndex)).Indoors = 0 Then
                    If Map(GetPlayerMap(MyIndex)).Weather <> 0 Then
                        Call BltMapWeather
                    End If
            
                    Call BltWeather
                End If

                If InEditor Then
                    If GridMode = 1 Then
                        For y = ScreenY To ScreenY2
                            For X = ScreenX To ScreenX2
                                Call BltTile2(X * PIC_X, y * PIC_Y, 0)
                            Next X
                        Next y
                    End If
                End If

                ' Lock the backbuffer so we can draw text and names
                TexthDC = DD_BackBuffer.GetDC

                If ScreenMode = 0 Then
                
                    ' Draw NPC's damage on player
                    If frmMirage.chkNpcDamage.Value = 1 Then
                        If frmMirage.chkPlayerName.Value = 0 Then
                            If GetTickCount < NPCDmgTime + 2000 Then
                                Call DrawText(TexthDC, (Int(Len(NPCDmgDamage)) / 2) * 3 + NewX + sx, NewY - 22 - ii + sx, NPCDmgDamage, QBColor(BRIGHTRED))
                            End If
                        Else
                            If GetPlayerGuild(MyIndex) <> vbNullString Then
                                If GetTickCount < NPCDmgTime + 2000 Then
                                    Call DrawText(TexthDC, (Int(Len(NPCDmgDamage)) / 2) * 3 + NewX + sx, NewY - 42 - ii + sx, NPCDmgDamage, QBColor(BRIGHTRED))
                                End If
                            Else
                                If GetTickCount < NPCDmgTime + 2000 Then
                                    Call DrawText(TexthDC, (Int(Len(NPCDmgDamage)) / 2) * 3 + NewX + sx, NewY - 22 - ii + sx, NPCDmgDamage, QBColor(BRIGHTRED))
                                End If
                            End If
                        End If
                        ii = ii + 1
                    End If

                    ' Draw player's damage on NPC
                    If frmMirage.chkPlayerDamage.Value = 1 Then
                        If NPCWho > 0 Then
                            If MapNpc(NPCWho).num > 0 Then
                                If frmMirage.chkNpcName.Value = 0 Then
                                    If Npc(MapNpc(NPCWho).num).Big = 0 Then
                                        If GetTickCount < DmgTime + 2000 Then
                                            Call DrawText(TexthDC, (MapNpc(NPCWho).X - NewPlayerX) * PIC_X + sx + (Int(Len(DmgDamage)) / 2) * 3 + MapNpc(NPCWho).xOffset - NewXOffset, (MapNpc(NPCWho).y - NewPlayerY) * PIC_Y + sx - 20 + MapNpc(NPCWho).yOffset - NewYOffset - iii, DmgDamage, QBColor(WHITE))
                                        End If
                                    Else
                                        If GetTickCount < DmgTime + 2000 Then
                                            Call DrawText(TexthDC, (MapNpc(NPCWho).X - NewPlayerX) * PIC_X + sx + (Int(Len(DmgDamage)) / 2) * 3 + MapNpc(NPCWho).xOffset - NewXOffset, (MapNpc(NPCWho).y - NewPlayerY) * PIC_Y + sx - 47 + MapNpc(NPCWho).yOffset - NewYOffset - iii, DmgDamage, QBColor(WHITE))
                                        End If
                                    End If
                                Else
                                    If Npc(MapNpc(NPCWho).num).Big = 0 Then
                                        If GetTickCount < DmgTime + 2000 Then
                                            Call DrawText(TexthDC, (MapNpc(NPCWho).X - NewPlayerX) * PIC_X + sx + (Int(Len(DmgDamage)) / 2) * 3 + MapNpc(NPCWho).xOffset - NewXOffset, (MapNpc(NPCWho).y - NewPlayerY) * PIC_Y + sx - 30 + MapNpc(NPCWho).yOffset - NewYOffset - iii, DmgDamage, QBColor(WHITE))
                                        End If
                                    Else
                                        If GetTickCount < DmgTime + 2000 Then
                                            Call DrawText(TexthDC, (MapNpc(NPCWho).X - NewPlayerX) * PIC_X + sx + (Int(Len(DmgDamage)) / 2) * 3 + MapNpc(NPCWho).xOffset - NewXOffset, (MapNpc(NPCWho).y - NewPlayerY) * PIC_Y + sx - 57 + MapNpc(NPCWho).yOffset - NewYOffset - iii, DmgDamage, QBColor(WHITE))
                                        End If
                                    End If
                                End If
                                iii = iii + 1
                            End If
                        End If
                    End If
                    
                    ' Draw player name and guild name
                    If frmMirage.chkPlayerName.Value = 1 Then
                        For I = 1 To MAX_PLAYERS
                            If IsPlaying(I) Then
                                If GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                                    Call BltPlayerGuildName(I)
                                    Call BltPlayerName(I)
                                End If
                            End If
                        Next I
                    End If

                    ' speech bubble stuffs
                    If ReadINI("CONFIG", "SpeechBubbles", App.Path & "\config.ini") = 1 Then
                        For I = 1 To MAX_PLAYERS
                            If IsPlaying(I) Then
                                If GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                                    If Bubble(I).Text <> vbNullString Then
                                        Call BltPlayerText(I)
                                    End If
    
                                    If GetTickCount() > Bubble(I).Created + DISPLAY_BUBBLE_TIME Then
                                        Bubble(I).Text = vbNullString
                                    End If
                                End If
                            End If
                        Next I
                    End If

                    ' scriptbubble stuffs
                    For z = 1 To MAX_BUBBLES
                        If IsPlaying(MyIndex) Then
                            If GetPlayerMap(MyIndex) = ScriptBubble(z).Map Then
                                If ScriptBubble(z).Text <> vbNullString Then
                                    Call Bltscriptbubble(z, ScriptBubble(z).Map, ScriptBubble(z).X, ScriptBubble(z).y, ScriptBubble(z).Colour)
                                End If
    
                                If GetTickCount() > ScriptBubble(z).Created + DISPLAY_BUBBLE_TIME Then
                                    ScriptBubble(z).Text = vbNullString
                                End If
                            End If
                        End If
                    Next z

                    ' Draw NPC Names
                    If ReadINI("CONFIG", "NPCName", App.Path & "\config.ini") = 1 Then
                        For I = LBound(MapNpc) To UBound(MapNpc)
                            If MapNpc(I).num > 0 Then
                                Call BltMapNPCName(I)
                            End If
                        Next I
                    End If

                    ' Blit out attribs if in editor
                    If InEditor Then
                        For y = 0 To MAX_MAPY
                            For X = 0 To MAX_MAPX
                                With Map(GetPlayerMap(MyIndex)).Tile(X, y)
                                    If .Type = TILE_TYPE_BLOCKED Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "B", QBColor(BRIGHTRED))
                                    End If
                                    If .Type = TILE_TYPE_WARP Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "W", QBColor(BRIGHTBLUE))
                                    End If
                                    If .Type = TILE_TYPE_ITEM Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "I", QBColor(WHITE))
                                    End If
                                    If .Type = TILE_TYPE_NPCAVOID Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "N", QBColor(WHITE))
                                    End If
                                    If .Type = TILE_TYPE_KEY Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "K", QBColor(WHITE))
                                    End If
                                    If .Type = TILE_TYPE_KEYOPEN Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "O", QBColor(WHITE))
                                    End If
                                    If .Type = TILE_TYPE_HEAL Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "H", QBColor(BRIGHTGREEN))
                                    End If
                                    If .Type = TILE_TYPE_KILL Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "K", QBColor(BRIGHTRED))
                                    End If
                                    If .Type = TILE_TYPE_SHOP Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "S", QBColor(YELLOW))
                                    End If
                                    If .Type = TILE_TYPE_CBLOCK Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "CB", QBColor(BLACK))
                                    End If
                                    If .Type = TILE_TYPE_ARENA Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "A", QBColor(BRIGHTGREEN))
                                    End If
                                    If .Type = TILE_TYPE_SOUND Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "PS", QBColor(YELLOW))
                                    End If
                                    If .Type = TILE_TYPE_SPRITE_CHANGE Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "SC", QBColor(GREY))
                                    End If
                                    If .Type = TILE_TYPE_SIGN Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "SI", QBColor(YELLOW))
                                    End If
                                    If .Type = TILE_TYPE_DOOR Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "D", QBColor(BLACK))
                                    End If
                                    If .Type = TILE_TYPE_NOTICE Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "N", QBColor(BRIGHTGREEN))
                                    End If
                                    If .Type = TILE_TYPE_CHEST Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "C", QBColor(BROWN))
                                    End If
                                    If .Type = TILE_TYPE_CLASS_CHANGE Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "CG", QBColor(WHITE))
                                    End If
                                    If .Type = TILE_TYPE_SCRIPTED Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "SC", QBColor(YELLOW))
                                    End If
                                    If .Type = TILE_TYPE_HOUSE Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "PH", QBColor(YELLOW))
                                    End If
                                    If .light > 0 Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 18 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 14 - (NewPlayerY * PIC_Y) - NewYOffset, "L", QBColor(YELLOW))
                                    End If
                                    If .Type = TILE_TYPE_BANK Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "BANK", QBColor(BRIGHTRED))
                                    End If
                                    If .Type = TILE_TYPE_GUILDBLOCK Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "GB", QBColor(MAGENTA))
                                    End If
                                    If .Type = TILE_TYPE_HOOKSHOT Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "GS", QBColor(WHITE))
                                    End If
                                    If .Type = TILE_TYPE_WALKTHRU Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "WT", QBColor(RED))
                                    End If
                                    If .Type = TILE_TYPE_ROOF Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "RF", QBColor(RED))
                                    End If
                                    If .Type = TILE_TYPE_ROOFBLOCK Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "RFB", QBColor(BRIGHTRED))
                                    End If
                                    If .Type = TILE_TYPE_ONCLICK Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "OC", QBColor(WHITE))
                                    End If
                                    If .Type = TILE_TYPE_LOWER_STAT Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "-S", QBColor(BRIGHTRED))
                                    End If
                                End With
                            Next X
                        Next y
                    End If

                    ' draw FPS
                    If BFPS Then
                        Call DrawText(TexthDC, 18 * PIC_X + sx, sx, "FPS: " & GameFPS, QBColor(YELLOW))
                    End If

                    ' draw cursor and player X and Y locations
                    If BLoc Then
                        Call DrawText(TexthDC, 0 + sx, 0 + sx, "Cursor (X: " & CurX & "; Y: " & CurY & ")", QBColor(YELLOW))
                        Call DrawText(TexthDC, 0 + sx, 15 + sx, "Location (X: " & GetPlayerX(MyIndex) & "; Y: " & GetPlayerY(MyIndex) & ")", QBColor(YELLOW))
                        Call DrawText(TexthDC, 0 + sx, 30 + sx, "Map #" & GetPlayerMap(MyIndex), QBColor(YELLOW))
                    End If

                    ' Draw map name
                    If Map(GetPlayerMap(MyIndex)).Moral = MAP_MORAL_NONE Then
                        Call DrawText(TexthDC, Int((20.5) * PIC_X / 2) - (Int(Len(Trim$(Map(GetPlayerMap(MyIndex)).Name)) / 2) * 8) + sx, 2 + sx, Trim$(Map(GetPlayerMap(MyIndex)).Name), QBColor(BRIGHTRED))
                    ElseIf Map(GetPlayerMap(MyIndex)).Moral = MAP_MORAL_HOUSE Then
                        Call DrawText(TexthDC, Int((20.5) * PIC_X / 2) - (Int(Len(Trim$(Map(GetPlayerMap(MyIndex)).Name)) / 2) * 8) + sx, 2 + sx, Trim$(Map(GetPlayerMap(MyIndex)).Name), QBColor(YELLOW))
                    ElseIf Map(GetPlayerMap(MyIndex)).Moral = MAP_MORAL_SAFE Then
                        Call DrawText(TexthDC, Int((20.5) * PIC_X / 2) - (Int(Len(Trim$(Map(GetPlayerMap(MyIndex)).Name)) / 2) * 8) + sx, 2 + sx, Trim$(Map(GetPlayerMap(MyIndex)).Name), QBColor(WHITE))
                    ElseIf Map(GetPlayerMap(MyIndex)).Moral = MAP_MORAL_NO_PENALTY Then
                        Call DrawText(TexthDC, Int((20.5) * PIC_X / 2) - (Int(Len(Trim$(Map(GetPlayerMap(MyIndex)).Name)) / 2) * 8) + sx, 2 + sx, Trim$(Map(GetPlayerMap(MyIndex)).Name), QBColor(BLACK))
                    End If

                    For I = 1 To MAX_BLT_LINE
                        If BattlePMsg(I).Index > 0 Then
                            If BattlePMsg(I).Time + 7000 > GetTickCount Then
                                Call DrawText(TexthDC, 1 + sx, BattlePMsg(I).y + frmMirage.picScreen.Height - 15 + sx, Trim$(BattlePMsg(I).Msg), QBColor(BattlePMsg(I).Color))
                            Else
                                BattlePMsg(I).Done = 0
                            End If
                        End If

                        If BattleMMsg(I).Index > 0 Then
                            If BattleMMsg(I).Time + 7000 > GetTickCount Then
                                Call DrawText(TexthDC, (frmMirage.picScreen.Width - (Len(BattleMMsg(I).Msg) * 8)) + sx, BattleMMsg(I).y + frmMirage.picScreen.Height - 15 + sx, Trim$(BattleMMsg(I).Msg), QBColor(BattleMMsg(I).Color))
                            Else
                                BattleMMsg(I).Done = 0
                            End If
                        End If
                    Next I
                        
                End If
                
            Else
                ' Lock the backbuffer so we can draw text
                TexthDC = DD_BackBuffer.GetDC
                
                ' Show player that a new map is loading
                Call DrawText(TexthDC, PIC_X, PIC_Y, "Receiving Map...", QBColor(BRIGHTCYAN))
            End If

            ' Release DC
            Call DD_BackBuffer.ReleaseDC(TexthDC)

            ' Get the rect for the back buffer to blit from
            sRECT.Top = 0
            sRECT.Bottom = (MAX_MAPY + 1) * PIC_Y
            sRECT.Left = 0
            sRECT.Right = (MAX_MAPX + 1) * PIC_X

            ' Get the rect to blit to
            Call DX.GetWindowRect(frmMirage.picScreen.hWnd, dRECT)
            dRECT.Bottom = dRECT.Top - sx + ((MAX_MAPY + 1) * PIC_Y)
            dRECT.Right = dRECT.Left - sx + ((MAX_MAPX + 1) * PIC_X)
            dRECT.Top = dRECT.Bottom - ((MAX_MAPY + 1) * PIC_Y)
            dRECT.Left = dRECT.Right - ((MAX_MAPX + 1) * PIC_X)

            ' Blit the backbuffer
            Call DD_PrimarySurf.Blt(dRECT, DD_BackBuffer, sRECT, DDBLT_WAIT)

            ' Check if player is trying to move
            Call CheckMovement

            ' Check to see if player is trying to attack
            Call CheckAttack

            ' Process player movements (actually move them)
            For I = 1 To MAX_PLAYERS
                If IsPlaying(I) Then
                    Call ProcessMovement(I)
                End If
            Next I

            ' Process npc movements (actually move them)
            For I = 1 To MAX_MAP_NPCS
                If Map(GetPlayerMap(MyIndex)).Npc(I) > 0 Then
                    Call ProcessNpcMovement(I)
                End If
            Next I

        End If

        ' Change map animation every 250 milliseconds
        If GetTickCount > MapAnimTimer + 250 Then
            If MapAnim = 0 Then
                MapAnim = 1
            Else
                MapAnim = 0
            End If
            MapAnimTimer = GetTickCount
        End If

        ' Lock fps
        Do While GetTickCount < Tick + 31
            DoEvents
            Sleep 1
        Loop

        ' Calculate fps
        If GetTickCount > TickFPS + 1000 Then
            GameFPS = FPS
            TickFPS = GetTickCount
            FPS = 0
        Else
            FPS = FPS + 1
        End If

        DoEvents
    Loop

    frmSendGetData.Visible = True

    Call SetStatus("Closing Game...")

    ' MsgBox "Connection lost!"

    ' Shutdown the game
    Call GameDestroy

    Exit Sub
End Sub

' Closes the game client.
Sub GameDestroy()
    ' Unloads all TCP-related things.
    Call TCPDestroy

    ' Unloads all DirectX objects.
    Call DestroyDirectX

    ' Unloads the BGM in memory (soon-to-be obsolete).
    Call StopBGM

    ' Closes the VB6 application.
    End
End Sub

Sub BltTile(ByVal X As Long, ByVal y As Long)
    Dim Ground As Long
    Dim Mask1 As Long
    Dim Anim1 As Long
    Dim Mask2 As Long
    Dim Anim2 As Long
    Dim GroundTileSet As Byte
    Dim Mask1TileSet As Byte
    Dim Anim1TileSet As Byte
    Dim Mask2TileSet As Byte
    Dim Anim2TileSet As Byte
    Dim rec As DXVBLib.RECT

    Ground = Map(GetPlayerMap(MyIndex)).Tile(X, y).Ground
    Mask1 = Map(GetPlayerMap(MyIndex)).Tile(X, y).Mask
    Anim1 = Map(GetPlayerMap(MyIndex)).Tile(X, y).Anim
    Mask2 = Map(GetPlayerMap(MyIndex)).Tile(X, y).Mask2
    Anim2 = Map(GetPlayerMap(MyIndex)).Tile(X, y).M2Anim

    GroundTileSet = Map(GetPlayerMap(MyIndex)).Tile(X, y).GroundSet
    Mask1TileSet = Map(GetPlayerMap(MyIndex)).Tile(X, y).MaskSet
    Anim1TileSet = Map(GetPlayerMap(MyIndex)).Tile(X, y).AnimSet
    Mask2TileSet = Map(GetPlayerMap(MyIndex)).Tile(X, y).Mask2Set
    Anim2TileSet = Map(GetPlayerMap(MyIndex)).Tile(X, y).M2AnimSet

    If TileFile(GroundTileSet) = 0 Then
        Exit Sub
    End If

    rec.Top = Int(Ground / TilesInSheets) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (Ground - Int(Ground / TilesInSheets) * TilesInSheets) * PIC_X
    rec.Right = rec.Left + PIC_X

    Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset + sx, (y - NewPlayerY) * PIC_Y - NewYOffset + sx, DD_TileSurf(GroundTileSet), rec, DDBLTFAST_WAIT)

    If MapAnim = 0 Or Anim1 = 0 Then
        If Mask1 > 0 Then
            If TileFile(Mask1TileSet) = 0 Then
                Exit Sub
            End If

            If TempTile(X, y).DoorOpen = NO Then
                rec.Top = Int(Mask1 / TilesInSheets) * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                rec.Left = (Mask1 - Int(Mask1 / TilesInSheets) * TilesInSheets) * PIC_X
                rec.Right = rec.Left + PIC_X
                
                Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset + sx, (y - NewPlayerY) * PIC_Y - NewYOffset + sx, DD_TileSurf(Mask1TileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        End If
    Else
        If Anim1 > 0 Then
            If TileFile(Anim1TileSet) = 0 Then
                Exit Sub
            End If

            rec.Top = Int(Anim1 / TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Anim1 - Int(Anim1 / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X

            Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset + sx, (y - NewPlayerY) * PIC_Y - NewYOffset + sx, DD_TileSurf(Anim1TileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If

    If MapAnim = 0 Or Anim2 = 0 Then
        If Mask2 > 0 Then
            If TileFile(Mask2TileSet) = 0 Then
                Exit Sub
            End If

            rec.Top = Int(Mask2 / TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Mask2 - Int(Mask2 / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X

            Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset + sx, (y - NewPlayerY) * PIC_Y - NewYOffset + sx, DD_TileSurf(Mask2TileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        If Anim2 > 0 Then
            If TileFile(Anim2TileSet) = 0 Then
                Exit Sub
            End If

            rec.Top = Int(Anim2 / TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Anim2 - Int(Anim2 / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X

            Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset + sx, (y - NewPlayerY) * PIC_Y - NewYOffset + sx, DD_TileSurf(Anim2TileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
End Sub

Sub BltItem(ByVal ItemNum As Long)
    Dim rec As DXVBLib.RECT

    rec.Top = Int(Item(MapItem(ItemNum).num).Pic / 6) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (Item(MapItem(ItemNum).num).Pic - Int(Item(MapItem(ItemNum).num).Pic / 6) * 6) * PIC_X
    rec.Right = rec.Left + PIC_X

    Call DD_BackBuffer.BltFast((MapItem(ItemNum).X - NewPlayerX) * PIC_X + sx - NewXOffset, (MapItem(ItemNum).y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltFringeTile(ByVal X As Long, ByVal y As Long)
    Dim Fringe As Long
    Dim FAnim As Long
    Dim FringeTileSet As Byte
    Dim FAnimTileSet As Byte
    Dim rec As DXVBLib.RECT

    Fringe = Map(GetPlayerMap(MyIndex)).Tile(X, y).Fringe
    FAnim = Map(GetPlayerMap(MyIndex)).Tile(X, y).FAnim

    FringeTileSet = Map(GetPlayerMap(MyIndex)).Tile(X, y).FringeSet
    FAnimTileSet = Map(GetPlayerMap(MyIndex)).Tile(X, y).FAnimSet

    If MapAnim = 0 Or FAnim = 0 Then
        If Fringe > 0 Then
            If TileFile(FringeTileSet) = 0 Then
                Exit Sub
            End If

            rec.Top = Int(Fringe / TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Fringe - Int(Fringe / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X

            Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_TileSurf(FringeTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        If FAnim > 0 Then
            If TileFile(FAnimTileSet) = 0 Then
                Exit Sub
            End If

            rec.Top = Int(FAnim / TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (FAnim - Int(FAnim / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X

            Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_TileSurf(FAnimTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
End Sub

Sub BltFringe2Tile(ByVal X As Integer, ByVal y As Integer)
    Dim Fringe2 As Long
    Dim F2Anim As Long
    Dim Fringe2TileSet As Byte
    Dim F2AnimTileSet As Byte
    Dim rec As DXVBLib.RECT

    Fringe2 = Map(GetPlayerMap(MyIndex)).Tile(X, y).Fringe2
    F2Anim = Map(GetPlayerMap(MyIndex)).Tile(X, y).F2Anim

    Fringe2TileSet = Map(GetPlayerMap(MyIndex)).Tile(X, y).Fringe2Set
    F2AnimTileSet = Map(GetPlayerMap(MyIndex)).Tile(X, y).F2AnimSet

    If MapAnim = 0 Or F2Anim = 0 Then
        If Fringe2 > 0 Then
            If TileFile(Fringe2TileSet) = 0 Then
                Exit Sub
            End If

            rec.Top = Int(Fringe2 / TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Fringe2 - Int(Fringe2 / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X

            Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_TileSurf(Fringe2TileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        If F2Anim > 0 Then
            If TileFile(F2AnimTileSet) = 0 Then
                Exit Sub
            End If

            rec.Top = Int(F2Anim / TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (F2Anim - Int(F2Anim / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X

            Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_TileSurf(F2AnimTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
End Sub

Sub BltPlayer(ByVal Index As Long)
    Dim Anim As Byte
    Dim X As Long, y As Integer
    Dim AttackSpeed As Long
    Dim temp As Long
    Dim attack_weaponslot As Long
    Dim attack_item As Long
    Dim rec As DXVBLib.RECT

    attack_weaponslot = Int(GetPlayerWeaponSlot(Index))

    If attack_weaponslot > 0 Then
        attack_item = Int(Player(Index).Inv(attack_weaponslot).num)
        If attack_item > 0 Then
            AttackSpeed = 1000 'Item(attack_item).AttackSpeed
        Else
            AttackSpeed = 1000
        End If
    Else
        AttackSpeed = 1000
    End If

    ' Only used if ever want to switch to blt rather then bltfast
    ' With rec_pos
    ' .Top = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset
    ' .Bottom = .Top + PIC_Y
    ' .Left = GetPlayerX(Index) * PIC_X + Player(Index).xOffset
    ' .Right = .Left + PIC_X
    ' End With

    ' Check for animation
    Anim = 0
    If Player(Index).Attacking = 0 Then
        Select Case GetPlayerDir(Index)
            Case DIR_UP
                If (Player(Index).yOffset < PIC_Y / 2) Then
                    Anim = 1
                End If
            Case DIR_DOWN
                If (Player(Index).yOffset < PIC_Y / 2 * -1) Then
                    Anim = 1
                End If
            Case DIR_LEFT
                If (Player(Index).xOffset < PIC_Y / 2) Then
                    Anim = 1
                End If
            Case DIR_RIGHT
                If (Player(Index).xOffset < PIC_Y / 2 * -1) Then
                    Anim = 1
                End If
        End Select
    Else
        If Player(Index).AttackTimer + Int(AttackSpeed / 2) > GetTickCount Then
            Anim = 2
        End If
    End If

    ' Check to see if we want to stop making him attack
    If Player(Index).AttackTimer + AttackSpeed < GetTickCount Then
        Player(Index).Attacking = 0
    '    Player(Index).AttackTimer = 0
    End If

    ' Configure what happens if theres no items there
    temp = GetPlayerShieldSlot(Index)
    If temp > 0 Then
        Player(Index).Shield = GetPlayerInvItemNum(Index, temp)
    Else
        Player(Index).Shield = 0
    End If
    
    temp = GetPlayerArmorSlot(Index)
    If temp > 0 Then
        Player(Index).Armor = GetPlayerInvItemNum(Index, temp)
    Else
        Player(Index).Armor = 0
    End If
    
    temp = GetPlayerHelmetSlot(Index)
    If temp > 0 Then
        Player(Index).Helmet = GetPlayerInvItemNum(Index, temp)
    Else
        Player(Index).Helmet = 0
    End If
    
    temp = GetPlayerWeaponSlot(Index)
    If temp > 0 Then
        Player(Index).Weapon = GetPlayerInvItemNum(Index, temp)
    Else
        Player(Index).Weapon = 0
    End If
    
    temp = GetPlayerRingSlot(Index)
    If temp > 0 Then
        Player(Index).Ring = GetPlayerInvItemNum(Index, temp)
    Else
        Player(Index).Ring = 0
    End If
    
    temp = GetPlayerNecklaceSlot(Index)
    If temp > 0 Then
        Player(Index).Necklace = GetPlayerInvItemNum(Index, temp)
    Else
        Player(Index).Necklace = 0
    End If
    
    temp = GetPlayerLegsSlot(Index)
    If temp > 0 Then
        Player(Index).legs = GetPlayerInvItemNum(Index, temp)
    Else
        Player(Index).legs = 0
    End If

    ' 32 X 64
    If SpriteSize = 1 Then

        ' 32 X 64
        If PaperDoll = 1 And GetPlayerPaperdoll(Index) = 1 Then

            rec.Left = (GetPlayerDir(Index) * 3 + Anim) * 32
            rec.Right = rec.Left + 32

            If Index = MyIndex Then
                X = NewX + sx
                y = NewY + sx

                ' PLAYER 32 X 64 IF DIR = UP
                If GetPlayerDir(MyIndex) = DIR_UP Then

                    ' PLAYER 32 X 64 BLIT SHIELD IF DIR = UP
                    If Player(MyIndex).Shield > 0 Then
                        rec.Top = Item(Player(MyIndex).Shield).Pic * 64 + PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' PLAYER 32 X 64 BLIT WEAPON IF DIR = UP
                    If Player(MyIndex).Weapon > 0 Then
                        rec.Top = Item(Player(MyIndex).Weapon).Pic * 64 + PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' PLAYER 32 X 64 BLIT NECKLACE IF DIR = UP
                    If Player(MyIndex).Necklace > 0 Then
                        rec.Top = Item(Player(MyIndex).Necklace).Pic * 64 + PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                End If

                If CUSTOM_PLAYERS = 0 Then
                    ' PLAYER 32 X 64 BLIT SPRITE
                    rec.Top = GetPlayerSprite(MyIndex) * 64 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    ' PLAYER 32 X 64 BLIT SPRITE
                    rec.Top = Player(Index).head * 64 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, y, DD_PlayerHead, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.Top = Player(Index).body * 64 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, y, DD_PlayerBody, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.Top = Player(Index).leg * 64 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, y, DD_PlayerLegs, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' PLAYER 32 X 64 BLIT LEGS
                If Player(MyIndex).legs > 0 Then
                    rec.Top = Item(Player(MyIndex).legs).Pic * 64 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' PLAYER 32 X 64 BLIT ARMOR
                If Player(MyIndex).Armor > 0 Then
                    rec.Top = Item(Player(MyIndex).Armor).Pic * 64 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' PLAYER 32 X 64 BLIT HELMET
                If Player(MyIndex).Helmet > 0 Then
                    rec.Top = Item(Player(MyIndex).Helmet).Pic * 64 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' PLAYER 32 X 64 DIR <> UP
                If GetPlayerDir(MyIndex) <> DIR_UP Then

                    ' PLAYER 32 X 64 BLIT SHIELD IF DIR <> UP
                    If Player(MyIndex).Shield > 0 Then
                        rec.Top = Item(Player(MyIndex).Shield).Pic * 64 + PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' PLAYER 32 X 64 BLIT WEAPON IF DIR <> UP
                    If Player(MyIndex).Weapon > 0 Then
                        rec.Top = Item(Player(MyIndex).Weapon).Pic * 64 + PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' PLAYER 32 X 64 BLIT NECKLACE IF DIR <> UP
                    If Player(MyIndex).Necklace > 0 Then
                        rec.Top = Item(Player(MyIndex).Necklace).Pic * 64 + PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                End If


            ' 32 X 64 IF OTHER PLAYER
            Else

                X = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset
                y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset

                ' IF BLIT IS OFFSCREEN ADJUST THE Y VALUE
                ' If y < 0 Then
                ' rec.tOp = rec.tOp + (y * -1)
                ' y = 0
                ' End If

                ' OTHER 32 X 64 IF DIR = UP
                If GetPlayerDir(Index) = DIR_UP Then

                    ' OTHER 32 X 64 BLIT SHIELD IF DIR = UP
                    If Player(Index).Shield > 0 Then
                        rec.Top = Item(Player(Index).Shield).Pic * 64 + PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' OTHER 32 X 64 BLIT WEAPON IF DIR = UP
                    If Player(Index).Weapon > 0 Then
                        rec.Top = Item(Player(Index).Weapon).Pic * 64 + PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' OTHER 32 X 64 BLIT NECKLACE IF DIR = UP
                    If Player(Index).Necklace > 0 Then
                        rec.Top = Item(Player(Index).Necklace).Pic * 64 + PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                End If

                ' OTHER 32 X 64 BLIT SPRITE
                If 0 + CUSTOM_PLAYERS = 0 Then
                    rec.Top = GetPlayerSprite(Index) * 64 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    rec.Top = Player(Index).head * 64 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_PlayerHead, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.Top = Player(Index).body * 64 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_PlayerBody, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.Top = Player(Index).leg * 64 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_PlayerLegs, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' OTHER 32 X 64 BLIT LEGS
                If Player(Index).legs > 0 Then
                    rec.Top = Item(Player(Index).legs).Pic * 64 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' OTHER 32 X 64 BLIT ARMOR
                If Player(Index).Armor > 0 Then
                    rec.Top = Item(Player(Index).Armor).Pic * 64 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' OTHER 32 X 64 BLIT HELMET
                If Player(Index).Helmet > 0 Then
                    rec.Top = Item(Player(Index).Helmet).Pic * 64 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' OTHER 32 X 64 IF DIR <> UP
                If GetPlayerDir(Index) <> DIR_UP Then

                    ' OTHER 32 X 64 BLIT SHIELD IF DIR <> UP
                    If Player(Index).Shield > 0 Then
                        rec.Top = Item(Player(Index).Shield).Pic * 64 + PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' OTHER 32 X 64 BLIT NECKLACE IF DIR <> UP
                    If Player(Index).Necklace > 0 Then
                        rec.Top = Item(Player(Index).Necklace).Pic * 64 + PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' 'OTHER 32 X 64 BLIT WEAPON IF DIR <> UP
                    If Player(Index).Weapon > 0 Then
                        rec.Top = Item(Player(Index).Weapon).Pic * 64 + PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                End If

            ' END OF PAPERDOLL FOR 32 X 64
            End If

        ' IF 32 X 64 AND NO PAPERDOLL
        Else
            rec.Left = (GetPlayerDir(Index) * 3 + Anim) * 32
            rec.Right = rec.Left + 32

            ' PLAYER 32 X 64
            If Index = MyIndex Then
                X = NewX + sx
                y = NewY + sx

                If 0 + CUSTOM_PLAYERS = 0 Then
                    ' PLAYER 32 X 64 BLIT SPRITE
                    rec.Top = GetPlayerSprite(MyIndex) * 64 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    ' PLAYER 32 X 64 BLIT SPRITE
                    rec.Top = Player(Index).head * 64 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, y, DD_PlayerHead, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.Top = Player(Index).body * 64 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, y, DD_PlayerBody, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.Top = Player(Index).leg * 64 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, y, DD_PlayerLegs, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

            ' OTHER 32 X 64
            Else
                X = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset
                y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset

' ADJUST IF OFF EDGE OF SCREEN
' If y < 0 Then
' rec.tOp = rec.tOp + (y * -1)
' y = 0
' 11111  End If

                ' OTHER 32 X 64 BLIT SPRITE
                If 0 + CUSTOM_PLAYERS = 0 Then
                    rec.Top = GetPlayerSprite(Index) * 64 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    rec.Top = Player(Index).head * 64 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_PlayerHead, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.Top = Player(Index).body * 64 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_PlayerBody, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.Top = Player(Index).leg * 64 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_PlayerLegs, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            End If

        ' END OF 32 X 64
        End If

    ' 32 X 32 LOOP
    ElseIf SpriteSize = 0 Then

        rec.Top = GetPlayerSprite(Index) * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = (GetPlayerDir(Index) * 3 + Anim) * PIC_X
        rec.Right = rec.Left + PIC_X

        ' 32 X 32 PLAYER
        If Index = MyIndex Then

            ' 32 X 32 PAPERDOLLED PLAYER
            If PaperDoll = 1 And GetPlayerPaperdoll(Index) = 1 Then
                X = NewX + sx
                y = NewY + sx

                ' PLAYER 32 X 32 IF DIR = UP
                If GetPlayerDir(MyIndex) = DIR_UP Then

                    ' PLAYER 32 X 32 BLIT SHIELD IF DIR = UP
                    If Player(MyIndex).Shield > 0 Then
                        rec.Top = Item(Player(MyIndex).Shield).Pic * PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' PLAYER 32 X 32 BLIT WEAPON IF DIR = UP
                    If Player(MyIndex).Weapon > 0 Then
                        rec.Top = Item(Player(MyIndex).Weapon).Pic * PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' PLAYER 32 X 32 BLIT NECKLACE IF DIR = UP
                    If Player(MyIndex).Necklace > 0 Then
                        rec.Top = Item(Player(MyIndex).Necklace).Pic * PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                End If

                If 0 + CUSTOM_PLAYERS = 0 Then
                    ' PLAYER 32 X 32 BLIT SPRITE
                    rec.Top = GetPlayerSprite(MyIndex) * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    ' PLAYER 32 X 32 BLIT SPRITE
                    rec.Top = Player(Index).head * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, y, DD_PlayerHead, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.Top = Player(Index).body * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, y, DD_PlayerBody, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.Top = Player(Index).leg * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, y, DD_PlayerLegs, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' PLAYER 32 X 32 BLIT LEGS
                If Player(MyIndex).legs > 0 Then
                    rec.Top = Item(Player(MyIndex).legs).Pic * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' PLAYER 32 X 32 BLIT ARMOR
                If Player(MyIndex).Armor > 0 Then
                    rec.Top = Item(Player(MyIndex).Armor).Pic * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' PLAYER 32 X 32 BLIT HELMET
                If Player(MyIndex).Helmet > 0 Then
                    rec.Top = Item(Player(MyIndex).Helmet).Pic * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' PLAYER 32 X 32 IF DIR <> UP
                If GetPlayerDir(MyIndex) <> DIR_UP Then

                    ' PLAYER 32 X 32 BLIT SHIELD IF DIR <> UP
                    If Player(MyIndex).Shield > 0 Then
                        rec.Top = Item(Player(MyIndex).Shield).Pic * PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' PLAYER 32 X 32 BLIT WEAPON IF DIR <> UP
                    If Player(MyIndex).Weapon > 0 Then
                        rec.Top = Item(Player(MyIndex).Weapon).Pic * PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' PLAYER 32 X 32 BLIT NECKLACE IF DIR <> UP
                    If Player(MyIndex).Necklace > 0 Then
                        rec.Top = Item(Player(MyIndex).Necklace).Pic * PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                End If

            ' 32 X 32 IF NO PAPERDOLL ON SELF BLIT JUST SPRITE
            Else
                X = NewX + sx
                y = NewY + sx
                If 0 + CUSTOM_PLAYERS = 0 Then
                    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    rec.Top = Player(Index).head * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, y, DD_PlayerHead, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.Top = Player(Index).body * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, y, DD_PlayerBody, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.Top = Player(Index).leg * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, y, DD_PlayerLegs, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            End If

        ' 32 X 32 OTHER LOOP
        Else
            X = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset
            y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset '- 4

            ' IF OFF TOP EDGE ADJUST
            If y < 0 Then
                rec.Top = rec.Top + (y * -1)
                y = 0
            End If

            ' 32 X 32 OTHER PAPERDOLL LOOP
            If PaperDoll = 1 And GetPlayerPaperdoll(Index) = 1 Then

                ' OTHER 32 X 32 IF DIR = UP
                If GetPlayerDir(Index) = DIR_UP Then

                    ' OTHER 32 X 32 BLIT SHIELD IF DIR = UP
                    If Player(Index).Shield > 0 Then
                        rec.Top = Item(Player(Index).Shield).Pic * PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' OTHER 32 X 32 BLIT WEAPON IF DIR = UP
                    If Player(Index).Weapon > 0 Then
                        rec.Top = Item(Player(Index).Weapon).Pic * PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' OTHER 32 X 32 BLIT NECKLACE IF DIR = UP
                    If Player(Index).Necklace > 0 Then
                        rec.Top = Item(Player(Index).Necklace).Pic * PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                End If

                ' OTHER 32 X 32 BLIT SPRITE
                If 0 + CUSTOM_PLAYERS = 0 Then
                    rec.Top = GetPlayerSprite(Index) * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    rec.Top = Player(Index).head * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_PlayerHead, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.Top = Player(Index).body * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_PlayerBody, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.Top = Player(Index).leg * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_PlayerLegs, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' OTHER 32 X 32 BLIT ARMOR
                If Player(Index).Armor > 0 Then
                    rec.Top = Item(Player(Index).Armor).Pic * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' OTHER 32 X 32 BLIT HELMET
                If Player(Index).Helmet > 0 Then
                    rec.Top = Item(Player(Index).Helmet).Pic * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' OTHER 32 X 32 BLIT LEGS
                If Player(Index).legs > 0 Then
                    rec.Top = Item(Player(Index).legs).Pic * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' OTHER 32 X 32 IF DIR <> UP
                If GetPlayerDir(Index) <> DIR_UP Then

                    ' OTHER 32 X 32 BLIT SHIELD IF DIR <> UP
                    If Player(Index).Shield > 0 Then
                        rec.Top = Item(Player(Index).Shield).Pic * PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' OTHER 32 X 32 BLIT WEAPON IF DIR <> UP
                    If Player(Index).Weapon > 0 Then
                        rec.Top = Item(Player(Index).Weapon).Pic * PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' OTHER 32 X 32 BLIT NECKLACE IF DIR <> UP
                    If Player(Index).Necklace > 0 Then
                        rec.Top = Item(Player(Index).Necklace).Pic * PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                End If

            ' OTHER 32 X 32 NON PAPERDOLL
            Else

                ' OTHER 32 X 32 BLIT NON-PD SPRITE
                If 0 + CUSTOM_PLAYERS = 0 Then
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    rec.Top = Player(Index).head * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_PlayerHead, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.Top = Player(Index).body * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_PlayerBody, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.Top = Player(Index).leg * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_PlayerLegs, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            End If

        End If
    Else
        ' 96 X 96
        If PaperDoll = 1 And GetPlayerPaperdoll(Index) = 1 Then

            rec.Left = (GetPlayerDir(Index) * 3 + Anim) * 96
            rec.Right = rec.Left + 96

            If Index = MyIndex Then
                X = Val(NewX + sx) - PIC_X
                y = Val(NewY + sx) + PIC_Y

                ' PLAYER 96 X 96 IF DIR = UP
                If GetPlayerDir(MyIndex) = DIR_UP Then

                    ' PLAYER 96 X 96 BLIT SHIELD IF DIR = UP
                    If Player(MyIndex).Shield > 0 Then
                        rec.Top = Item(Player(MyIndex).Shield).Pic * 96 + PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' PLAYER 96 X 96 BLIT WEAPON IF DIR = UP
                    If Player(MyIndex).Weapon > 0 Then
                        rec.Top = Item(Player(MyIndex).Weapon).Pic * 96 + PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' PLAYER 96 X 96 BLIT NECKLACE IF DIR = UP
                    If Player(MyIndex).Necklace > 0 Then
                        rec.Top = Item(Player(MyIndex).Necklace).Pic * 96 + PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' PLAYER 96 X 96 BLIT SHIELD IF DIR = UP
                    If Player(MyIndex).Shield > 0 Then
                        rec.Top = Item(Player(MyIndex).Shield).Pic * 96 + PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                End If

                If CUSTOM_PLAYERS = 0 Then
                    ' PLAYER 96 X 96 BLIT SPRITE
                    rec.Top = GetPlayerSprite(MyIndex) * 96 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    ' PLAYER 96 X 96 BLIT SPRITE
                    rec.Top = Player(Index).head * 96 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, y, DD_PlayerHead, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.Top = Player(Index).body * 96 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, y, DD_PlayerBody, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.Top = Player(Index).leg * 96 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, y, DD_PlayerLegs, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' PLAYER 96 X 96 BLIT ARMOR
                If Player(MyIndex).Armor > 0 Then
                    rec.Top = Item(Player(MyIndex).Armor).Pic * 96 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' PLAYER 96 X 96 BLIT LEGS
                If Player(MyIndex).legs > 0 Then
                    rec.Top = Item(Player(MyIndex).legs).Pic * 96 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' PLAYER 96 X 96 BLIT HELMET
                If Player(MyIndex).Helmet > 0 Then
                    rec.Top = Item(Player(MyIndex).Helmet).Pic * 96 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' PLAYER 96 X 96 DIR <> UP
                If GetPlayerDir(MyIndex) <> DIR_UP Then

                    ' PLAYER 96 X 96 BLIT SHIELD IF DIR <> UP
                    If Player(MyIndex).Shield > 0 Then
                        rec.Top = Item(Player(MyIndex).Shield).Pic * 96 + PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' PLAYER 96 X 96 BLIT WEAPON IF DIR <> UP
                    If Player(MyIndex).Weapon > 0 Then
                        rec.Top = Item(Player(MyIndex).Weapon).Pic * 96 + PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' PLAYER 96 X 96 BLIT NECKLACE IF DIR <> UP
                    If Player(MyIndex).Necklace > 0 Then
                        rec.Top = Item(Player(MyIndex).Necklace).Pic * 96 + PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                End If


            ' 96 X 96 IF OTHER PLAYER
            Else

                X = Val(GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset) - PIC_X
                y = Val(GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset) + PIC_Y

                ' IF BLIT IS OFFSCREEN ADJUST THE Y VALUE
                ' 11111 If y < 0 Then
                ' rec.tOp = rec.tOp + (y * -1)
                ' y = 0
                ' End If

                ' OTHER 96 X 96 IF DIR = UP
                If GetPlayerDir(Index) = DIR_UP Then

                    ' OTHER 96 X 96 BLIT SHIELD IF DIR = UP
                    If Player(Index).Shield > 0 Then
                        rec.Top = Item(Player(Index).Shield).Pic * 96 + PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' OTHER 96 X 96 BLIT WEAPON IF DIR = UP
                    If Player(Index).Weapon > 0 Then
                        rec.Top = Item(Player(Index).Weapon).Pic * 96 + PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' OTHER 96 X 96 BLIT NECKLACE IF DIR = UP
                    If Player(Index).Necklace > 0 Then
                        rec.Top = Item(Player(Index).Necklace).Pic * 96 + PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                End If

                ' OTHER 96 X 96 BLIT SPRITE
                If 0 + CUSTOM_PLAYERS = 0 Then
                    rec.Top = GetPlayerSprite(Index) * 96 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    rec.Top = Player(Index).head * 96 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_PlayerHead, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.Top = Player(Index).body * 96 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_PlayerBody, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.Top = Player(Index).leg * 96 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_PlayerLegs, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' OTHER 96 X 96 BLIT ARMOR
                If Player(Index).Armor > 0 Then
                    rec.Top = Item(Player(Index).Armor).Pic * 96 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' OTHER 96 X 96 BLIT LEGS
                If Player(Index).legs > 0 Then
                    rec.Top = Item(Player(Index).legs).Pic * 96 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' OTHER 96 X 96 BLIT HELMET
                If Player(Index).Helmet > 0 Then
                    rec.Top = Item(Player(Index).Helmet).Pic * 96 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' OTHER 96 X 96 IF DIR <> UP
                If GetPlayerDir(Index) <> DIR_UP Then

                    ' OTHER 96 X 96 BLIT SHIELD IF DIR <> UP
                    If Player(Index).Shield > 0 Then
                        rec.Top = Item(Player(Index).Shield).Pic * 96 + PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' OTHER 96 X 96 BLIT NECKLACE IF DIR <> UP
                    If Player(Index).Necklace > 0 Then
                        rec.Top = Item(Player(Index).Necklace).Pic * 96 + PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' 'OTHER 96 X 96 BLIT WEAPON IF DIR <> UP
                    If Player(Index).Weapon > 0 Then
                        rec.Top = Item(Player(Index).Weapon).Pic * 96 + PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                End If

            ' END OF PAPERDOLL FOR 96 X 96
            End If

        ' IF 96 X 96 AND NO PAPERDOLL
        Else
            rec.Left = (GetPlayerDir(Index) * 3 + Anim) * 96
            rec.Right = rec.Left + 96

            ' PLAYER 96 X 96
            If Index = MyIndex Then
                X = NewX + sx
                y = NewY + sx

                If 0 + CUSTOM_PLAYERS = 0 Then
                    ' PLAYER 96 X 96 BLIT SPRITE
                    rec.Top = GetPlayerSprite(MyIndex) * 96 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    ' PLAYER 96 X 96 BLIT SPRITE
                    rec.Top = Player(Index).head * 96 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, y, DD_PlayerHead, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.Top = Player(Index).body * 96 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, y, DD_PlayerBody, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.Top = Player(Index).leg * 96 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, y, DD_PlayerLegs, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

            ' OTHER 96 X 96
            Else
                X = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset
                y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset

                ' ADJUST IF OFF EDGE OF SCREEN
                ' If y < 0 Then
                ' rec.tOp = rec.tOp + (y * -1)
                ' y = 0
                ' End If

                ' OTHER 96 X 96 BLIT SPRITE
                If 0 + CUSTOM_PLAYERS = 0 Then
                    rec.Top = GetPlayerSprite(Index) * 96 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    rec.Top = Player(Index).head * 96 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_PlayerHead, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.Top = Player(Index).body * 96 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_PlayerBody, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.Top = Player(Index).leg * 96 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_PlayerLegs, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            End If

        ' END OF 96 X 96
        End If
    End If

End Sub
Sub BltPlayerTop(ByVal Index As Long)
    Dim Anim As Byte
    Dim X As Long
    Dim y As Long
    Dim yMod As Long
    Dim AttackSpeed As Long
    Dim rec As DXVBLib.RECT

    If SpriteSize = 1 Then
        If GetPlayerWeaponSlot(Index) > 0 Then
            AttackSpeed = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AttackSpeed
        Else
            AttackSpeed = 1000
        End If

        ' Only used if ever want to switch to blt rather then bltfast
        ' With rec_pos
        ' .Top = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset
        ' .Bottom = .Top + PIC_Y
        ' .Left = GetPlayerX(Index) * PIC_X + Player(Index).xOffset
        ' .Right = .Left + PIC_X
        ' End With

        ' Check for animation
        Anim = 0
        If Player(Index).Attacking = 0 Then
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    If (Player(Index).yOffset < PIC_Y / 2) Then
                        Anim = 1
                    End If
                Case DIR_DOWN
                    If (Player(Index).yOffset < PIC_Y / 2 * -1) Then
                        Anim = 1
                    End If
                Case DIR_LEFT
                    If (Player(Index).xOffset < PIC_Y / 2) Then
                        Anim = 1
                    End If
                Case DIR_RIGHT
                    If (Player(Index).xOffset < PIC_Y / 2 * -1) Then
                        Anim = 1
                    End If
            End Select
        Else
            If Player(Index).AttackTimer + Int(AttackSpeed / 2) > GetTickCount Then
                Anim = 2
            End If
        End If

        ' Check to see if we want to stop making him attack
        If Player(Index).AttackTimer + AttackSpeed < GetTickCount Then
            Player(Index).Attacking = 0
            'Player(Index).AttackTimer = 0
        End If

        If PaperDoll = 1 And GetPlayerPaperdoll(Index) = 1 Then
            rec.Left = (GetPlayerDir(Index) * 3 + Anim) * PIC_X
            rec.Right = rec.Left + PIC_X

            If Index = MyIndex Then
                X = NewX + sx
                y = NewY + sx - 32
                
                ' Fixing "Player head disspear" bug - Emblem
                ' It was caused by trying to blt to a invalid location.
                If y < 0 Then
                    yMod = y
                    y = 0
                End If

                If 0 + CUSTOM_PLAYERS = 0 Then
                    rec.Top = GetPlayerSprite(Index) * 64 - yMod
                    rec.Bottom = rec.Top + PIC_Y + yMod

                    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    rec.Top = GetPlayerHead(Index) * 64 - yMod
                    rec.Bottom = rec.Top + PIC_Y + yMod

                    Call DD_BackBuffer.BltFast(X, y, DD_PlayerHead, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                If GetPlayerDir(Index) = DIR_UP Then
                    If Player(MyIndex).Shield > 0 Then
                        rec.Top = Item(Player(MyIndex).Shield).Pic * 64 - yMod
                        rec.Bottom = rec.Top + PIC_Y + yMod
                        Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                    If Player(MyIndex).Weapon > 0 Then
                        rec.Top = Item(Player(MyIndex).Weapon).Pic * 64 - yMod
                        rec.Bottom = rec.Top + PIC_Y + yMod
                        Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                    If Player(MyIndex).Necklace > 0 Then
                        rec.Top = Item(Player(MyIndex).Necklace).Pic * 64 - yMod
                        rec.Bottom = rec.Top + PIC_Y + yMod
                        Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                End If

                If Player(MyIndex).Armor > 0 Then
                    rec.Top = Item(Player(MyIndex).Armor).Pic * 64 - yMod
                    rec.Bottom = rec.Top + PIC_Y + yMod
                    Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                If Player(MyIndex).legs > 0 Then
                    rec.Top = Item(Player(MyIndex).legs).Pic * 64 - yMod
                    rec.Bottom = rec.Top + PIC_Y + yMod
                    Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                If Player(MyIndex).Helmet > 0 Then
                    rec.Top = Item(Player(MyIndex).Helmet).Pic * 64 - yMod
                    rec.Bottom = rec.Top + PIC_Y + yMod
                    Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                If GetPlayerDir(Index) <> DIR_UP Then
                    If Player(MyIndex).Shield > 0 Then
                        rec.Top = Item(Player(MyIndex).Shield).Pic * 64 - yMod
                        rec.Bottom = rec.Top + PIC_Y + yMod
                        Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                    If Player(MyIndex).Necklace > 0 Then
                        rec.Top = Item(Player(MyIndex).Necklace).Pic * 64 - yMod
                        rec.Bottom = rec.Top + PIC_Y + yMod
                        Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                    If Player(MyIndex).Weapon > 0 Then
                        rec.Top = Item(Player(MyIndex).Weapon).Pic * 64 - yMod
                        rec.Bottom = rec.Top + PIC_Y + yMod
                        Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                End If


            Else
                X = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset
                y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset - 32

                If y < 0 Then
                    yMod = y
                    y = 0
                End If

                If 0 + CUSTOM_PLAYERS = 0 Then
                    rec.Top = GetPlayerSprite(Index) * 64 - yMod
                    rec.Bottom = rec.Top + PIC_Y + yMod

                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    rec.Top = GetPlayerHead(Index) * 64 - yMod
                    rec.Bottom = rec.Top + PIC_Y + yMod

                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_PlayerHead, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                If GetPlayerDir(Index) = DIR_UP Then
                    If Player(Index).Shield > 0 Then
                        rec.Top = Item(Player(Index).Shield).Pic * 64 - yMod
                        rec.Bottom = rec.Top + PIC_Y + yMod
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                    If Player(Index).Necklace > 0 Then
                        rec.Top = Item(Player(Index).Necklace).Pic * 64 - yMod
                        rec.Bottom = rec.Top + PIC_Y + yMod
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                    If Player(Index).Weapon > 0 Then
                        rec.Top = Item(Player(Index).Weapon).Pic * 64 - yMod
                        rec.Bottom = rec.Top + PIC_Y + yMod
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                End If

                If Player(Index).Armor > 0 Then
                    rec.Top = Item(Player(Index).Armor).Pic * 64 - yMod
                    rec.Bottom = rec.Top + PIC_Y + yMod
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                If Player(Index).legs > 0 Then
                    rec.Top = Item(Player(Index).legs).Pic * 64 - yMod
                    rec.Bottom = rec.Top + PIC_Y + yMod
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                If Player(Index).Helmet > 0 Then
                    rec.Top = Item(Player(Index).Helmet).Pic * 64 - yMod
                    rec.Bottom = rec.Top + PIC_Y + yMod
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                If GetPlayerDir(Index) <> DIR_UP Then
                    If Player(Index).Shield > 0 Then
                        rec.Top = Item(Player(Index).Shield).Pic * 64 - yMod
                        rec.Bottom = rec.Top + PIC_Y + yMod
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                    If Player(Index).Necklace > 0 Then
                        rec.Top = Item(Player(Index).Necklace).Pic * 64 - yMod
                        rec.Bottom = rec.Top + PIC_Y + yMod
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                    If Player(Index).Weapon > 0 Then
                        rec.Top = Item(Player(Index).Weapon).Pic * 64 - yMod
                        rec.Bottom = rec.Top + PIC_Y + yMod
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                End If
            End If
        Else
            If Index = MyIndex Then
                X = NewX + sx
                y = NewY + sx - 32
                
            Else
                X = (GetPlayerX(Index) - NewPlayerX) * PIC_X + sx + Player(Index).xOffset - NewXOffset
                y = (GetPlayerY(Index) - NewPlayerY) * PIC_Y + sx + Player(Index).yOffset - NewYOffset - 32
            End If
            
            If y < 0 Then
                yMod = y
                y = 0
            End If
            
            rec.Top = GetPlayerSprite(Index) * 64 - yMod
            rec.Bottom = rec.Top + PIC_Y + yMod
            rec.Left = (GetPlayerDir(Index) * 3 + Anim) * PIC_X
            rec.Right = rec.Left + PIC_X
            
            If 0 + CUSTOM_PLAYERS = 0 Then
                Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            Else
                rec.Top = Player(Index).head * 64 - yMod
                rec.Bottom = rec.Top + PIC_Y + yMod
                Call DD_BackBuffer.BltFast(X, y, DD_PlayerHead, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        End If
    Else

        If GetPlayerWeaponSlot(Index) > 0 Then
            AttackSpeed = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AttackSpeed
        Else
            AttackSpeed = 1000
        End If

        ' Only used if ever want to switch to blt rather then bltfast
        ' With rec_pos
        ' .Top = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset
        ' .Bottom = .Top + PIC_Y
        ' .Left = GetPlayerX(Index) * PIC_X + Player(Index).xOffset
        ' .Right = .Left + 96
        ' End With

        ' Check for animation
        Anim = 0
        If Player(Index).Attacking = 0 Then
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    If (Player(Index).yOffset < PIC_Y / 2) Then
                        Anim = 1
                    End If
                Case DIR_DOWN
                    If (Player(Index).yOffset < PIC_Y / 2 * -1) Then
                        Anim = 1
                    End If
                Case DIR_LEFT
                    If (Player(Index).xOffset < PIC_Y / 2) Then
                        Anim = 1
                    End If
                Case DIR_RIGHT
                    If (Player(Index).xOffset < PIC_Y / 2 * -1) Then
                        Anim = 1
                    End If
            End Select
        Else
            If Player(Index).AttackTimer + Int(AttackSpeed / 2) > GetTickCount Then
                Anim = 2
            End If
        End If

        ' Check to see if we want to stop making him attack
        If Player(Index).AttackTimer + AttackSpeed < GetTickCount Then
            Player(Index).Attacking = 0
            Player(Index).AttackTimer = 0
        End If

        If PaperDoll = 1 And GetPlayerPaperdoll(Index) = 1 Then
            rec.Left = (GetPlayerDir(Index) * 3 + Anim) * PIC_X
            rec.Right = rec.Left + 96

            If Index = MyIndex Then
                X = NewX + sx - PIC_X
                y = NewY + sx - 32
                If y < 0 Then
                    yMod = y
                    y = 0
                End If

                If 0 + CUSTOM_PLAYERS = 0 Then
                    rec.Top = GetPlayerSprite(Index) * 96 - yMod
                    rec.Bottom = rec.Top + PIC_Y + yMod

                    Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    rec.Top = GetPlayerHead(Index) * 96 - yMod
                    rec.Bottom = rec.Top + PIC_Y + yMod

                    Call DD_BackBuffer.BltFast(X, y, DD_PlayerHead, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                If GetPlayerDir(Index) = DIR_UP Then
                    If Player(MyIndex).Shield > 0 Then
                        rec.Top = Item(Player(MyIndex).Shield).Pic * 96 - yMod
                        rec.Bottom = rec.Top + PIC_Y + yMod
                        Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                    If Player(MyIndex).Weapon > 0 Then
                        rec.Top = Item(Player(MyIndex).Weapon).Pic * 96 - yMod
                        rec.Bottom = rec.Top + PIC_Y + yMod
                        Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                    If Player(MyIndex).Necklace > 0 Then
                        rec.Top = Item(Player(MyIndex).Necklace).Pic * 96 - yMod
                        rec.Bottom = rec.Top + PIC_Y + yMod
                        Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                End If

                If Player(MyIndex).Armor > 0 Then
                    rec.Top = Item(Player(MyIndex).Armor).Pic * 96 - yMod
                    rec.Bottom = rec.Top + PIC_Y + -yMod
                    Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                If Player(MyIndex).legs > 0 Then
                    rec.Top = Item(Player(MyIndex).legs).Pic * 96 - yMod
                    rec.Bottom = rec.Top + PIC_Y + yMod
                    Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                If Player(MyIndex).Helmet > 0 Then
                    rec.Top = Item(Player(MyIndex).Helmet).Pic * 96 - yMod
                    rec.Bottom = rec.Top + PIC_Y + yMod
                    Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                If GetPlayerDir(Index) <> DIR_UP Then
                    If Player(MyIndex).Shield > 0 Then
                        rec.Top = Item(Player(MyIndex).Shield).Pic * 96 - yMod
                        rec.Bottom = rec.Top + PIC_Y + yMod
                        Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                    If Player(MyIndex).Necklace > 0 Then
                        rec.Top = Item(Player(MyIndex).Necklace).Pic * 96 - yMod
                        rec.Bottom = rec.Top + PIC_Y + yMod
                        Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                    If Player(MyIndex).Weapon > 0 Then
                        rec.Top = Item(Player(MyIndex).Weapon).Pic * 96 - yMod
                        rec.Bottom = rec.Top + PIC_Y + yMod
                        Call DD_BackBuffer.BltFast(X, y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                End If

            Else
                X = (GetPlayerX(Index) - NewPlayerX) * PIC_X + sx + Player(Index).xOffset - PIC_X - NewXOffset
                y = (GetPlayerY(Index) - NewPlayerY) * PIC_Y + sx + Player(Index).yOffset - 32 + PIC_Y - NewXOffset
                
                If y < 0 Then
                    yMod = y
                    y = 0
                End If

                If GetPlayerDir(Index) = DIR_UP Then
                    If Player(Index).Shield > 0 Then
                        rec.Top = Item(Player(Index).Shield).Pic * 96 - yMod
                        rec.Bottom = rec.Top + PIC_Y + yMod
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                    If Player(Index).Necklace > 0 Then
                        rec.Top = Item(Player(Index).Necklace).Pic * 96 - yMod
                        rec.Bottom = rec.Top + PIC_Y + yMod
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                    If Player(Index).Weapon > 0 Then
                        rec.Top = Item(Player(Index).Weapon).Pic * 96 - yMod
                        rec.Bottom = rec.Top + PIC_Y + yMod
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                End If

                rec.Top = GetPlayerSprite(Index) * 96 - yMod
                rec.Bottom = rec.Top + PIC_Y + yMod

                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

                If Player(Index).Armor > 0 Then
                    rec.Top = Item(Player(Index).Armor).Pic * 96 - yMod
                    rec.Bottom = rec.Top + PIC_Y + yMod
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                If Player(Index).legs > 0 Then
                    rec.Top = Item(Player(Index).legs).Pic * 96 - yMod
                    rec.Bottom = rec.Top + PIC_Y + yMod
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                If Player(Index).Helmet > 0 Then
                    rec.Top = Item(Player(Index).Helmet).Pic * 96 - yMod
                    rec.Bottom = rec.Top + PIC_Y + yMod
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                If GetPlayerDir(Index) <> DIR_UP Then
                    If Player(Index).Shield > 0 Then
                        rec.Top = Item(Player(Index).Shield).Pic * 96 - yMod
                        rec.Bottom = rec.Top + PIC_Y + yMod
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                    If Player(Index).Necklace > 0 Then
                        rec.Top = Item(Player(Index).Necklace).Pic * 96 - yMod
                        rec.Bottom = rec.Top + PIC_Y + yMod
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                    If Player(Index).Weapon > 0 Then
                        rec.Top = Item(Player(Index).Weapon).Pic * 96 - yMod
                        rec.Bottom = rec.Top + PIC_Y + yMod
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                End If
            End If
        Else
            

            If Index = MyIndex Then
                X = NewX + sx + PIC_X
                y = NewY + sx - 32 - PIC_Y

            Else
                X = (GetPlayerX(Index) - NewPlayerX) * PIC_X + Player(Index).xOffset - NewXOffset
                y = (GetPlayerY(Index) - NewPlayerY) * PIC_Y + Player(Index).yOffset - NewXOffset - 32

                'Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
            If y < 0 Then
                yMod = y
                y = 0
            End If
            
            rec.Top = GetPlayerSprite(Index) * 96 - yMod
            rec.Bottom = rec.Top + PIC_Y + yMod
            rec.Left = (GetPlayerDir(Index) * 3 + Anim) * PIC_X
            rec.Right = rec.Left + 96

            If 0 + CUSTOM_PLAYERS = 0 Then
                Call DD_BackBuffer.BltFast(X, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            Else
                rec.Top = Player(Index).head * 96 - yMod
                rec.Bottom = rec.Top + PIC_Y + yMod
                Call DD_BackBuffer.BltFast(X, y, DD_PlayerHead, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        End If
    End If
End Sub

Sub BltMapNPCName(ByVal Index As Long)
    Dim TextX As Long
    Dim TextY As Long

    If Npc(MapNpc(Index).num).Big = 0 And Npc(MapNpc(Index).num).SpriteSize = 0 Then
        TextX = MapNpc(Index).X * PIC_X + sx + MapNpc(Index).xOffset + CLng(PIC_X / 2) - ((Len(Trim$(Npc(MapNpc(Index).num).Name)) / 2) * 8)
        TextY = MapNpc(Index).y * PIC_Y + sx + MapNpc(Index).yOffset - CLng(PIC_Y / 2) - 4
    Else
        TextX = MapNpc(Index).X * PIC_X + sx + MapNpc(Index).xOffset + CLng(PIC_X / 2) - ((Len(Trim$(Npc(MapNpc(Index).num).Name)) / 2) * 8)
        TextY = MapNpc(Index).y * PIC_Y + sx + MapNpc(Index).yOffset - CLng(PIC_Y / 2) - 32
    End If

    Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, Trim$(Npc(MapNpc(Index).num).Name), vbWhite)
End Sub

Sub BltNpcBody(ByVal MapNpcNum As Long)
    Dim Anim As Byte
    Dim X As Long
    Dim y As Long
    Dim modify As Long
    Dim rec As DXVBLib.RECT

' Only used if ever want to switch to blt rather then bltfast
' With rec_pos
' .Top = MapNpc(MapNpcNum).Y * PIC_Y + MapNpc(MapNpcNum).YOffset
' .Bottom = .Top + PIC_Y
' .Left = MapNpc(MapNpcNum).X * PIC_X + MapNpc(MapNpcNum).XOffset
' .Right = .Left + PIC_X
' End With

    ' Check for animation
    Anim = 0
    If MapNpc(MapNpcNum).Attacking = 0 Then
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                If (MapNpc(MapNpcNum).yOffset < PIC_Y / 2) Then
                    Anim = 1
                End If
            Case DIR_DOWN
                If (MapNpc(MapNpcNum).yOffset < PIC_Y / 2 * -1) Then
                    Anim = 1
                End If
            Case DIR_LEFT
                If (MapNpc(MapNpcNum).xOffset < PIC_Y / 2) Then
                    Anim = 1
                End If
            Case DIR_RIGHT
                If (MapNpc(MapNpcNum).xOffset < PIC_Y / 2 * -1) Then
                    Anim = 1
                End If
        End Select
    Else
        If MapNpc(MapNpcNum).AttackTimer + 500 > GetTickCount Then
            Anim = 2
        End If
    End If

    ' Check to see if we want to stop making him attack
    If MapNpc(MapNpcNum).AttackTimer + 1000 < GetTickCount Then
        MapNpc(MapNpcNum).Attacking = 0
        MapNpc(MapNpcNum).AttackTimer = 0
    End If

    If Npc(MapNpc(MapNpcNum).num).Big = 1 Then
        rec.Top = Npc(MapNpc(MapNpcNum).num).Sprite * 64 + 32
        rec.Bottom = rec.Top + 32
        rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * 64
        rec.Right = rec.Left + 64

        X = MapNpc(MapNpcNum).X * 32 + sx - 16 + MapNpc(MapNpcNum).xOffset
        y = MapNpc(MapNpcNum).y * 32 + sx + MapNpc(MapNpcNum).yOffset

        If y < 0 Then
            modify = -y
            rec.Top = rec.Top + modify
            rec.Bottom = rec.Top + 32
            y = 0
        End If

        If X < 0 Then
            ' rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * 64 + 16
            ' modify = -X
            ' rec.Left = rec.Left + modify - 16
            ' rec.Right = rec.Left + 48
            ' X = 0
            modify = -X
            rec.Left = rec.Left + modify
            rec.Right = rec.Left + 48
            X = 0
        End If

        If 32 + X >= (MAX_MAPX * 32) Then
            modify = X - (MAX_MAPX * 32)
            rec.Left = rec.Left + modify + 16
            rec.Right = rec.Left + 32 - modify
        End If

        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_BigSpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    Else

        If Npc(MapNpc(MapNpcNum).num).SpriteSize = 1 Then
            rec.Top = Npc(MapNpc(MapNpcNum).num).Sprite * 64 + 32
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
            rec.Right = rec.Left + PIC_X

            X = MapNpc(MapNpcNum).X * PIC_X + sx + MapNpc(MapNpcNum).xOffset
            y = MapNpc(MapNpcNum).y * PIC_Y + sx + MapNpc(MapNpcNum).yOffset

' Check if its out of bounds because of the offset

            If y < 0 Then
                rec.Top = rec.Top + (y * -1)
                y = 0
            End If

            ' Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        Else
            rec.Top = Npc(MapNpc(MapNpcNum).num).Sprite * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
            rec.Right = rec.Left + PIC_X

            X = MapNpc(MapNpcNum).X * PIC_X + sx + MapNpc(MapNpcNum).xOffset
            y = MapNpc(MapNpcNum).y * PIC_Y + sx + MapNpc(MapNpcNum).yOffset

            ' Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
End Sub

Sub BltNpcTop(ByVal MapNpcNum As Long)
    Dim Anim As Byte
    Dim X As Long
    Dim y As Long
    Dim NPC_number As Long
    Dim modify As Long
    Dim rec As DXVBLib.RECT

    ' Get the NPC number
    NPC_number = MapNpc(MapNpcNum).num

    If Npc(NPC_number).Big = 0 Then
        If Npc(MapNpc(MapNpcNum).num).SpriteSize = 0 Then
            Exit Sub
        End If
    End If

' Only used if ever want to switch to blt rather then bltfast
' With rec_pos
' .Top = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).yOffset
' .Bottom = .Top + PIC_Y
' .Left = MapNpc(MapNpcNum).x * PIC_X + MapNpc(MapNpcNum).xOffset
' .Right = .Left + PIC_X
' End With

    ' Check for animation
    Anim = 0
    If MapNpc(MapNpcNum).Attacking = 0 Then
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                If (MapNpc(MapNpcNum).yOffset < PIC_Y / 2) Then
                    Anim = 1
                End If
            Case DIR_DOWN
                If (MapNpc(MapNpcNum).yOffset < PIC_Y / 2 * -1) Then
                    Anim = 1
                End If
            Case DIR_LEFT
                If (MapNpc(MapNpcNum).xOffset < PIC_Y / 2) Then
                    Anim = 1
                End If
            Case DIR_RIGHT
                If (MapNpc(MapNpcNum).xOffset < PIC_Y / 2 * -1) Then
                    Anim = 1
                End If
        End Select
    Else
        If MapNpc(MapNpcNum).AttackTimer + 500 > GetTickCount Then
            Anim = 2
        End If
    End If

    ' Check to see if we want to stop making him attack
    If MapNpc(MapNpcNum).AttackTimer + 1000 < GetTickCount Then
        MapNpc(MapNpcNum).Attacking = 0
        MapNpc(MapNpcNum).AttackTimer = 0
    End If

    If Npc(MapNpc(MapNpcNum).num).Big = 0 Then
        rec.Top = Npc(MapNpc(MapNpcNum).num).Sprite * 64
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
        rec.Right = rec.Left + PIC_X

        X = MapNpc(MapNpcNum).X * PIC_X + sx + MapNpc(MapNpcNum).xOffset
        y = MapNpc(MapNpcNum).y * PIC_Y + sx + MapNpc(MapNpcNum).yOffset - 32

        ' Check if its out of bounds because of the offset
        If y < 0 Then
            rec.Top = rec.Top + (y * -1)
            y = 0
        End If

        ' Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    Else
        rec.Top = Npc(MapNpc(MapNpcNum).num).Sprite * PIC_Y

        rec.Top = Npc(MapNpc(MapNpcNum).num).Sprite * 64
        rec.Bottom = rec.Top + 32
        rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * 64
        rec.Right = rec.Left + 64

        X = MapNpc(MapNpcNum).X * 32 + sx - 16 + MapNpc(MapNpcNum).xOffset
        y = MapNpc(MapNpcNum).y * 32 + sx - 32 + MapNpc(MapNpcNum).yOffset

        If y < 0 Then
            modify = -y
            rec.Top = rec.Top + modify
            rec.Bottom = rec.Top + 32
            y = 0
        End If

        If X < 0 Then
            ' rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * 64 + 16
            ' modify = -X
            ' rec.Left = rec.Left + modify - 16
            ' rec.Right = rec.Left + 48
            ' X = 0
            modify = -X
            rec.Left = rec.Left + modify
            rec.Right = rec.Left + 48
            X = 0
        End If

        If 32 + X >= (MAX_MAPX * 32) Then
            modify = X - (MAX_MAPX * 32)
            rec.Left = rec.Left + modify + 16
            rec.Right = rec.Left + 32 - modify
        End If

        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_BigSpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
End Sub
Sub BltPlayerName(ByVal Index As Long)
    Dim TextX As Long
    Dim TextY As Long
    Dim Color As Long

    If Player(Index).Color <> 0 Then
        If Player(Index).Color > 16 Then
            Exit Sub
        Else
            Color = QBColor(Val(Player(Index).Color - 1))
        End If
    Else
        ' Check access level
        If GetPlayerPK(Index) = NO Then
            Color = QBColor(YELLOW)
            Select Case GetPlayerAccess(Index)
                Case 0
                    Color = QBColor(BROWN)
                Case 1
                    Color = QBColor(DARKGREY)
                Case 2
                    Color = QBColor(CYAN)
                Case 3
                    Color = QBColor(BLUE)
                Case 4
                    Color = QBColor(PINK)
            End Select
        Else
            Color = QBColor(BRIGHTRED)
        End If
    End If

    If SpriteSize = 1 Then
        If Index = MyIndex Then
            If DISPLAY_LEVEL >= 1 Then
                TextX = NewX + sx + Int(PIC_X / 2) - ((Len(GetPlayerName(MyIndex) & " Lvl: " & GetPlayerLevel(Index)) / 2) * 8)
            Else
                TextX = NewX + sx + Int(PIC_X / 2) - ((Len(GetPlayerName(MyIndex)) / 2) * 8)
            End If

            TextY = NewY + sx - 50
            If DISPLAY_LEVEL >= 1 Then
                Call DrawText(TexthDC, TextX, TextY, GetPlayerName(MyIndex) & " Lvl: " & GetPlayerLevel(Index), Color)
            Else
                Call DrawText(TexthDC, TextX, TextY, GetPlayerName(MyIndex), Color)
            End If
        Else
            ' Draw name
            If DISPLAY_LEVEL >= 1 Then
                TextX = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset + Int(PIC_X / 2) - ((Len(GetPlayerName(Index) & " Lvl: " & GetPlayerLevel(Index)) / 2) * 8)
            Else
                TextX = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset + Int(PIC_X / 2) - ((Len(GetPlayerName(Index)) / 2) * 8)
            End If

            TextY = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset - Int(PIC_Y / 2) - 32

            If DISPLAY_LEVEL >= 1 Then
                Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, GetPlayerName(Index) & " Lvl: " & GetPlayerLevel(Index), Color)
            Else
                Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, GetPlayerName(Index), Color)
            End If
        End If
    Else
        If SpriteSize = 2 Then
            If Index = MyIndex Then
                If DISPLAY_LEVEL >= 1 Then
                    TextX = NewX + sx + Int(PIC_X / 2) - ((Len(GetPlayerName(MyIndex) & " Lvl: " & GetPlayerLevel(Index)) / 2) * 8)
                Else
                    TextX = NewX + sx + Int(PIC_X / 2) - ((Len(GetPlayerName(MyIndex)) / 2) * 8)
                End If

                TextY = NewY + sx - 50
                If DISPLAY_LEVEL >= 1 Then
                    Call DrawText(TexthDC, TextX, TextY - PIC_Y, GetPlayerName(MyIndex) & " Lvl: " & GetPlayerLevel(Index), Color)
                Else
                    Call DrawText(TexthDC, TextX, TextY - PIC_Y, GetPlayerName(MyIndex), Color)
                End If
            Else
                ' Draw name
                If DISPLAY_LEVEL >= 1 Then
                    TextX = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset + Int(PIC_X / 2) - ((Len(GetPlayerName(Index) & " Lvl: " & GetPlayerLevel(Index)) / 2) * 8)
                Else
                    TextX = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset + Int(PIC_X / 2) - ((Len(GetPlayerName(Index)) / 2) * 8)
                End If

                TextY = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset - Int(PIC_Y / 2) - 32

                If DISPLAY_LEVEL >= 1 Then
                    Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, GetPlayerName(Index) & " Lvl: " & GetPlayerLevel(Index), Color)
                Else
                    Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset - PIC_Y, GetPlayerName(Index), Color)
                End If
            End If
        Else
            If Index = MyIndex Then
                If DISPLAY_LEVEL >= 1 Then
                    TextX = NewX + sx + Int(PIC_X / 2) - ((Len(GetPlayerName(MyIndex) & " Lvl: " & GetPlayerLevel(Index)) / 2) * 8)
                Else
                    TextX = NewX + sx + Int(PIC_X / 2) - ((Len(GetPlayerName(MyIndex)) / 2) * 8)
                End If
                TextY = NewY + sx - Int(PIC_Y / 2)

                If DISPLAY_LEVEL >= 1 Then
                    Call DrawText(TexthDC, TextX, TextY, GetPlayerName(MyIndex) & " Lvl: " & GetPlayerLevel(Index), Color)
                Else
                    Call DrawText(TexthDC, TextX, TextY, GetPlayerName(MyIndex), Color)
                End If
            Else
                ' Draw name
                If DISPLAY_LEVEL >= 1 Then
                    TextX = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset + Int(PIC_X / 2) - ((Len(GetPlayerName(Index) & " Lvl: " & GetPlayerLevel(Index)) / 2) * 8)
                Else
                    TextX = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset + Int(PIC_X / 2) - ((Len(GetPlayerName(Index)) / 2) * 8)
                End If

                TextY = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset - Int(PIC_Y / 2)

                If DISPLAY_LEVEL >= 1 Then
                    Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, GetPlayerName(Index) & " Lvl: " & GetPlayerLevel(Index), Color)
                Else
                    Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, GetPlayerName(Index), Color)
                End If
            End If
        End If
    End If
End Sub


Sub BltPlayerGuildName(ByVal Index As Long)
    Dim TextX As Long
    Dim TextY As Long
    Dim Color As Long

    ' Check access level.
    If GetPlayerPK(Index) = NO Then
        Select Case GetPlayerGuildAccess(Index)
            Case 0
                Color = QBColor(RED)
            Case 1
                Color = QBColor(BRIGHTCYAN)
            Case 2
                Color = QBColor(PINK)
            Case 3
                Color = QBColor(BRIGHTGREEN)
            Case 4
                Color = QBColor(YELLOW)
        End Select
    Else
        Color = QBColor(BRIGHTRED)
    End If

    ' Draw the players guild.
    If Index = MyIndex Then
        TextX = NewX + sx + Int(PIC_X / 2) - ((Len(GetPlayerGuild(MyIndex)) / 2) * 8)

        If SpriteSize = 1 Then
            TextY = NewY + sx - Int(PIC_Y / 4) - 52
        Else
            TextY = NewY + sx - Int(PIC_Y / 4) - 20
        End If

        Call DrawText(TexthDC, TextX, TextY, GetPlayerGuild(MyIndex), Color)
    Else
        TextX = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset + Int(PIC_X / 2) - ((Len(GetPlayerGuild(Index)) / 2) * 8)

        If SpriteSize = 1 Then
            TextY = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset - Int(PIC_Y / 2) - 44
        Else
            TextY = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset - Int(PIC_Y / 2) - 12
        End If

        Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, GetPlayerGuild(Index), Color)
    End If
End Sub

Sub ProcessMovement(ByVal Index As Long)
    ' Check if player is walking, and if so process moving them over
    If Player(Index).Moving = MOVING_WALKING Then
        If Player(Index).Access > 0 Then
            If SS_WALK_SPEED <> 0 Then
                Select Case GetPlayerDir(Index)
                    Case DIR_UP
                        Player(Index).yOffset = Player(Index).yOffset - SS_WALK_SPEED
                    Case DIR_DOWN
                        Player(Index).yOffset = Player(Index).yOffset + SS_WALK_SPEED
                    Case DIR_LEFT
                        Player(Index).xOffset = Player(Index).xOffset - SS_WALK_SPEED
                    Case DIR_RIGHT
                        Player(Index).xOffset = Player(Index).xOffset + SS_WALK_SPEED
                End Select
            Else
                Select Case GetPlayerDir(Index)
                    Case DIR_UP
                        Player(Index).yOffset = Player(Index).yOffset - GM_WALK_SPEED
                    Case DIR_DOWN
                        Player(Index).yOffset = Player(Index).yOffset + GM_WALK_SPEED
                    Case DIR_LEFT
                        Player(Index).xOffset = Player(Index).xOffset - GM_WALK_SPEED
                    Case DIR_RIGHT
                        Player(Index).xOffset = Player(Index).xOffset + GM_WALK_SPEED
                End Select
            End If
        Else
            If SS_WALK_SPEED <> 0 Then
                Select Case GetPlayerDir(Index)
                    Case DIR_UP
                        Player(Index).yOffset = Player(Index).yOffset - SS_WALK_SPEED
                    Case DIR_DOWN
                        Player(Index).yOffset = Player(Index).yOffset + SS_WALK_SPEED
                    Case DIR_LEFT
                        Player(Index).xOffset = Player(Index).xOffset - SS_WALK_SPEED
                    Case DIR_RIGHT
                        Player(Index).xOffset = Player(Index).xOffset + SS_WALK_SPEED
                End Select
            Else
                Select Case GetPlayerDir(Index)
                    Case DIR_UP
                        Player(Index).yOffset = Player(Index).yOffset - WALK_SPEED
                    Case DIR_DOWN
                        Player(Index).yOffset = Player(Index).yOffset + WALK_SPEED
                    Case DIR_LEFT
                        Player(Index).xOffset = Player(Index).xOffset - WALK_SPEED
                    Case DIR_RIGHT
                        Player(Index).xOffset = Player(Index).xOffset + WALK_SPEED
                End Select
            End If
        End If

        ' Check if completed walking over to the next tile
        If (Player(Index).xOffset = 0) And (Player(Index).yOffset = 0) Then
            Player(Index).Moving = 0
        End If
    Else
        ' Check if player is running, and if so process moving them over
        If Player(Index).Moving = MOVING_RUNNING Then
            If GetPlayerSP(Index) > 0 Then
                ' Removed until the server supports SP. [Mellowz]
                ' Call SetPlayerSP(Index, GetPlayerSP(Index) - 1)
                frmMirage.lblSP.Caption = Int(GetPlayerSP(MyIndex) / GetPlayerMaxSP(MyIndex) * 100) & "%"
                If Player(Index).Access > 0 Then
                    If SS_RUN_SPEED <> 0 Then
                        Select Case GetPlayerDir(Index)
                            Case DIR_UP
                                Player(Index).yOffset = Player(Index).yOffset - SS_RUN_SPEED
                            Case DIR_DOWN
                                Player(Index).yOffset = Player(Index).yOffset + SS_RUN_SPEED
                            Case DIR_LEFT
                                Player(Index).xOffset = Player(Index).xOffset - SS_RUN_SPEED
                            Case DIR_RIGHT
                                Player(Index).xOffset = Player(Index).xOffset + SS_RUN_SPEED
                        End Select
                    Else
                        Select Case GetPlayerDir(Index)
                            Case DIR_UP
                                Player(Index).yOffset = Player(Index).yOffset - GM_RUN_SPEED
                            Case DIR_DOWN
                                Player(Index).yOffset = Player(Index).yOffset + GM_RUN_SPEED
                            Case DIR_LEFT
                                Player(Index).xOffset = Player(Index).xOffset - GM_RUN_SPEED
                            Case DIR_RIGHT
                                Player(Index).xOffset = Player(Index).xOffset + GM_RUN_SPEED
                        End Select
                    End If
                Else
                    If SS_RUN_SPEED <> 0 Then
                        Select Case GetPlayerDir(Index)
                            Case DIR_UP
                                Player(Index).yOffset = Player(Index).yOffset - SS_RUN_SPEED
                            Case DIR_DOWN
                                Player(Index).yOffset = Player(Index).yOffset + SS_RUN_SPEED
                            Case DIR_LEFT
                                Player(Index).xOffset = Player(Index).xOffset - SS_RUN_SPEED
                            Case DIR_RIGHT
                                Player(Index).xOffset = Player(Index).xOffset + SS_RUN_SPEED
                        End Select
                    Else
                        Select Case GetPlayerDir(Index)
                            Case DIR_UP
                                Player(Index).yOffset = Player(Index).yOffset - RUN_SPEED
                            Case DIR_DOWN
                                Player(Index).yOffset = Player(Index).yOffset + RUN_SPEED
                            Case DIR_LEFT
                                Player(Index).xOffset = Player(Index).xOffset - RUN_SPEED
                            Case DIR_RIGHT
                                Player(Index).xOffset = Player(Index).xOffset + RUN_SPEED
                        End Select
                    End If
                End If
            Else
                ' Call AddText("You are to tired to run.", Blue)
                Player(Index).Moving = MOVING_WALKING
            End If


            ' Check if completed walking over to the next tile
            If (Player(Index).xOffset = 0) And (Player(Index).yOffset = 0) Then
                Player(Index).Moving = 0
            End If
        End If
    End If
End Sub

Sub ProcessNpcMovement(ByVal MapNpcNum As Long)
    ' Check if npc is walking, and if so process moving them over
    If MapNpc(MapNpcNum).Moving = MOVING_WALKING Then
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                MapNpc(MapNpcNum).yOffset = MapNpc(MapNpcNum).yOffset - WALK_SPEED
            Case DIR_DOWN
                MapNpc(MapNpcNum).yOffset = MapNpc(MapNpcNum).yOffset + WALK_SPEED
            Case DIR_LEFT
                MapNpc(MapNpcNum).xOffset = MapNpc(MapNpcNum).xOffset - WALK_SPEED
            Case DIR_RIGHT
                MapNpc(MapNpcNum).xOffset = MapNpc(MapNpcNum).xOffset + WALK_SPEED
        End Select

        ' Check if completed walking over to the next tile
        If (MapNpc(MapNpcNum).xOffset = 0) And (MapNpc(MapNpcNum).yOffset = 0) Then
            MapNpc(MapNpcNum).Moving = 0
        End If
    End If
End Sub

Public Sub Event_SignText()
    Dim PlayerMap As Long
    Dim PlayerX As Long
    Dim PlayerY As Long

    ' Get the players coordinates.
    PlayerMap = GetPlayerMap(MyIndex)
    PlayerX = GetPlayerX(MyIndex)
    PlayerY = GetPlayerY(MyIndex)

    ' Check to make sure we don't check out-of-bounds.
    If Not (PlayerY - 1) > -1 Then Exit Sub

    ' Check if the attribute on the tile is a sign.
    If Not Map(PlayerMap).Tile(PlayerX, PlayerY - 1).Type = TILE_TYPE_SIGN Then Exit Sub

    ' Check if the player is facing north.
    If Not GetPlayerDir(MyIndex) = DIR_UP Then Exit Sub

    ' Display the sign to the player.
    Call AddText("You read the following:", BLACK)

    ' Display the first line of the sign.
    If Trim$(Map(PlayerMap).Tile(PlayerX, PlayerY - 1).String1) <> vbNullString Then
        Call AddText(Trim$(Map(PlayerMap).Tile(PlayerX, PlayerY - 1).String1), GREY)
    End If

    ' Display the second line of the sign.
    If Trim$(Map(PlayerMap).Tile(PlayerX, PlayerY - 1).String2) <> vbNullString Then
        Call AddText(Trim$(Map(PlayerMap).Tile(PlayerX, PlayerY - 1).String2), GREY)
    End If

    ' Display the third line of the sign.
    If Trim$(Map(PlayerMap).Tile(PlayerX, PlayerY - 1).String3) <> vbNullString Then
        Call AddText(Trim$(Map(PlayerMap).Tile(PlayerX, PlayerY - 1).String3), GREY)
    End If
End Sub

Sub HandleKeypresses(ByVal KeyAscii As Integer)
    Dim ChatText As String
    Dim Name As String
    Dim I As Long

    ' Get the message or command.
    MyText = frmMirage.txtMyTextBox.Text

    ' Handle when the player presses the return key
    If KeyAscii = vbKeyReturn Then

        ' Reset the message control.
        frmMirage.txtMyTextBox.Text = vbNullString

        ' Check for the sign event.
        Call Event_SignText

        ' Broadcast message
        If Mid$(MyText, 1, 1) = "'" Then
            ChatText = Mid$(MyText, 2, Len(MyText) - 1)
            If Len(Trim$(ChatText)) > 0 Then
                Call BroadcastMsg(ChatText)
            End If
            MyText = vbNullString
            Exit Sub
        End If

        ' Emote message
        If Mid$(MyText, 1, 1) = "-" Then
            ChatText = Mid$(MyText, 2, Len(MyText) - 1)
            If Len(Trim$(ChatText)) > 0 Then
                Call EmoteMsg(ChatText)
            End If
            MyText = vbNullString
            Exit Sub
        End If

        ' Player message
        If Mid$(MyText, 1, 1) = "!" Then
            ChatText = Mid$(MyText, 2, Len(MyText) - 1)
            Name = vbNullString

            ' Get the desired player from the user text
            For I = 1 To Len(ChatText)
                If Mid$(ChatText, I, 1) <> " " Then
                    Name = Name & Mid$(ChatText, I, 1)
                Else
                    Exit For
                End If
            Next I

            ' Make sure they are actually sending something
            If Len(ChatText) - I > 0 Then
                ChatText = Mid$(ChatText, I + 1, Len(ChatText) - I)

                ' Send the message to the player
                Call PlayerMsg(ChatText, Name)
            Else
                Call AddText("Usage: !playername msghere", AlertColor)
            End If
            MyText = vbNullString
            Exit Sub
        End If

        ' // Commands //
        ' Verification User
        If LCase$(Mid$(MyText, 1, 5)) = "/info" Then
            ChatText = Mid$(MyText, 7, Len(MyText) - 5)

            If LenB(ChatText) <> 0 Then
                Call SendData(POut.GetStats & SEP_CHAR & ChatText & END_CHAR)
            Else
                Call AddText("Please enter a player name.", BRIGHTRED)
            End If

            MyText = vbNullString
            Exit Sub
        End If

        ' Whos Online
        If LCase$(Mid$(MyText, 1, 4)) = "/who" Then
            Call SendWhosOnline
            MyText = vbNullString
            Exit Sub
        End If
        
        If LCase$(Mid$(MyText, 1, 6)) = "/where" Then
            Call AddText("Map: " & GetPlayerMap(MyIndex) & "; X: " & GetPlayerX(MyIndex) & "; Y: " & GetPlayerY(MyIndex), GREY)
            MyText = vbNullString
            Exit Sub
        End If

        ' Checking fps
        If Mid$(MyText, 1, 4) = "/fps" Then
            If BFPS = False Then
                BFPS = True
            Else
                BFPS = False
            End If
            MyText = vbNullString
            Exit Sub
        End If

        ' Show inventory
        If LCase$(Mid$(MyText, 1, 4)) = "/inv" Then
            frmMirage.picInventory.Visible = True
            MyText = vbNullString
            Exit Sub
        End If

        ' Request stats
        If LCase$(Mid$(MyText, 1, 6)) = "/stats" Then
            Call SendData(POut.GetStats & SEP_CHAR & GetPlayerName(MyIndex) & END_CHAR)
            MyText = vbNullString
            Exit Sub
        End If

        ' Refresh Player
        If LCase$(Mid$(MyText, 1, 8)) = "/refresh" Then
            Call SendData(POut.Refresh & END_CHAR)
            MyText = vbNullString
            Exit Sub
        End If

        ' Decline Chat
        If LCase$(Mid$(MyText, 1, 12)) = "/chatdecline" Then
            Call SendData(POut.DeclineChat & END_CHAR)
            MyText = vbNullString
            Exit Sub
        End If

        ' Accept Chat
        If LCase$(Mid$(MyText, 1, 5)) = "/chat" Then
            Call SendData(POut.AcceptChat & END_CHAR)
            MyText = vbNullString
            Exit Sub
        End If

        If LCase$(Mid$(MyText, 1, 6)) = "/trade" Then
            ' Make sure they are actually sending something
            If Len(MyText) > 7 Then
                ChatText = Mid$(MyText, 8, Len(MyText) - 7)
                Call SendTradeRequest(ChatText)
            Else
                Call AddText("Usage: /trade playernamehere", AlertColor)
            End If
            MyText = vbNullString
            Exit Sub
        End If

        ' Accept Trade
        If LCase$(Mid$(MyText, 1, 7)) = "/accept" Then
            Call SendAcceptTrade
            MyText = vbNullString
            Exit Sub
        End If

        ' Decline Trade
        If LCase$(Mid$(MyText, 1, 8)) = "/decline" Then
            Call SendDeclineTrade
            MyText = vbNullString
            Exit Sub
        End If

        ' Party request
        If LCase$(Mid$(MyText, 1, 8)) = "/pcreate" Then
            ' Make sure they are actually sending something
            If Len(MyText) > 9 Then
                ChatText = Mid$(MyText, 10, Len(MyText) - 9)
                Call SendData(POut.PartyCreate & SEP_CHAR & ChatText & END_CHAR)
            Else
                Call AddText("Usage: /pcreate <Party Name>", AlertColor)
            End If
            MyText = vbNullString
            Exit Sub
        End If
        
        If LCase$(Mid$(MyText, 1, 9)) = "/pdisband" Then
            Call SendData(POut.PartyDisband & END_CHAR)
            MyText = vbNullString
            Exit Sub
        End If
        
        If LCase$(Mid$(MyText, 1, 8)) = "/pinvite" Then
            ' Make sure they are actually sending something
            If Len(MyText) > 9 Then
                ChatText = Mid$(MyText, 10, Len(MyText) - 9)
                Call SendData(POut.PartyInvite & SEP_CHAR & ChatText & END_CHAR)
            Else
                Call AddText("Usage: /pinvite <Player Name>", AlertColor)
            End If
            MyText = vbNullString
            Exit Sub
        End If

        ' Join party
        If LCase$(Mid$(MyText, 1, 8)) = "/paccept" Then
            Call SendData(POut.PartyInviteAccept & END_CHAR)
            MyText = vbNullString
            Exit Sub
        End If

        ' Join party
        If LCase$(Mid$(MyText, 1, 9)) = "/pdecline" Then
            Call SendData(POut.PartyInviteDecline & END_CHAR)
            MyText = vbNullString
            Exit Sub
        End If

        ' Leave party
        If LCase$(Mid$(MyText, 1, 7)) = "/pleave" Then
            Call SendData(POut.PartyLeave & END_CHAR)
            MyText = vbNullString
            Exit Sub
        End If

        If LCase$(Mid$(MyText, 1, 8)) = "/pleader" Then
            ' Make sure they are actually sending something
            If Len(MyText) > 9 Then
                ChatText = Mid$(MyText, 10, Len(MyText) - 9)
                Call SendData(POut.PartyChangeLeader & SEP_CHAR & ChatText & END_CHAR)
            Else
                Call AddText("Usage: /pleader <Player Name>", AlertColor)
            End If
            MyText = vbNullString
            Exit Sub
        End If

        ' // Moniter Admin Commands //
        If GetPlayerAccess(MyIndex) > 0 Then
            ' weather command
            If LCase$(Mid$(MyText, 1, 8)) = "/weather" Then
                If Len(MyText) > 8 Then
                    MyText = Mid$(MyText, 9, Len(MyText) - 8)
                    If IsNumeric(MyText) = True Then
                        Call SendData(POut.Weather & SEP_CHAR & Val(MyText) & END_CHAR)
                    Else
                        If Trim$(LCase$(MyText)) = "none" Then
                            I = 0
                        End If
                        If Trim$(LCase$(MyText)) = "rain" Then
                            I = 1
                        End If
                        If Trim$(LCase$(MyText)) = "snow" Then
                            I = 2
                        End If
                        If Trim$(LCase$(MyText)) = "thunder" Then
                            I = 3
                        End If
                        Call SendData(POut.Weather & SEP_CHAR & I & END_CHAR)
                    End If
                End If
                MyText = vbNullString
                Exit Sub
            End If

            ' Kicking a player
            If LCase$(Mid$(MyText, 1, 5)) = "/kick" Then
                If Len(MyText) > 6 Then
                    MyText = Mid$(MyText, 7, Len(MyText) - 6)
                    Call SendKick(MyText)
                End If
                MyText = vbNullString
                Exit Sub
            End If

            ' Global Message
            If Mid$(MyText, 1, 1) = "'" Then
                ChatText = Mid$(MyText, 2, Len(MyText) - 1)
                If Len(Trim$(ChatText)) > 0 Then
                    Call GlobalMsg(ChatText)
                End If
                MyText = vbNullString
                Exit Sub
            End If

            ' Admin Message
            If Mid$(MyText, 1, 1) = "=" Then
                ChatText = Mid$(MyText, 2, Len(MyText) - 1)
                If Len(Trim$(ChatText)) > 0 Then
                    Call AdminMsg(ChatText)
                End If
                MyText = vbNullString
                Exit Sub
            End If
        End If

        ' // Mapper Admin Commands //
        If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
            ' Location
            If Mid$(MyText, 1, 4) = "/loc" Then
                If BLoc = False Then
                    BLoc = True
                Else
                    BLoc = False
                End If
                MyText = vbNullString
                Exit Sub
            End If

            ' Map Editor
            If LCase$(Mid$(MyText, 1, 10)) = "/mapeditor" Then
                Call SendRequestEditMap
                MyText = vbNullString
                Exit Sub
            End If

            ' Map report
            If LCase$(Mid$(MyText, 1, 10)) = "/mapreport" Then
                Call SendData(POut.MapReport & END_CHAR)
                MyText = vbNullString
                Exit Sub
            End If

            ' Setting sprite
            If LCase$(Mid$(MyText, 1, 10)) = "/setsprite" Then
                If Len(MyText) > 11 Then
                    ' Get sprite #
                    MyText = Mid$(MyText, 12, Len(MyText) - 11)

                    Call SendSetPlayerSprite(GetPlayerName(MyIndex), Val(MyText))
                End If
                MyText = vbNullString
                Exit Sub
            End If

            ' Setting player sprite
            If LCase$(Mid$(MyText, 1, 16)) = "/setplayersprite" Then
                If Len(MyText) > 19 Then
                    I = Val(Mid$(MyText, 17, 1))

                    MyText = Mid$(MyText, 18, Len(MyText) - 17)
                    Call SendSetPlayerSprite(I, Val(MyText))
                End If
                MyText = vbNullString
                Exit Sub
            End If

            ' Respawn request
            If Mid$(MyText, 1, 8) = "/respawn" Then
                Call SendMapRespawn
                MyText = vbNullString
                Exit Sub
            End If

            ' MOTD change
            If Mid$(MyText, 1, 5) = "/motd" Then
                If Len(MyText) > 6 Then
                    MyText = Mid$(MyText, 7, Len(MyText) - 6)
                    If Trim$(MyText) <> vbNullString Then
                        Call SendMOTDChange(MyText)
                    End If
                End If
                MyText = vbNullString
                Exit Sub
            End If

            ' Check the ban list
            If Mid$(MyText, 1, 3) = "/banlist" Then
                Call SendBanList
                MyText = vbNullString
                Exit Sub
            End If

            ' Banning a player
            If LCase$(Mid$(MyText, 1, 4)) = "/ban" Then
                If Len(MyText) > 5 Then
                    MyText = Mid$(MyText, 6, Len(MyText) - 5)
                    Call SendBan(MyText)
                    MyText = vbNullString
                End If
                Exit Sub
            End If
        End If

        ' // Developer Admin Commands //
        If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
            ' Editing item request
            If Mid$(MyText, 1, 9) = "/edititem" Then
                Call SendRequestEditItem
                MyText = vbNullString
                Exit Sub
            End If

            ' Day/Night
            If LCase$(Mid$(MyText, 1, 9)) = "/daynight" Then
                Call SendData(POut.DayNight & END_CHAR)
                MyText = vbNullString
                Exit Sub
            End If

            ' Editing emoticon request
            If Mid$(MyText, 1, 13) = "/editemoticon" Then
                Call SendRequestEditEmoticon
                MyText = vbNullString
                Exit Sub
            End If

            ' Editing emoticon request
            If Mid$(MyText, 1, 12) = "/editelement" Then
                Call SendRequestEditElement
                MyText = vbNullString
                Exit Sub
            End If

            ' Editing arrow request
            If Mid$(MyText, 1, 13) = "/editarrow" Then
                Call SendRequestEditArrow
                MyText = vbNullString
                Exit Sub
            End If

            ' Editing npc request
            If Mid$(MyText, 1, 8) = "/editnpc" Then
                Call SendRequestEditNPC
                MyText = vbNullString
                Exit Sub
            End If

            ' Editing shop request
            If Mid$(MyText, 1, 9) = "/editshop" Then
                Call SendRequestEditShop
                MyText = vbNullString
                Exit Sub
            End If

            ' Editing spell request
            If LCase$(Trim$(MyText)) = "/editspell" Then
                Call SendRequestEditSpell
                MyText = vbNullString
                Exit Sub
            End If
        End If

        ' // Creator Admin Commands //
        If GetPlayerAccess(MyIndex) >= ADMIN_CREATOR Then
            ' Giving another player access
            If LCase$(Mid$(MyText, 1, 10)) = "/setaccess" Then
                ' Get access #
                I = Val(Mid$(MyText, 12, 1))

                MyText = Mid$(MyText, 14, Len(MyText) - 13)

                Call SendSetAccess(MyText, I)
                MyText = vbNullString
                Exit Sub
            End If

            ' Reload Scripts
            If LCase$(Trim$(MyText)) = "/reload" Then
                Call SendReloadScripts
                MyText = vbNullString
                Exit Sub
            End If

            ' Ban destroy
            If LCase$(Mid$(MyText, 1, 15)) = "/destroybanlist" Then
                Call SendBanDestroy
                MyText = vbNullString
                Exit Sub
            End If
        End If

        ' Tell them its not a valid command
        If Left$(Trim$(MyText), 1) = "/" Then
            For I = 0 To MAX_EMOTICONS
                If Trim$(Emoticons(I).Command) = Trim$(MyText) And Trim$(Emoticons(I).Command) <> "/" Then
                    Call SendData(POut.CheckEmoticons & SEP_CHAR & I & END_CHAR)
                    MyText = vbNullString
                    Exit Sub
                End If
            Next I
            Call SendData(POut.CheckCommands & SEP_CHAR & MyText & END_CHAR)
            MyText = vbNullString
            Exit Sub
        End If

        ' Say message
        If Len(Trim$(MyText)) > 0 Then
            Call SayMsg(MyText)
        End If
        MyText = vbNullString
        Exit Sub
    End If
End Sub

Sub CheckMapGetItem()
    If GetTickCount > Player(MyIndex).MapGetTimer + 250 And Trim$(MyText) = vbNullString Then
        Player(MyIndex).MapGetTimer = GetTickCount
        Call SendData(POut.MapGetItem & END_CHAR)
    End If
End Sub

Sub CheckAttack()
    Dim AttackSpeed As Long

    If GetPlayerWeaponSlot(MyIndex) > 0 Then
        AttackSpeed = Item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AttackSpeed
    Else
        AttackSpeed = 1000
    End If

    If ControlDown Then
        If Player(MyIndex).AttackTimer + AttackSpeed < GetTickCount Then
            If Player(MyIndex).Attacking = 0 Then
                Player(MyIndex).Attacking = 1
                Player(MyIndex).AttackTimer = GetTickCount
                Call SendData(POut.Attack & END_CHAR)
            End If
        End If
    End If
End Sub

Sub CheckInput(ByVal KeyState As Byte, ByVal KeyCode As Integer, ByVal Shift As Integer)
    If Not GettingMap Then
        If KeyState = 1 Then
            If KeyCode = vbKeyReturn Then
                Call CheckMapGetItem
            End If

            If KeyCode = vbKeyControl Then
                ControlDown = True
            End If

            If KeyCode = vbKeyUp Then
                DirUp = True
                DirDown = False
                DirLeft = False
                DirRight = False
            End If

            If KeyCode = vbKeyDown Then
                DirUp = False
                DirDown = True
                DirLeft = False
                DirRight = False
            End If

            If KeyCode = vbKeyLeft Then
                DirUp = False
                DirDown = False
                DirLeft = True
                DirRight = False
            End If

            If KeyCode = vbKeyRight Then
                DirUp = False
                DirDown = False
                DirLeft = False
                DirRight = True
            End If

            If KeyCode = vbKeyShift Then
                ShiftDown = True
            End If
        End If
    End If
End Sub

Function IsTryingToMove() As Boolean
    If DirUp Or DirDown Or DirLeft Or DirRight Then
        IsTryingToMove = True
    End If
End Function

Function CanMove() As Boolean
    Dim I As Long
    Dim X As Long
    Dim y As Long

    CanMove = True

    If Player(MyIndex).Moving <> 0 Then
        CanMove = False
        Exit Function
    End If

    ' Make sure they haven't just casted a spell
    If Player(MyIndex).CastedSpell = YES Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            Player(MyIndex).CastedSpell = NO
        Else
            CanMove = False
            Exit Function
        End If
    End If

    X = GetPlayerX(MyIndex)
    y = GetPlayerY(MyIndex)

    If DirUp Then
        Call SetPlayerDir(MyIndex, DIR_UP)
        y = y - 1
    ElseIf DirDown Then
        Call SetPlayerDir(MyIndex, DIR_DOWN)
        y = y + 1
    ElseIf DirLeft Then
        Call SetPlayerDir(MyIndex, DIR_LEFT)
        X = X - 1
    Else
        Call SetPlayerDir(MyIndex, DIR_RIGHT)
        X = X + 1
    End If

    If y < 0 Then
        If Map(GetPlayerMap(MyIndex)).Up > 0 Then
            Call SendPlayerRequestNewMap(DIR_UP)
            GettingMap = True
        End If
        CanMove = False
        Exit Function
    ElseIf y > MAX_MAPY Then
        If Map(GetPlayerMap(MyIndex)).Down > 0 Then
            Call SendPlayerRequestNewMap(DIR_DOWN)
            GettingMap = True
        End If
        CanMove = False
        Exit Function
    ElseIf X < 0 Then
        If Map(GetPlayerMap(MyIndex)).Left > 0 Then
            Call SendPlayerRequestNewMap(DIR_LEFT)
            GettingMap = True
        End If
        CanMove = False
        Exit Function
    ElseIf X > MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Right > 0 Then
            Call SendPlayerRequestNewMap(DIR_RIGHT)
            GettingMap = True
        End If
        CanMove = False
        Exit Function
    End If

    If Not GetPlayerDir(MyIndex) = LAST_DIR Then
        LAST_DIR = GetPlayerDir(MyIndex)
        Call SendPlayerDir
    End If

    If Map(GetPlayerMap(MyIndex)).Tile(X, y).Type = TILE_TYPE_BLOCKED Or Map(GetPlayerMap(MyIndex)).Tile(X, y).Type = TILE_TYPE_SIGN Or Map(GetPlayerMap(MyIndex)).Tile(X, y).Type = TILE_TYPE_ROOFBLOCK Then
        CanMove = False
        Exit Function
    End If

    If Map(GetPlayerMap(MyIndex)).Tile(X, y).Type = TILE_TYPE_CBLOCK Then
        If Map(GetPlayerMap(MyIndex)).Tile(X, y).Data1 = Player(MyIndex).Class Then
            Exit Function
        End If
        If Map(GetPlayerMap(MyIndex)).Tile(X, y).Data2 = Player(MyIndex).Class Then
            Exit Function
        End If
        If Map(GetPlayerMap(MyIndex)).Tile(X, y).Data3 = Player(MyIndex).Class Then
            Exit Function
        End If
        CanMove = False
    End If

    If Map(GetPlayerMap(MyIndex)).Tile(X, y).Type = TILE_TYPE_GUILDBLOCK And Map(GetPlayerMap(MyIndex)).Tile(X, y).String1 <> GetPlayerGuild(MyIndex) Then
        CanMove = False
    End If

    If Map(GetPlayerMap(MyIndex)).Tile(X, y).Type = TILE_TYPE_KEY Or Map(GetPlayerMap(MyIndex)).Tile(X, y).Type = TILE_TYPE_DOOR Then
        If TempTile(X, y).DoorOpen = NO Then
            CanMove = False
            Exit Function
        End If
    End If

    If Map(GetPlayerMap(MyIndex)).Tile(X, y).Type = TILE_TYPE_WALKTHRU Then
        Exit Function
    Else
        For I = 1 To MAX_PLAYERS
            If IsPlaying(I) Then
                If GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                    If GetPlayerX(I) = X Then
                        If GetPlayerY(I) = y Then
                            CanMove = False
                            Exit Function
                        End If
                    End If
                End If
            End If
        Next I
    End If

    For I = 1 To MAX_MAP_NPCS
        If MapNpc(I).num > 0 Then
            If MapNpc(I).X = X Then
                If MapNpc(I).y = y Then
                    CanMove = False
                    Exit Function
                End If
            End If
        End If
    Next I
End Function

Sub CheckMovement()
    If Not GettingMap Then
        If IsTryingToMove Then
            If CanMove Then
                ' Check if player has the shift key down for running
                If ShiftDown Then
                    Player(MyIndex).Moving = MOVING_RUNNING
                Else
                    Player(MyIndex).Moving = MOVING_WALKING
                End If

                Select Case GetPlayerDir(MyIndex)
                    Case DIR_UP
                        Call SendPlayerMove
                        Player(MyIndex).yOffset = PIC_Y
                        Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) - 1)

                    Case DIR_DOWN
                        Call SendPlayerMove
                        Player(MyIndex).yOffset = PIC_Y * -1
                        Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)

                    Case DIR_LEFT
                        Call SendPlayerMove
                        Player(MyIndex).xOffset = PIC_X
                        Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)

                    Case DIR_RIGHT
                        Call SendPlayerMove
                        Player(MyIndex).xOffset = PIC_X * -1
                        Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) + 1)
                End Select

                ' Gotta check :)
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_WARP Then
                    GettingMap = True
                End If
            End If
        End If
    End If
End Sub

Public Function FindPlayer(ByVal Name As String) As Long
    Dim I As Long

    ' Capitalize and trim the player name.
    Name = UCase$(Trim$(Name))

    ' Loop through all of the players.
    For I = 1 To MAX_PLAYERS

        ' Check if the player is playing.
        If IsPlaying(I) Then

            ' Makes sure we don't check names too small, in bytes.
            If LenB(GetPlayerName(I)) >= LenB(Name) Then

                ' Compare the names in upper-case.
                If UCase$(GetPlayerName(I)) = Name Then
                    FindPlayer = I
                    Exit Function
                End If
            End If
        End If
    Next I
End Function

Public Sub UpdateTradeInventory()
    Dim I As Long

    frmPlayerTrade.PlayerInv1.Clear

    For I = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, I) > 0 And GetPlayerInvItemNum(MyIndex, I) <= MAX_ITEMS Then
            If Item(GetPlayerInvItemNum(MyIndex, I)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, I)).Stackable = 1 Then
                frmPlayerTrade.PlayerInv1.addItem I & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, I)).Name) & " (" & GetPlayerInvItemValue(MyIndex, I) & ")"
            Else
                If GetPlayerWeaponSlot(MyIndex) = I Or GetPlayerArmorSlot(MyIndex) = I Or GetPlayerHelmetSlot(MyIndex) = I Or GetPlayerShieldSlot(MyIndex) = I Or GetPlayerLegsSlot(MyIndex) = I Or GetPlayerRingSlot(MyIndex) = I Or GetPlayerNecklaceSlot(MyIndex) = I Then
                    frmPlayerTrade.PlayerInv1.addItem I & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, I)).Name) & " (worn)"
                Else
                    frmPlayerTrade.PlayerInv1.addItem I & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, I)).Name)
                End If
            End If
        Else
            frmPlayerTrade.PlayerInv1.addItem "<Nothing>"
        End If
    Next I

    frmPlayerTrade.PlayerInv1.ListIndex = 0
End Sub

Sub PlayerSearch(Button As Integer, Shift As Integer, X As Long, y As Long)
    If CurX >= 0 And CurX <= MAX_MAPX Then
        If CurY >= 0 And CurY <= MAX_MAPY Then
            ' Disabled until we get a better movement system. [Mellowz]
            ' Call MoveCharacter(CurX, CurY)
            Call SendData(POut.Search & SEP_CHAR & CurX & SEP_CHAR & CurY & END_CHAR)
        End If
    End If
End Sub

Public Sub UpdateVisInv()
    Dim sRECT As DXVBLib.RECT
    Dim dRECT As DXVBLib.RECT

    ' Reset all of the equipment slots.
    frmMirage.ShieldImage.Picture = LoadPicture()
    frmMirage.WeaponImage.Picture = LoadPicture()
    frmMirage.HelmetImage.Picture = LoadPicture()
    frmMirage.ArmorImage.Picture = LoadPicture()
    frmMirage.LegsImage.Picture = LoadPicture()
    frmMirage.RingImage.Picture = LoadPicture()
    frmMirage.NecklaceImage.Picture = LoadPicture()

    ' Define the destination rectangle.
    dRECT.Top = 0
    dRECT.Bottom = dRECT.Top + PIC_Y
    dRECT.Left = 0
    dRECT.Right = dRECT.Left + PIC_X

    ' Draw the shield slot.
    If GetPlayerShieldSlot(MyIndex) > 0 Then
        sRECT.Top = Int(Item(GetPlayerInvItemNum(MyIndex, GetPlayerShieldSlot(MyIndex))).Pic / 6) * PIC_Y
        sRECT.Bottom = sRECT.Top + PIC_Y
        sRECT.Left = (Item(GetPlayerInvItemNum(MyIndex, GetPlayerShieldSlot(MyIndex))).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, GetPlayerShieldSlot(MyIndex))).Pic / 6) * 6) * PIC_X
        sRECT.Right = sRECT.Left + PIC_X
        
        Call DD_ItemSurf.BltToDC(frmMirage.ShieldImage.hDC, sRECT, dRECT)
        
        frmMirage.ShieldImage.Refresh
    End If

    ' Draw the weapon slot.
    If GetPlayerWeaponSlot(MyIndex) > 0 Then
        sRECT.Top = Int(Item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).Pic / 6) * PIC_Y
        sRECT.Bottom = sRECT.Top + PIC_Y
        sRECT.Left = (Item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).Pic / 6) * 6) * PIC_X
        sRECT.Right = sRECT.Left + PIC_X
        
        Call DD_ItemSurf.BltToDC(frmMirage.WeaponImage.hDC, sRECT, dRECT)
        
        frmMirage.WeaponImage.Refresh
    End If

    ' Draw the helmet slot.
    If GetPlayerHelmetSlot(MyIndex) > 0 Then
        sRECT.Top = Int(Item(GetPlayerInvItemNum(MyIndex, GetPlayerHelmetSlot(MyIndex))).Pic / 6) * PIC_Y
        sRECT.Bottom = sRECT.Top + PIC_Y
        sRECT.Left = (Item(GetPlayerInvItemNum(MyIndex, GetPlayerHelmetSlot(MyIndex))).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, GetPlayerHelmetSlot(MyIndex))).Pic / 6) * 6) * PIC_X
        sRECT.Right = sRECT.Left + PIC_X

        Call DD_ItemSurf.BltToDC(frmMirage.HelmetImage.hDC, sRECT, dRECT)

        frmMirage.HelmetImage.Refresh
    End If

    ' Draw the armor slot.
    If GetPlayerArmorSlot(MyIndex) > 0 Then
        sRECT.Top = Int(Item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).Pic / 6) * PIC_Y
        sRECT.Bottom = sRECT.Top + PIC_Y
        sRECT.Left = (Item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).Pic / 6) * 6) * PIC_X
        sRECT.Right = sRECT.Left + PIC_X
        
        Call DD_ItemSurf.BltToDC(frmMirage.ArmorImage.hDC, sRECT, dRECT)

        frmMirage.ArmorImage.Refresh
    End If

    ' Draw the leg slot.
    If GetPlayerLegsSlot(MyIndex) > 0 Then
        sRECT.Top = Int(Item(GetPlayerInvItemNum(MyIndex, GetPlayerLegsSlot(MyIndex))).Pic / 6) * PIC_Y
        sRECT.Bottom = sRECT.Top + PIC_Y
        sRECT.Left = (Item(GetPlayerInvItemNum(MyIndex, GetPlayerLegsSlot(MyIndex))).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, GetPlayerLegsSlot(MyIndex))).Pic / 6) * 6) * PIC_X
        sRECT.Right = sRECT.Left + PIC_X

        Call DD_ItemSurf.BltToDC(frmMirage.LegsImage.hDC, sRECT, dRECT)
        
        frmMirage.LegsImage.Refresh
    End If

    ' Draw the ring slot.
    If GetPlayerRingSlot(MyIndex) > 0 Then
        sRECT.Top = Int(Item(GetPlayerInvItemNum(MyIndex, GetPlayerRingSlot(MyIndex))).Pic / 6) * PIC_Y
        sRECT.Bottom = sRECT.Top + PIC_Y
        sRECT.Left = (Item(GetPlayerInvItemNum(MyIndex, GetPlayerRingSlot(MyIndex))).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, GetPlayerRingSlot(MyIndex))).Pic / 6) * 6) * PIC_X
        sRECT.Right = sRECT.Left + PIC_X

        Call DD_ItemSurf.BltToDC(frmMirage.RingImage.hDC, sRECT, dRECT)
        
        frmMirage.RingImage.Refresh
    End If

    ' Draw the necklace slot.
    If GetPlayerNecklaceSlot(MyIndex) > 0 Then
        sRECT.Top = Int(Item(GetPlayerInvItemNum(MyIndex, GetPlayerNecklaceSlot(MyIndex))).Pic / 6) * PIC_Y
        sRECT.Bottom = sRECT.Top + PIC_Y
        sRECT.Left = (Item(GetPlayerInvItemNum(MyIndex, GetPlayerNecklaceSlot(MyIndex))).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, GetPlayerNecklaceSlot(MyIndex))).Pic / 6) * 6) * PIC_X
        sRECT.Right = sRECT.Left + PIC_X

        Call DD_ItemSurf.BltToDC(frmMirage.NecklaceImage.hDC, sRECT, dRECT)

        frmMirage.NecklaceImage.Refresh
    End If

    ' Hide all of the worn indicators.
    frmMirage.EquipS(0).Visible = False
    frmMirage.EquipS(1).Visible = False
    frmMirage.EquipS(2).Visible = False
    frmMirage.EquipS(3).Visible = False
    frmMirage.EquipS(4).Visible = False
    frmMirage.EquipS(5).Visible = False
    frmMirage.EquipS(6).Visible = False

    ' Draw the weapon slot.
    If GetPlayerWeaponSlot(MyIndex) > 0 Then
        frmMirage.EquipS(0).Top = frmMirage.picInv(GetPlayerWeaponSlot(MyIndex)).Top - 2
        frmMirage.EquipS(0).Left = frmMirage.picInv(GetPlayerWeaponSlot(MyIndex)).Left - 2
        frmMirage.EquipS(0).Visible = True
    End If

    ' Draw the armor slot.
    If GetPlayerArmorSlot(MyIndex) > 0 Then
        frmMirage.EquipS(1).Top = frmMirage.picInv(GetPlayerArmorSlot(MyIndex)).Top - 2
        frmMirage.EquipS(1).Left = frmMirage.picInv(GetPlayerArmorSlot(MyIndex)).Left - 2
        frmMirage.EquipS(1).Visible = True
    End If

    ' Draw the helmet slot.
    If GetPlayerHelmetSlot(MyIndex) > 0 Then
        frmMirage.EquipS(2).Top = frmMirage.picInv(GetPlayerHelmetSlot(MyIndex)).Top - 2
        frmMirage.EquipS(2).Left = frmMirage.picInv(GetPlayerHelmetSlot(MyIndex)).Left - 2
        frmMirage.EquipS(2).Visible = True
    End If

    ' Draw the shield slot.
    If GetPlayerShieldSlot(MyIndex) > 0 Then
        frmMirage.EquipS(3).Top = frmMirage.picInv(GetPlayerShieldSlot(MyIndex)).Top - 2
        frmMirage.EquipS(3).Left = frmMirage.picInv(GetPlayerShieldSlot(MyIndex)).Left - 2
        frmMirage.EquipS(3).Visible = True
    End If

    ' Draw the leg slot.
    If GetPlayerLegsSlot(MyIndex) > 0 Then
        frmMirage.EquipS(4).Top = frmMirage.picInv(GetPlayerLegsSlot(MyIndex)).Top - 2
        frmMirage.EquipS(4).Left = frmMirage.picInv(GetPlayerLegsSlot(MyIndex)).Left - 2
        frmMirage.EquipS(4).Visible = True
    End If

    ' Draw the ring slot.
    If GetPlayerRingSlot(MyIndex) > 0 Then
        frmMirage.EquipS(5).Top = frmMirage.picInv(GetPlayerRingSlot(MyIndex)).Top - 2
        frmMirage.EquipS(5).Left = frmMirage.picInv(GetPlayerRingSlot(MyIndex)).Left - 2
        frmMirage.EquipS(5).Visible = True
    End If

    ' Draw the necklace slot.
    If GetPlayerNecklaceSlot(MyIndex) > 0 Then
        frmMirage.EquipS(6).Top = frmMirage.picInv(GetPlayerNecklaceSlot(MyIndex)).Top - 2
        frmMirage.EquipS(6).Left = frmMirage.picInv(GetPlayerNecklaceSlot(MyIndex)).Left - 2
        frmMirage.EquipS(6).Visible = True
    End If
End Sub

Sub SendGameTime()
    Call SendData(POut.GMTime & SEP_CHAR & GameTime & END_CHAR)
End Sub

Sub UpdateBank()
    Dim I As Long

    frmBank.lstInventory.Clear
    frmBank.lstBank.Clear

    For I = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, I) > 0 Then
            If Item(GetPlayerInvItemNum(MyIndex, I)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, I)).Stackable = 1 Then
                frmBank.lstInventory.addItem I & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, I)).Name) & " (" & GetPlayerInvItemValue(MyIndex, I) & ")"
            Else
                frmBank.lstInventory.addItem I & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, I)).Name)
            End If
        Else
            frmBank.lstInventory.addItem I & "> Empty"
        End If
    Next I

    For I = 1 To MAX_BANK
        If GetPlayerBankItemNum(MyIndex, I) > 0 Then
            If Item(GetPlayerBankItemNum(MyIndex, I)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerBankItemNum(MyIndex, I)).Stackable = 1 Then
                frmBank.lstBank.addItem I & "> " & Trim$(Item(GetPlayerBankItemNum(MyIndex, I)).Name) & " (" & GetPlayerBankItemValue(MyIndex, I) & ")"
            Else
                frmBank.lstBank.addItem I & "> " & Trim$(Item(GetPlayerBankItemNum(MyIndex, I)).Name)
            End If
        Else
            frmBank.lstBank.addItem I & "> Empty"
        End If
    Next I

    frmBank.lstBank.ListIndex = 0
    frmBank.lstInventory.ListIndex = 0
End Sub

Sub UseItem()
    ' Send the item they want to use to the server.
    Call SendUseItem(Inventory)

    ' Update any equipment that has been worn or taken off.
    Call UpdateVisInv
End Sub

Sub DropItem()
    Dim GoldAmount As String

    ' Check to make sure it's a valid inventory slot.
    If GetPlayerInvItemNum(MyIndex, Inventory) < 1 Then Exit Sub
    If GetPlayerInvItemNum(MyIndex, Inventory) > MAX_ITEMS Then Exit Sub

    ' Check if the item is bound to the character.
    If Item(GetPlayerInvItemNum(MyIndex, Inventory)).Bound = 1 Then
        Call AddText("You cannot drop items bound to your character.", WHITE)
        Exit Sub
    End If

    ' Check if the item is a type of currency or is stackable.
    If Item(GetPlayerInvItemNum(MyIndex, Inventory)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, Inventory)).Stackable = 1 Then

        ' Prompt the user the amount of the specified item to drop.
        GoldAmount = InputBox("How much " & Trim$(Item(GetPlayerInvItemNum(MyIndex, Inventory)).Name) & "(" & GetPlayerInvItemValue(MyIndex, Inventory) & ") would you like to drop?", "Drop " & Trim$(Item(GetPlayerInvItemNum(MyIndex, Inventory)).Name), 0, frmMirage.Left, frmMirage.Top)

        ' Check if the item is numeric numbers only.
        If IsNumeric(GoldAmount) Then

            ' Check if the item amount is in a valid drop range.
            If CLng(GoldAmount) < 1 Or CLng(GoldAmount) > 100000000 Then
                Call AddText("Please enter a valid amount for that item!", BRIGHTRED)
                Exit Sub
            End If

            ' Drop the item on the ground.
            Call SendDropItem(Inventory, CLng(GoldAmount))
        End If
    Else
        ' Drop the item on the ground.
        Call SendDropItem(Inventory, 0)
    End If

    ' Update the visual inventory and equipment slots.
    Call UpdateVisInv
End Sub

' Sets the speed of a character based on speed
Sub SetSpeed(ByVal run As String, ByVal Speed As Long)
    If LCase$(run) = "walk" Then
        SS_WALK_SPEED = Speed
    ElseIf LCase$(run) = "run" Then
        SS_RUN_SPEED = Speed
    End If
' Ignore all other cases
End Sub

Sub MoveCharacter(ByVal MX As Integer, ByVal MY As Integer)
    If Player(MyIndex).input = 0 Then
        Exit Sub
    End If
    If GetPlayerY(MyIndex) = MAX_MAPY Then
        If MY = GetPlayerY(MyIndex) Then
            Call SendChangeDir
        End If
    Else
        If MY > GetPlayerY(MyIndex) And Val(MY - GetPlayerY(MyIndex)) > Val(MX - GetPlayerX(MyIndex)) Then
            Call SetPlayerDir(MyIndex, 1)
            If CanMove = True Then
                Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)
                DirDown = True
                Call SendPlayerMoveMouse
                Exit Sub
            End If
        End If
    End If

    If GetPlayerY(MyIndex) = 0 Then
        If MY = GetPlayerY(MyIndex) Then
            Call SendChangeDir
        End If
    Else
        If MY < GetPlayerY(MyIndex) And Val(MY - GetPlayerY(MyIndex)) < Val(MX - GetPlayerX(MyIndex)) Then
            Call SetPlayerDir(MyIndex, 0)
            If CanMove = True Then
                Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) - 1)
                DirUp = True
                Call SendPlayerMoveMouse
                Exit Sub
            End If
        End If
    End If

    If GetPlayerX(MyIndex) + 1 = MAX_MAPX Then
        If MX = GetPlayerX(MyIndex) Then
            Call SendChangeDir
        End If
    Else
        If MX > GetPlayerX(MyIndex) And Val(MY - GetPlayerY(MyIndex)) < Val(MX - GetPlayerX(MyIndex)) Then
            Call SetPlayerDir(MyIndex, 3)
            If CanMove = True Then
                Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) + 1)
                DirRight = True
                Call SendPlayerMoveMouse
                Exit Sub
            End If
        End If

    End If

    If GetPlayerX(MyIndex) = 0 Then
        If MX = GetPlayerX(MyIndex) Then
            Call SendChangeDir
        End If
    Else
        If MX < GetPlayerX(MyIndex) And Val(MY - GetPlayerY(MyIndex)) > Val(MX - GetPlayerX(MyIndex)) Then
            Call SetPlayerDir(MyIndex, 2)
            If CanMove = True Then
                Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)
                DirLeft = True
                Call SendPlayerMoveMouse
                Exit Sub
            End If
        End If
    End If
End Sub

Public Sub AlwaysOnTop(FormName As Form, bOnTop As Boolean)
    If Not bOnTop Then
        Call SetWindowPos(FormName.hWnd, HWND_TOPMOST, 0, 0, 0, 0, 3)
    Else
        Call SetWindowPos(FormName.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, 3)
    End If
End Sub

Sub GoShop(ByVal Shop As Integer)
    ' Close any other shop windows
    frmNewShop.Hide

    ' Initialize the shop
    frmNewShop.loadShop Shop
    snumber = Shop

    ' Hide panel
    frmNewShop.picItemInfo.Visible = False

    ' Show shop
    frmNewShop.Show vbModeless, frmMirage


    On Error Resume Next
    
    ' Set focus
    frmNewShop.SetFocus
    
End Sub

Sub IncrementGameClock()
    Dim CurTime As String

    Seconds = Seconds + Gamespeed

    If Seconds > 59 Then
        Minutes = Minutes + 1
        Seconds = Seconds - 60
    End If

    If Minutes > 59 Then
        Hours = Hours + 1
        Minutes = 0
    End If

    If Hours > 24 Then
        Hours = 1
    End If

    If Hours > 12 Then
        CurTime = CStr(Hours - 12)
    Else
        CurTime = Hours
    End If

    If Minutes < 10 Then
        CurTime = CurTime & ":0" & Minutes
    Else
        CurTime = CurTime & ":" & Minutes
    End If

    If Seconds < 10 Then
        CurTime = CurTime & ":0" & Seconds
    Else
        CurTime = CurTime & ":" & Seconds
    End If

    If Hours > 12 Then
        CurTime = CurTime & " PM"
    Else
        CurTime = CurTime & " AM"
    End If

    frmMirage.lblGameClock.Caption = CurTime
End Sub

' Returns true if the tile is a roof tile and the player is under that section of roof
Function IsTileRoof(ByVal X As Integer, ByVal y As Integer) As Boolean
    Dim IsRoof As Boolean
    
    If Map(GetPlayerMap(MyIndex)).Tile(X, y).Type = TILE_TYPE_ROOF Or Map(GetPlayerMap(MyIndex)).Tile(X, y).Type = TILE_TYPE_ROOFBLOCK Then 'If the tile is a roof or a roofblock
        If Map(GetPlayerMap(MyIndex)).Tile(X, y).String1 = Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).String1 Then 'If the roof ID is the same
            IsTileRoof = True
            Exit Function
        End If
    End If

    IsTileRoof = False
End Function

Public Sub GUI_PictureLoad(ByRef FormName As Form, ByVal FilePath As String)
    If FileExists(FilePath & ".gif") Then FormName.Picture = LoadPicture(App.Path & "\" & FilePath & ".gif")
    If FileExists(FilePath & ".jpg") Then FormName.Picture = LoadPicture(App.Path & "\" & FilePath & ".jpg")
    If FileExists(FilePath & ".png") Then FormName.Picture = LoadPicture(App.Path & "\" & FilePath & ".png")
    If FileExists(FilePath & ".bmp") Then FormName.Picture = LoadPicture(App.Path & "\" & FilePath & ".bmp")
End Sub

Function ItemIsEquipped(ByVal Index As Long, ByVal ItemNum As Long) As Boolean
    If GetPlayerWeaponSlot(Index) = ItemNum Then
        ItemIsEquipped = True
        Exit Function
    End If

    If GetPlayerArmorSlot(Index) = ItemNum Then
        ItemIsEquipped = True
        Exit Function
    End If

    If GetPlayerShieldSlot(Index) = ItemNum Then
        ItemIsEquipped = True
        Exit Function
    End If

    If GetPlayerHelmetSlot(Index) = ItemNum Then
        ItemIsEquipped = True
        Exit Function
    End If

    If GetPlayerLegsSlot(Index) = ItemNum Then
        ItemIsEquipped = True
        Exit Function
    End If

    If GetPlayerRingSlot(Index) = ItemNum Then
        ItemIsEquipped = True
        Exit Function
    End If

    If GetPlayerNecklaceSlot(Index) = ItemNum Then
        ItemIsEquipped = True
        Exit Function
    End If
End Function
