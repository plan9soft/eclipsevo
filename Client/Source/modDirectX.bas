Attribute VB_Name = "modDirectX"
Option Explicit

Public DX As DirectX7
Public DD As DirectDraw7

Public DD_Clip As DirectDrawClipper

Public DD_PrimarySurf As DirectDrawSurface7
Public DDSD_Primary As DDSURFACEDESC2

Public DD_SpriteSurf As DirectDrawSurface7
Public DDSD_Sprite As DDSURFACEDESC2

Public DD_ItemSurf As DirectDrawSurface7
Public DDSD_Item As DDSURFACEDESC2

Public DD_EmoticonSurf As DirectDrawSurface7
Public DDSD_Emoticon As DDSURFACEDESC2

Public DD_BackBuffer As DirectDrawSurface7
Public DDSD_BackBuffer As DDSURFACEDESC2

Public DD_BigSpriteSurf As DirectDrawSurface7
Public DDSD_BigSprite As DDSURFACEDESC2

Public DD_SpellAnim As DirectDrawSurface7
Public DDSD_SpellAnim As DDSURFACEDESC2

Public DD_BigSpellAnim As DirectDrawSurface7
Public DDSD_BigSpellAnim As DDSURFACEDESC2

Public DD_TileSurf(0 To ExtraSheets) As DirectDrawSurface7
Public DDSD_Tile(0 To ExtraSheets) As DDSURFACEDESC2
Public TileFile(0 To ExtraSheets) As Byte

Public DDSD_ArrowAnim As DDSURFACEDESC2
Public DD_ArrowAnim As DirectDrawSurface7

Public DD_PlayerHead As DirectDrawSurface7
Public DDSD_PlayerHead As DDSURFACEDESC2

Public DD_PlayerBody As DirectDrawSurface7
Public DDSD_PlayerBody As DDSURFACEDESC2

Public DD_PlayerLegs As DirectDrawSurface7
Public DDSD_PlayerLegs As DDSURFACEDESC2

Public Sub InitDirectX()
    On Error GoTo DXErr

    ' Initialize the DirextX 7 object.
    Set DX = New DirectX7

    ' Initialize the DirectDraw object.
    Set DD = DX.DirectDrawCreate(vbNullString)

    ' Inform DirectX that we're using windowed mode.
    Call DD.SetCooperativeLevel(frmMirage.hWnd, DDSCL_NORMAL)

    ' Prepare the primary surface description.
    DDSD_Primary.lFlags = DDSD_CAPS
    DDSD_Primary.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_SYSTEMMEMORY
    
    ' Create the primary surface.
    Set DD_PrimarySurf = DD.CreateSurface(DDSD_Primary)

    ' Create the clipper.
    Set DD_Clip = DD.CreateClipper(0)

    ' Associate the game screen with the clipper.
    DD_Clip.SetHWnd frmMirage.picScreen.hWnd

    ' Clip the blits to the game screen.
    DD_PrimarySurf.SetClipper DD_Clip

    ' Create all of the surfaces.
    Call InitSurfaces

    Exit Sub

DXErr:
    ' Failed to load the DirectX 7 object.
    Call MsgBox("Failed to initialize the  DirectX 7 object!" & vbNewLine & "Err: " & Err.Number & ", Desc: " & Err.Description, vbCritical)
    Call GameDestroy
End Sub

Public Sub BackBuffer_Create(Optional ByVal UseVideoMemory As Boolean = False)
    Dim DDSCAPS_MEMORYMODE As Long

    ' Set the surface to nothing. This is useful if we need to reload the surfaces.
    If Not DD_BackBuffer Is Nothing Then
        Set DD_BackBuffer = Nothing
    End If

    ' Should we use system memory or the video card?
    ' Please note: Eclipse doesn't run well with video memory yet!
    If UseVideoMemory Then
        DDSCAPS_MEMORYMODE = DDSCAPS_VIDEOMEMORY
    Else
        DDSCAPS_MEMORYMODE = DDSCAPS_SYSTEMMEMORY
    End If

    ' Create the back buffer surface description.
    DDSD_BackBuffer.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_BackBuffer.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_MEMORYMODE
    DDSD_BackBuffer.lWidth = (MAX_MAPX + 1) * PIC_X
    DDSD_BackBuffer.lHeight = (MAX_MAPY + 1) * PIC_Y

    ' Create the back buffer.
    Set DD_BackBuffer = DD.CreateSurface(DDSD_BackBuffer)
End Sub

Public Sub InitSurfaces()
    Dim TileID As Long

    ' Initiate the sprite surface description and create the surface.
    Call InitDDSurf("Sprites.bmp", DD_SpriteSurf, DDSD_Sprite)

    ' Initiate the tile surface description and create the surface.
    For TileID = 0 To ExtraSheets
        If FileExistsNew(App.Path & "\GFX\Tiles" & TileID & ".bmp") Then
            Call InitDDSurf("Tiles" & TileID & ".bmp", DD_TileSurf(TileID), DDSD_Tile(TileID))
            TileFile(TileID) = 1
        Else
            TileFile(TileID) = 0
        End If
    Next TileID

    ' Initiate the item surface description and create the surface.
    Call InitDDSurf("Items.bmp", DD_ItemSurf, DDSD_Item)

    ' Initiate the big sprites surface description and create the surface.
    Call InitDDSurf("BigSprites.bmp", DD_BigSpriteSurf, DDSD_BigSprite)

    ' Initiate the emoticon surface description and create the surface.
    Call InitDDSurf("Emoticons.bmp", DD_EmoticonSurf, DDSD_Emoticon)

    ' Initiate the spell surface description and create the surface.
    Call InitDDSurf("Spells.bmp", DD_SpellAnim, DDSD_SpellAnim)

    ' Initiate the big spells surface description and create the surface.
    Call InitDDSurf("BigSpells.bmp", DD_BigSpellAnim, DDSD_BigSpellAnim)

    ' Initiate the arrow surface description and create the surface.
    Call InitDDSurf("Arrows.bmp", DD_ArrowAnim, DDSD_ArrowAnim)

    ' Initiate the custom head surface description and create the surface.
    Call InitDDSurf("Heads.bmp", DD_PlayerHead, DDSD_PlayerHead)

    ' Initiate the custom body surface description and create the surface.
    Call InitDDSurf("Bodys.bmp", DD_PlayerBody, DDSD_PlayerBody)

    ' Initiate the custom leg surface description and create the surface.
    Call InitDDSurf("Legs.bmp", DD_PlayerLegs, DDSD_PlayerLegs)
End Sub

Public Sub InitDDSurf(ByVal FileName As String, ByRef DDSurf As DirectDrawSurface7, ByRef DDSurfDesc As DDSURFACEDESC2, Optional ByVal UseVideoMemory As Boolean = False)
    Dim DDSCAPS_MEMORYMODE As Long
    Dim TempFileName As String

    ' Set the path to the file.
    TempFileName = App.Path & "\GFX\" & FileName

    ' Check if the file exists before loading it.
    If Not FileExistsNew(TempFileName) Then
        Call MsgBox("The following file is missing: " & FileName & ".")
        Call GameDestroy
    End If

    ' Set the surface to nothing. This is useful if we need to reload the surfaces.
    If Not DDSurf Is Nothing Then
        Set DDSurf = Nothing
    End If

    ' Should we use system memory or the video card?
    ' Please note: Eclipse doesn't run well with video memory yet!
    If UseVideoMemory Then
        DDSCAPS_MEMORYMODE = DDSCAPS_VIDEOMEMORY
    Else
        DDSCAPS_MEMORYMODE = DDSCAPS_SYSTEMMEMORY
    End If

    ' Set the sufaces description.
    DDSurfDesc.lFlags = DDSD_CAPS
    DDSurfDesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_MEMORYMODE

    ' Create the surface from the description.
    Set DDSurf = DD.CreateSurfaceFromFile(TempFileName, DDSurfDesc)

    Call SetMaskColor(DDSurf)
End Sub

Public Sub DestroyDirectX()
    Dim TileID As Long

    ' Destroy DirectMusic7 and DirectSound7.
    Call DestroyDirectMusic
    Call DestroyDirectSound

    ' Remove all of the tile set objects from memory.
    For TileID = 0 To ExtraSheets
        If TileFile(TileID) = 1 Then
            Set DD_TileSurf(TileID) = Nothing
        End If
    Next TileID

    ' Remove the in-game objects from memory.
    Set DD_PrimarySurf = Nothing
    Set DD_SpriteSurf = Nothing
    Set DD_ItemSurf = Nothing
    Set DD_BigSpriteSurf = Nothing
    Set DD_EmoticonSurf = Nothing
    Set DD_SpellAnim = Nothing
    Set DD_BigSpellAnim = Nothing
    Set DD_ArrowAnim = Nothing

    ' Remove the custom player objects from memory.
    Set DD_PlayerHead = Nothing
    Set DD_PlayerBody = Nothing
    Set DD_PlayerLegs = Nothing

    ' Clear the back buffer.
    Set DD_BackBuffer = Nothing

    ' Clear the clipper.
    Set DD_Clip = Nothing

    ' Destroy the DirectX 7 objects.
    Set DD = Nothing
    Set DX = Nothing
End Sub

Public Function NeedToRestoreSurfaces() As Boolean
    Dim TestCoopRes As Long

    ' Checks if we're in the same display mode.
    TestCoopRes = DD.TestCooperativeLevel

    ' Inform the DirectX we need to restart.
    If TestCoopRes <> DD_OK Then
        NeedToRestoreSurfaces = True
    End If
End Function

Public Sub SetMaskColor(ByRef DDSurf As DirectDrawSurface7)
    Dim ColorKey As DDCOLORKEY

    ' Define the color key.
    ColorKey.low = 0
    ColorKey.high = 0

    ' Set the color key.
    Call DDSurf.SetColorKey(DDCKEY_SRCBLT, ColorKey)
End Sub

Public Function CreateRECT(ByVal Top As Long, ByVal Bottom As Long, ByVal Left As Long, ByVal Right As Long) As RECT
    CreateRECT.Top = Top
    CreateRECT.Bottom = CreateRECT.Top + Bottom
    CreateRECT.Left = Left
    CreateRECT.Right = CreateRECT.Left + Right
End Function

Sub DisplayFx(ByRef surfDisplay As DirectDrawSurface7, intX As Long, intY As Long, intWidth As Long, intHeight As Long, lngROP As Long, blnFxCap As Boolean, Tile As Long)
    Dim lngSrcDC As Long
    Dim lngDestDC As Long

    lngDestDC = DD_BackBuffer.GetDC
    lngSrcDC = surfDisplay.GetDC
    BitBlt lngDestDC, intX, intY, intWidth, intHeight, lngSrcDC, (Tile - Int(Tile / TilesInSheets) * TilesInSheets) * PIC_X, Int(Tile / TilesInSheets) * PIC_Y, lngROP
    surfDisplay.ReleaseDC lngSrcDC
    DD_BackBuffer.ReleaseDC lngDestDC
End Sub

Public Function GetScreenLeft(ByVal Index As Long) As Long
    GetScreenLeft = GetPlayerX(Index) - 11
End Function

Public Function GetScreenTop(ByVal Index As Long) As Long
    GetScreenTop = GetPlayerY(Index) - 8
End Function

Public Function GetScreenRight(ByVal Index As Long) As Long
    GetScreenRight = GetPlayerX(Index) + 10
End Function

Public Function GetScreenBottom(ByVal Index As Long) As Long
    GetScreenBottom = GetPlayerY(Index) + 8
End Function

Sub Night()
    Dim X As Long, y As Long
    
    ' Check if the night effect is enabled or disabled.
    If frmMirage.chkNightEffect.Value = vbUnchecked Then Exit Sub

    If TileFile(10) = 0 Then
        Exit Sub
    End If

    For y = ScreenY To ScreenY2
        For X = ScreenX To ScreenX2
            If Map(GetPlayerMap(MyIndex)).Tile(X, y).light <= 0 Then
                DisplayFx DD_TileSurf(10), (X - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, 32, 32, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT, 31
            Else
                DisplayFx DD_TileSurf(10), (X - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, 32, 32, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT, CLng(Map(GetPlayerMap(MyIndex)).Tile(X, y).light)
            End If
        Next X
    Next y
End Sub

Sub BltTile2(ByVal X As Long, ByVal y As Long, ByVal Tile As Long)
    Dim rec As DXVBLib.RECT

    If TileFile(10) = 0 Then
        Exit Sub
    End If

    rec.Top = Int(Tile / TilesInSheets) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (Tile - Int(Tile / TilesInSheets) * TilesInSheets) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) + sx - NewXOffset, y - (NewPlayerY * PIC_Y) + sx - NewYOffset, DD_TileSurf(10), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
'    DisplayFx DD_TileSurf(10), (x - NewPlayerX * PIC_X) + sx - NewXOffset, y - (NewPlayerY * PIC_Y) + sx - NewYOffset, 32, 16, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT, Tile
End Sub

Sub BltPlayerText(ByVal Index As Long)
    Dim TextX As Long
    Dim TextY As Long
    Dim intLoop As Integer
    Dim intLoop2 As Integer

    Dim bytLineCount As Byte
    Dim bytLineLength As Byte
    Dim strLine(0 To MAX_LINES - 1) As String
    Dim strWords() As String
    strWords() = Split(Bubble(Index).Text, " ")

    If Len(Bubble(Index).Text) < MAX_LINE_LENGTH Then
        DISPLAY_BUBBLE_WIDTH = 2 + ((Len(Bubble(Index).Text) * 9) \ PIC_X)

        If DISPLAY_BUBBLE_WIDTH > MAX_BUBBLE_WIDTH Then
            DISPLAY_BUBBLE_WIDTH = MAX_BUBBLE_WIDTH
        End If
    Else
        DISPLAY_BUBBLE_WIDTH = 6
    End If

    TextX = GetPlayerX(Index) * PIC_X + Player(Index).xOffset + Int(PIC_X) - ((DISPLAY_BUBBLE_WIDTH * 32) / 2) - 6
    TextY = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset - Int(PIC_Y) + 75

    Call DD_BackBuffer.ReleaseDC(TexthDC)

    ' Draw the fancy box with tiles.
    Call BltTile2(TextX - 10, TextY - 10, 1)
    Call BltTile2(TextX + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X - 10, TextY - 10, 2)

    For intLoop = 1 To (DISPLAY_BUBBLE_WIDTH - 2)

        Call BltTile2(TextX - 10 + (intLoop * PIC_X), TextY - 10, 16)
    Next intLoop

    TexthDC = DD_BackBuffer.GetDC

    ' Loop through all the words.
    For intLoop = 0 To UBound(strWords)
        ' Increment the line length.
        bytLineLength = bytLineLength + Len(strWords(intLoop)) + 1

        ' If we have room on the current line.
        If bytLineLength < MAX_LINE_LENGTH Then
            ' Add the text to the current line.
            strLine(bytLineCount) = strLine(bytLineCount) & strWords(intLoop) & " "
        Else
            bytLineCount = bytLineCount + 1

            If bytLineCount = MAX_LINES Then
                bytLineCount = bytLineCount - 1
                Exit For
            End If

            strLine(bytLineCount) = Trim$(strWords(intLoop)) & " "
            bytLineLength = 0
        End If
    Next intLoop

    Call DD_BackBuffer.ReleaseDC(TexthDC)

    If bytLineCount > 0 Then
        For intLoop = 6 To (bytLineCount - 2) * PIC_Y + 6
            Call BltTile2(TextX - 10, TextY - 10 + intLoop, 19)
            Call BltTile2(TextX - 10 + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X, TextY - 10 + intLoop, 17)

            For intLoop2 = 1 To DISPLAY_BUBBLE_WIDTH - 2
                Call BltTile2(TextX - 10 + (intLoop2 * PIC_X), TextY + intLoop - 10, 5)
            Next intLoop2
        Next intLoop
    End If

    Call BltTile2(TextX - 10, TextY + (bytLineCount * 4) - 4, 3)
    Call BltTile2(TextX + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X - 10, TextY + (bytLineCount * 4) - 4, 4)

    For intLoop = 1 To (DISPLAY_BUBBLE_WIDTH - 2)
        Call BltTile2(TextX - 10 + (intLoop * PIC_X), TextY + (bytLineCount * 16) - 4, 15)
    Next intLoop

    TexthDC = DD_BackBuffer.GetDC

    For intLoop = 0 To (MAX_LINES - 1)
        If strLine(intLoop) <> vbNullString Then
            Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) + sx - NewXOffset + (((DISPLAY_BUBBLE_WIDTH) * PIC_X) / 2) - ((Len(strLine(intLoop)) * 8) \ 2) - 4, TextY - (NewPlayerY * PIC_Y) + sx - NewYOffset, strLine(intLoop), QBColor(WHITE))
            TextY = TextY + 16
        End If
    Next intLoop
End Sub

Sub Bltscriptbubble(ByVal Index As Long, ByVal MapNum As Long, ByVal X As Long, ByVal y As Long, ByVal Colour As Long)
    Dim TextX As Long
    Dim TextY As Long
    Dim intLoop As Integer
    Dim intLoop2 As Integer

    Dim bytLineCount As Byte
    Dim bytLineLength As Byte
    Dim strLine(0 To MAX_LINES - 1) As String
    Dim strWords() As String

    strWords() = Split(ScriptBubble(Index).Text, " ")

    If Len(ScriptBubble(Index).Text) < MAX_LINE_LENGTH Then
        DISPLAY_BUBBLE_WIDTH = 2 + ((Len(ScriptBubble(Index).Text) * 9) \ PIC_X)

        If DISPLAY_BUBBLE_WIDTH > MAX_BUBBLE_WIDTH Then
            DISPLAY_BUBBLE_WIDTH = MAX_BUBBLE_WIDTH
        End If
    Else
        DISPLAY_BUBBLE_WIDTH = 6
    End If

    ' TextX = X * PIC_X + Int(PIC_X) - ((DISPLAY_BUBBLE_WIDTH * 32) / 2) - 6
    TextX = X * PIC_X - 22
    TextY = y * PIC_Y - 22

    Call DD_BackBuffer.ReleaseDC(TexthDC)

    ' Draw the fancy box with tiles.
    Call BltTile2(TextX - 10, TextY - 10, 1)
    Call BltTile2(TextX + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X - 10, TextY - 10, 2)

    For intLoop = 1 To (DISPLAY_BUBBLE_WIDTH - 2)
        Call BltTile2(TextX - 10 + (intLoop * PIC_X), TextY - 10, 16)
    Next intLoop

    TexthDC = DD_BackBuffer.GetDC

    ' Loop through all the words.
    For intLoop = 0 To UBound(strWords)
        ' Increment the line length.
        bytLineLength = bytLineLength + Len(strWords(intLoop)) + 1

        ' If we have room on the current line.
        If bytLineLength < MAX_LINE_LENGTH Then
            ' Add the text to the current line.
            strLine(bytLineCount) = strLine(bytLineCount) & strWords(intLoop) & " "
        Else
            bytLineCount = bytLineCount + 1

            If bytLineCount = MAX_LINES Then
                bytLineCount = bytLineCount - 1
                Exit For
            End If

            strLine(bytLineCount) = Trim$(strWords(intLoop)) & " "
            bytLineLength = 0
        End If
    Next intLoop

    Call DD_BackBuffer.ReleaseDC(TexthDC)

    If bytLineCount > 0 Then
        For intLoop = 6 To (bytLineCount - 2) * PIC_Y + 6
            Call BltTile2(TextX - 10, TextY - 10 + intLoop, 19)
            Call BltTile2(TextX - 10 + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X, TextY - 10 + intLoop, 17)

            For intLoop2 = 1 To DISPLAY_BUBBLE_WIDTH - 2
                Call BltTile2(TextX - 10 + (intLoop2 * PIC_X), TextY + intLoop - 10, 5)
            Next intLoop2
        Next intLoop
    End If

    Call BltTile2(TextX - 10, TextY + (bytLineCount * 16) - 4, 3)
    Call BltTile2(TextX + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X - 10, TextY + (bytLineCount * 16) - 4, 4)

    For intLoop = 1 To (DISPLAY_BUBBLE_WIDTH - 2)
        Call BltTile2(TextX - 10 + (intLoop * PIC_X), TextY + (bytLineCount * 16) - 4, 15)
    Next intLoop

    TexthDC = DD_BackBuffer.GetDC

    For intLoop = 0 To (MAX_LINES - 1)
        If strLine(intLoop) <> vbNullString Then
            Call DrawText(TexthDC, TextX + (((DISPLAY_BUBBLE_WIDTH) * PIC_X) / 2) - ((Len(strLine(intLoop)) * 8) \ 2) - 7, TextY, strLine(intLoop), QBColor(Colour))
            TextY = TextY + 16
        End If
    Next intLoop
End Sub

Sub BltPlayerBars(ByVal Index As Long)
    Dim X As Long, y As Long

    X = (GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset) - (NewPlayerX * PIC_X) - NewXOffset
    y = (GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset) - (NewPlayerY * PIC_Y) - NewYOffset


    If Player(Index).HP = 0 Then
        Exit Sub
    End If
    If SpriteSize = 1 Then
        ' draws the back bars
        Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
        Call DD_BackBuffer.DrawBox(X, y - 30, X + 32, y - 34)

        ' draws HP
        Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
        Call DD_BackBuffer.DrawBox(X, y - 30, X + ((Player(Index).HP / 100) / (Player(Index).MaxHP / 100) * 32), y - 34)
    Else
        If SpriteSize = 2 Then
            ' draws the back bars
            Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
            Call DD_BackBuffer.DrawBox(X, y - 30 - PIC_Y, X + 32, y - 34 - PIC_Y)

            ' draws HP
            Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
            Call DD_BackBuffer.DrawBox(X, y - 30 - PIC_Y, X + ((Player(Index).HP / 100) / (Player(Index).MaxHP / 100) * 32), y - 34 - PIC_Y)
        Else
            ' draws the back bars
            Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
            Call DD_BackBuffer.DrawBox(X, y + 2, X + 32, y - 2)

            ' draws HP
            Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
            Call DD_BackBuffer.DrawBox(X, y + 2, X + ((Player(Index).HP / 100) / (Player(Index).MaxHP / 100) * 32), y - 2)
        End If
    End If
End Sub

Sub BltNpcBars(ByVal Index As Long)
    Dim X As Long, y As Long

    On Error GoTo BltNpcBars_Error

    If MapNpc(Index).HP = 0 Then
        Exit Sub
    End If
    If MapNpc(Index).num < 1 Then
        Exit Sub
    End If

    If Npc(MapNpc(Index).num).Big = 1 Then
        X = (MapNpc(Index).X * PIC_X + sx - 9 + MapNpc(Index).xOffset) - (NewPlayerX * PIC_X) - NewXOffset
        y = (MapNpc(Index).y * PIC_Y + sx + MapNpc(Index).yOffset) - (NewPlayerY * PIC_Y) - NewYOffset

        Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
        Call DD_BackBuffer.DrawBox(X, y + 32, X + 50, y + 36)
        Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
        If MapNpc(Index).MaxHP < 1 Then
            Call DD_BackBuffer.DrawBox(X, y + 32, X + ((MapNpc(Index).HP / 100) / ((MapNpc(Index).MaxHP + 1) / 100) * 50), y + 36)
        Else
            Call DD_BackBuffer.DrawBox(X, y + 32, X + ((MapNpc(Index).HP / 100) / (MapNpc(Index).MaxHP / 100) * 50), y + 36)
        End If
    Else
        X = (MapNpc(Index).X * PIC_X + sx + MapNpc(Index).xOffset) - (NewPlayerX * PIC_X) - NewXOffset
        y = (MapNpc(Index).y * PIC_Y + sx + MapNpc(Index).yOffset) - (NewPlayerY * PIC_Y) - NewYOffset

        Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
        Call DD_BackBuffer.DrawBox(X, y + 32, X + 32, y + 36)
        Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))

        If MapNpc(Index).MaxHP < 1 Then
            Call DD_BackBuffer.DrawBox(X, y + 32, X + ((MapNpc(Index).HP / 100) / ((MapNpc(Index).MaxHP + 1) / 100) * 32), y + 36)
        Else
            Call DD_BackBuffer.DrawBox(X, y + 32, X + ((MapNpc(Index).HP / 100) / (MapNpc(Index).MaxHP / 100) * 32), y + 36)
        End If

    End If


    On Error GoTo 0
    Exit Sub

BltNpcBars_Error:

    If Err.Number = DDERR_CANTCREATEDC Then

    End If

End Sub

Sub BltWeather()
    Dim rec As DXVBLib.RECT
    Dim i As Long

    Call DD_BackBuffer.SetForeColor(RGB(0, 0, 200))

    If GameWeather = WEATHER_RAINING Or GameWeather = WEATHER_THUNDER Then
        For i = 1 To MAX_RAINDROPS
            If DropRain(i).Randomized = False Then
                If frmMirage.tmrRainDrop.Enabled = False Then
                    BLT_RAIN_DROPS = 1
                    frmMirage.tmrRainDrop.Enabled = True
                    If frmMirage.tmrRainDrop.Tag = vbNullString Then
                        frmMirage.tmrRainDrop.Interval = 100
                        frmMirage.tmrRainDrop.Tag = "123"
                    End If
                End If
            End If
        Next i
    ElseIf GameWeather = WEATHER_SNOWING Then
        For i = 1 To MAX_RAINDROPS
            If DropSnow(i).Randomized = False Then
                If frmMirage.tmrSnowDrop.Enabled = False Then
                    BLT_SNOW_DROPS = 1
                    frmMirage.tmrSnowDrop.Enabled = True
                    If frmMirage.tmrSnowDrop.Tag = vbNullString Then
                        frmMirage.tmrSnowDrop.Interval = 200
                        frmMirage.tmrSnowDrop.Tag = "123"
                    End If
                End If
            End If
        Next i
    Else
        If BLT_RAIN_DROPS > 0 And BLT_RAIN_DROPS <= RainIntensity Then
            Call ClearRainDrop(BLT_RAIN_DROPS)
        End If
        frmMirage.tmrRainDrop.Tag = vbNullString
    End If

    For i = 1 To MAX_RAINDROPS
        If Not ((DropRain(i).X = 0) Or (DropRain(i).y = 0)) Then
            rec.Top = 0
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = 6 * PIC_X
            rec.Right = rec.Left + PIC_X
            DropRain(i).X = DropRain(i).X + DropRain(i).Speed
            DropRain(i).y = DropRain(i).y + DropRain(i).Speed
            Call DD_BackBuffer.BltFast(DropRain(i).X + DropRain(i).Speed, DropRain(i).y + DropRain(i).Speed, DD_TileSurf(10), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            If (DropRain(i).X > (MAX_MAPX + 1) * PIC_X) Or (DropRain(i).y > (MAX_MAPY + 1) * PIC_Y) Then
                DropRain(i).Randomized = False
            End If
        End If
    Next i
    If TileFile(10) = 1 Then
        rec.Top = Int(14 / TilesInSheets) * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = (14 - Int(14 / TilesInSheets) * TilesInSheets) * PIC_X
        rec.Right = rec.Left + PIC_X
        For i = 1 To MAX_RAINDROPS
            If Not ((DropSnow(i).X = 0) Or (DropSnow(i).y = 0)) Then
                DropSnow(i).X = DropSnow(i).X + DropSnow(i).Speed
                DropSnow(i).y = DropSnow(i).y + DropSnow(i).Speed
                Call DD_BackBuffer.BltFast(DropSnow(i).X + DropSnow(i).Speed, DropSnow(i).y + DropSnow(i).Speed, DD_TileSurf(10), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                If (DropSnow(i).X > (MAX_MAPX + 1) * PIC_X) Or (DropSnow(i).y > (MAX_MAPY + 1) * PIC_Y) Then
                    DropSnow(i).Randomized = False
                End If
            End If
        Next i
    End If

    ' If it's thunder, make the screen randomly flash white
    If GameWeather = WEATHER_THUNDER Then
        If Int((100 - 1 + 1) * Rnd) + 1 = 8 Then
            DD_BackBuffer.SetFillColor RGB(255, 255, 255)

            Call DD_BackBuffer.DrawBox(0, 0, (MAX_MAPX + 1) * PIC_X, (MAX_MAPY + 1) * PIC_Y)
        End If
    End If
End Sub

Sub BltMapWeather()
    Dim rec As DXVBLib.RECT
    Dim i As Long

    Call DD_BackBuffer.SetForeColor(RGB(0, 0, 200))

    If Map(GetPlayerMap(MyIndex)).Weather = 1 Or Map(GetPlayerMap(MyIndex)).Weather = 3 Then
        For i = 1 To MAX_RAINDROPS
            If DropRain(i).Randomized = False Then
                If frmMirage.tmrRainDrop.Enabled = False Then
                    BLT_RAIN_DROPS = 1
                    frmMirage.tmrRainDrop.Enabled = True
                End If
            End If
        Next i
        For i = 1 To MAX_RAINDROPS
            If Not ((DropRain(i).X = 0) Or (DropRain(i).y = 0)) Then
                rec.Top = (14 - Int(14 / TilesInSheets)) * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                rec.Left = 6 * PIC_X
                rec.Right = rec.Left + PIC_X
                DropRain(i).X = DropRain(i).X + DropRain(i).Speed
                DropRain(i).y = DropRain(i).y + DropRain(i).Speed
                Call DD_BackBuffer.BltFast(DropRain(i).X + DropRain(i).Speed, DropRain(i).y + DropRain(i).Speed, DD_TileSurf(10), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                If (DropRain(i).X > (MAX_MAPX + 1) * PIC_X) Or (DropRain(i).y > (MAX_MAPY + 1) * PIC_Y) Then
                    DropRain(i).Randomized = False
                End If
            End If
        Next i

        If Map(GetPlayerMap(MyIndex)).Weather = 3 Then
            If Int((100 - 1 + 1) * Rnd) + 1 < 3 Then
                DD_BackBuffer.SetFillColor RGB(255, 255, 255)

                Call DD_BackBuffer.DrawBox(0, 0, (MAX_MAPX + 1) * PIC_X, (MAX_MAPY + 1) * PIC_Y)
            End If
        End If

    ElseIf Map(GetPlayerMap(MyIndex)).Weather = 2 Then
        For i = 1 To MAX_RAINDROPS
            If DropSnow(i).Randomized = False Then
                If frmMirage.tmrSnowDrop.Enabled = False Then
                    BLT_SNOW_DROPS = 1
                    frmMirage.tmrSnowDrop.Enabled = True
                End If
            End If
        Next i
        If TileFile(10) = 1 Then
            rec.Top = Int(14 / TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (14 - Int(14 / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X

            For i = 1 To MAX_RAINDROPS
                If Not ((DropSnow(i).X = 0) Or (DropSnow(i).y = 0)) Then
                    DropSnow(i).X = DropSnow(i).X + DropSnow(i).Speed
                    DropSnow(i).y = DropSnow(i).y + DropSnow(i).Speed
                    Call DD_BackBuffer.BltFast(DropSnow(i).X + DropSnow(i).Speed, DropSnow(i).y + DropSnow(i).Speed, DD_TileSurf(10), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    If (DropSnow(i).X > (MAX_MAPX + 1) * PIC_X) Or (DropSnow(i).y > (MAX_MAPY + 1) * PIC_Y) Then
                        DropSnow(i).Randomized = False
                    End If
                End If
            Next i
        End If
    Else
        If BLT_RAIN_DROPS > 0 And BLT_RAIN_DROPS <= frmMirage.tmrRainDrop.Interval Then
            Call ClearRainDrop(BLT_RAIN_DROPS)
        End If
        frmMirage.tmrRainDrop.Tag = vbNullString
    End If
End Sub

Sub RNDRainDrop(ByVal RDNumber As Long)
Start:
    DropRain(RDNumber).X = Int((((MAX_MAPX + 1) * PIC_X) * Rnd) + 1)
    DropRain(RDNumber).y = Int((((MAX_MAPY + 1) * PIC_Y) * Rnd) + 1)
    If (DropRain(RDNumber).y > (MAX_MAPY + 1) * PIC_Y / 4) And (DropRain(RDNumber).X > (MAX_MAPX + 1) * PIC_X / 4) Then
        GoTo Start
    End If
    DropRain(RDNumber).Speed = Int((10 * Rnd) + 6)
    DropRain(RDNumber).Randomized = True
End Sub

Sub ClearRainDrop(ByVal RDNumber As Long)
    On Error Resume Next
    DropRain(RDNumber).X = 0
    DropRain(RDNumber).y = 0
    DropRain(RDNumber).Speed = 0
    DropRain(RDNumber).Randomized = False
End Sub

Sub RNDSnowDrop(ByVal RDNumber As Long)
Start:
    DropSnow(RDNumber).X = Int((((MAX_MAPX + 1) * PIC_X) * Rnd) + 1)
    DropSnow(RDNumber).y = Int((((MAX_MAPY + 1) * PIC_Y) * Rnd) + 1)
    If (DropSnow(RDNumber).y > (MAX_MAPY + 1) * PIC_Y / 4) And (DropSnow(RDNumber).X > (MAX_MAPX + 1) * PIC_X / 4) Then
        GoTo Start
    End If
    DropSnow(RDNumber).Speed = Int((10 * Rnd) + 6)
    DropSnow(RDNumber).Randomized = True
End Sub

Sub ClearSnowDrop(ByVal RDNumber As Long)
    On Error Resume Next
    DropSnow(RDNumber).X = 0
    DropSnow(RDNumber).y = 0
    DropSnow(RDNumber).Speed = 0
    DropSnow(RDNumber).Randomized = False
End Sub

Sub BltSpell(ByVal Index As Long)
    Dim rec As DXVBLib.RECT
    Dim X As Long
    Dim y As Long
    Dim i As Long

    If Player(Index).SpellNum <= 0 Or Player(Index).SpellNum > MAX_SPELLS Then
        Exit Sub
    End If


    For i = 1 To MAX_SPELL_ANIM
        ' IF SPELL IS NOT BIG
        If Spell(Player(Index).SpellNum).Big = 0 Then
            If Player(Index).SpellAnim(i).CastedSpell = YES Then
                If Player(Index).SpellAnim(i).SpellDone < Spell(Player(Index).SpellNum).SpellDone Then

                    rec.Top = Spell(Player(Index).SpellNum).SpellAnim * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    rec.Left = Player(Index).SpellAnim(i).SpellVar * PIC_X
                    rec.Right = rec.Left + PIC_X

                    If Player(Index).SpellAnim(i).TargetType = 0 Then

                        ' SMALL: IF TARGET IS A PLAYER
                        If Player(Index).SpellAnim(i).Target > 0 Then

                            ' SMALL: IF TARGET IS SELF
                            If Player(Index).SpellAnim(i).Target = MyIndex Then
                                X = NewX + sx
                                y = NewY + sx
                                Call DD_BackBuffer.BltFast(X, y, DD_SpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

                            ' SMALL: IF TARGET IS ANOTHER PLAYER
                            Else
                                X = GetPlayerX(Player(Index).SpellAnim(i).Target) * PIC_X + sx + Player(Player(Index).SpellAnim(i).Target).xOffset
                                y = GetPlayerY(Player(Index).SpellAnim(i).Target) * PIC_Y + sx + Player(Player(Index).SpellAnim(i).Target).yOffset
                                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                            End If
                        End If

                    ' SMALL: IF TARGET IS AN NPC
                    Else
                        X = MapNpc(Player(Index).SpellAnim(i).Target).X * PIC_X + sx + MapNpc(Player(Index).SpellAnim(i).Target).xOffset
                        y = MapNpc(Player(Index).SpellAnim(i).Target).y * PIC_Y + sx + MapNpc(Player(Index).SpellAnim(i).Target).yOffset
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If


' SMALL: ADVANCE SPELL ONE CYCLE

                    If GetTickCount > Player(Index).SpellAnim(i).SpellTime + Spell(Player(Index).SpellNum).SpellTime Then
                        Player(Index).SpellAnim(i).SpellTime = GetTickCount
                        Player(Index).SpellAnim(i).SpellVar = Player(Index).SpellAnim(i).SpellVar + 1
                    End If

                    If Player(Index).SpellAnim(i).SpellVar > 12 Then
                        Player(Index).SpellAnim(i).SpellDone = Player(Index).SpellAnim(i).SpellDone + 1
                        Player(Index).SpellAnim(i).SpellVar = 0
                    End If

                Else
                    Player(Index).SpellAnim(i).CastedSpell = NO
                End If
            End If
        Else
            If Player(Index).SpellAnim(i).CastedSpell = YES Then
                If Player(Index).SpellAnim(i).SpellDone < Spell(Player(Index).SpellNum).SpellDone Then

                    rec.Top = Spell(Player(Index).SpellNum).SpellAnim * (PIC_Y * 3)
                    rec.Bottom = rec.Top + PIC_Y + 64
                    rec.Left = Player(Index).SpellAnim(i).SpellVar * PIC_X
                    rec.Right = rec.Left + PIC_X + 64

                    If Player(Index).SpellAnim(i).TargetType = 0 Then

                        ' BIG: IF TARGET IS A PLAYER
                        If Player(Index).SpellAnim(i).Target > 0 Then

                            ' BIG: IF TARGET IS SELF
                            If Player(Index).SpellAnim(i).Target = MyIndex Then
                                X = NewX + sx - 32
                                y = NewY + sx - 32

                                If y < 0 Then
                                    rec.Top = rec.Top + (y * -1)
                                    y = 0
                                End If

                                If X < 0 Then
                                    rec.Left = rec.Left + (X * -1)
                                    X = 0
                                End If

                                If (X + 64) > (MAX_MAPX * 32) Then
                                    rec.Right = rec.Left + 64
                                End If

                                If (y + 64) > (MAX_MAPY * 32) Then
                                    rec.Bottom = rec.Top + 64
                                End If

                                Call DD_BackBuffer.BltFast(X, y, DD_BigSpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

                            ' BIG: IF TARGET IS A DIFFERENT PLAYER
                            Else
                                X = GetPlayerX(Player(Index).SpellAnim(i).Target) * PIC_X + sx - 32 + Player(Player(Index).SpellAnim(i).Target).xOffset
                                y = GetPlayerY(Player(Index).SpellAnim(i).Target) * PIC_Y + sx - 32 + Player(Player(Index).SpellAnim(i).Target).yOffset

                                If y < 0 Then
                                    rec.Top = rec.Top + (y * -1)
                                    y = 0
                                End If

                                If X < 0 Then
                                    rec.Left = rec.Left + (X * -1)
                                    X = 0
                                End If

                                If (X + 64) > (MAX_MAPX * 32) Then
                                    rec.Right = rec.Left + 64
                                End If

                                If (y + 64) > (MAX_MAPY * 32) Then
                                    rec.Bottom = rec.Top + 64
                                End If

                                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_BigSpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                            End If
                        End If

                    ' BIG: IF TARGET IS AN NPC
                    Else
                        X = MapNpc(Player(Index).SpellAnim(i).Target).X * PIC_X + sx - 32 + MapNpc(Player(Index).SpellAnim(i).Target).xOffset
                        y = MapNpc(Player(Index).SpellAnim(i).Target).y * PIC_Y + sx - 32 + MapNpc(Player(Index).SpellAnim(i).Target).yOffset

                        If y < 0 Then
                            rec.Top = rec.Top + (y * -1)
                            y = 0
                        End If

                        If X < 0 Then
                            rec.Left = rec.Left + (X * -1)
                            X = 0
                        End If

                        If (X + 64) > (MAX_MAPX * 32) Then
                            rec.Right = rec.Left + 64
                        End If

                        If (y + 64) > (MAX_MAPY * 32) Then
                            rec.Bottom = rec.Top + 64
                        End If

                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_BigSpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' BIG: ADVANCE SPELL ONE CYCLE
                    If GetTickCount > Player(Index).SpellAnim(i).SpellTime + Spell(Player(Index).SpellNum).SpellTime Then
                        Player(Index).SpellAnim(i).SpellTime = GetTickCount
                        Player(Index).SpellAnim(i).SpellVar = Player(Index).SpellAnim(i).SpellVar + 3
                    End If

                    If Player(Index).SpellAnim(i).SpellVar > 36 Then
                        Player(Index).SpellAnim(i).SpellDone = Player(Index).SpellAnim(i).SpellDone + 1
                        Player(Index).SpellAnim(i).SpellVar = 0
                    End If

                Else
                    Player(Index).SpellAnim(i).CastedSpell = NO
                End If
            End If
        End If
    Next i
End Sub

' Scripted Spell
Sub BltScriptSpell(ByVal i As Long)
    Dim rec As RECT
    Dim X As Long, y As Long

    X = ScriptSpell(i).X
    y = ScriptSpell(i).y

    If Spell(ScriptSpell(i).SpellNum).Big = 0 Then
        If ScriptSpell(i).SpellDone < Spell(ScriptSpell(i).SpellNum).SpellDone Then
            rec.Top = Spell(ScriptSpell(i).SpellNum).SpellAnim * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = ScriptSpell(i).SpellVar * PIC_X
            rec.Right = rec.Left + PIC_X

            X = X * PIC_X + sx
            y = y * PIC_Y + sx

            If ScriptSpell(i).SpellVar > 10 Then
                ScriptSpell(i).SpellDone = ScriptSpell(i).SpellDone + 1
                ScriptSpell(i).SpellVar = 0
            End If

            If GetTickCount > ScriptSpell(i).SpellTime + Spell(ScriptSpell(i).SpellNum).SpellTime Then
                ScriptSpell(i).SpellTime = GetTickCount
                ScriptSpell(i).SpellVar = ScriptSpell(i).SpellVar + 1
            End If

            Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X), y - (NewPlayerY * PIC_Y), DD_SpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

        Else ' spell is done
            ScriptSpell(i).CastedSpell = NO
        End If
    Else
        If ScriptSpell(i).SpellDone < Spell(ScriptSpell(i).SpellNum).SpellDone Then
            rec.Top = Spell(ScriptSpell(i).SpellNum).SpellAnim * (PIC_Y * 3)
            rec.Bottom = rec.Top + PIC_Y + 64
            rec.Left = ScriptSpell(i).SpellVar * PIC_X
            rec.Right = rec.Left + PIC_X + 64

            X = X * PIC_X + sx - 32
            y = y * PIC_Y + sx - 32

            If y < 0 Then
                rec.Top = rec.Top + (y * -1)
                y = 0
            End If

            If X < 0 Then
                rec.Left = rec.Left + (X * -1)
                X = 0
            End If

            If (X + 64) > (MAX_MAPX * 32) Then
                rec.Right = rec.Left + 64
            End If

            If (y + 64) > (MAX_MAPY * 32) Then
                rec.Bottom = rec.Top + 64
            End If

            If ScriptSpell(i).SpellVar > 30 Then
                ScriptSpell(i).SpellDone = ScriptSpell(i).SpellDone + 1
                ScriptSpell(i).SpellVar = 0
            End If

            If GetTickCount > ScriptSpell(i).SpellTime + Spell(ScriptSpell(i).SpellNum).SpellTime Then
                ScriptSpell(i).SpellTime = GetTickCount
                ScriptSpell(i).SpellVar = ScriptSpell(i).SpellVar + 3
            End If

            Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X), y - (NewPlayerY * PIC_Y), DD_BigSpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        Else 'spell is done
            ScriptSpell(i).CastedSpell = NO
        End If
    End If
End Sub

Sub BltEmoticons(ByVal Index As Long)
    Dim rec As DXVBLib.RECT
    Dim x2 As Long
    Dim y2 As Long
    Dim ETime As Long

    ETime = 1300

    If Player(Index).EmoticonNum < 0 Then
        Exit Sub
    End If

    If Player(Index).EmoticonTime + ETime > GetTickCount Then
        If GetTickCount < Player(Index).EmoticonTime + Int((ETime / 12) * 1) Then
            Player(Index).EmoticonVar = 0
        ElseIf GetTickCount < Player(Index).EmoticonTime + Int((ETime / 12) * 2) Then
            Player(Index).EmoticonVar = 1
        ElseIf GetTickCount < Player(Index).EmoticonTime + Int((ETime / 12) * 3) Then
            Player(Index).EmoticonVar = 2
        ElseIf GetTickCount < Player(Index).EmoticonTime + Int((ETime / 12) * 4) Then
            Player(Index).EmoticonVar = 3
        ElseIf GetTickCount < Player(Index).EmoticonTime + Int((ETime / 12) * 5) Then
            Player(Index).EmoticonVar = 4
        ElseIf GetTickCount < Player(Index).EmoticonTime + Int((ETime / 12) * 6) Then
            Player(Index).EmoticonVar = 5
        ElseIf GetTickCount < Player(Index).EmoticonTime + Int((ETime / 12) * 7) Then
            Player(Index).EmoticonVar = 6
        ElseIf GetTickCount < Player(Index).EmoticonTime + Int((ETime / 12) * 8) Then
            Player(Index).EmoticonVar = 7
        ElseIf GetTickCount < Player(Index).EmoticonTime + Int((ETime / 12) * 9) Then
            Player(Index).EmoticonVar = 8
        ElseIf GetTickCount < Player(Index).EmoticonTime + Int((ETime / 12) * 10) Then
            Player(Index).EmoticonVar = 9
        ElseIf GetTickCount < Player(Index).EmoticonTime + Int((ETime / 12) * 11) Then
            Player(Index).EmoticonVar = 10
        ElseIf GetTickCount < Player(Index).EmoticonTime + Int((ETime / 12) * 12) Then
            Player(Index).EmoticonVar = 11
        End If

        rec.Top = Player(Index).EmoticonNum * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = Player(Index).EmoticonVar * PIC_X
        rec.Right = rec.Left + PIC_X

        If Index = MyIndex Then
            x2 = NewX + sx + 16
            y2 = NewY + sx - 32

            If y2 < 0 Then
                Exit Sub
            End If

            Call DD_BackBuffer.BltFast(x2, y2, DD_EmoticonSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        Else
            x2 = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset + 16
            y2 = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset - 32

            If y2 < 0 Then
                Exit Sub
            End If

            Call DD_BackBuffer.BltFast(x2 - (NewPlayerX * PIC_X) - NewXOffset, y2 - (NewPlayerY * PIC_Y) - NewYOffset, DD_EmoticonSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
End Sub

Sub Bltgrapple(ByVal Index As Long)
    Dim rec As DXVBLib.RECT
    Dim z As Integer
    Dim BX As Long
    Dim BY As Long

    If Player(Index).HookShotX > 0 Or Player(Index).HookShotY <> 0 Then

        Select Case Player(Index).HookShotDir
            Case 0
                z = 1
            Case 1
                z = 0
            Case 2
                z = 3
            Case 3
                z = 2
        End Select

        rec.Top = Player(Index).HookShotAnim * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = z * PIC_X
        rec.Right = rec.Left + PIC_X

        If GetTickCount > Player(Index).HookShotTime + 50 Then
            If Player(Index).HookShotSucces = 1 Then
                If Index = MyIndex Then
                Call SendData(POut.EndShot & SEP_CHAR & 1 & END_CHAR)
                End If
                Player(Index).HookShotX = 0
                Player(Index).HookShotY = 0
            Else
                If Index = MyIndex Then
                Call SendData(POut.EndShot & SEP_CHAR & 0 & END_CHAR)
                End If
                Player(Index).HookShotX = 0
                Player(Index).HookShotY = 0
            End If
        End If

        BX = GetPlayerX(Index)
        BY = GetPlayerY(Index)

        If Player(Index).HookShotDir = DIR_DOWN Then
            Do While BY <= Player(Index).HookShotToY
                If BY <= MAX_MAPY Then
                    Call DD_BackBuffer.BltFast((BX - NewPlayerX) * PIC_X + sx - NewXOffset, (BY - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                BY = BY + 1
            Loop
        End If

        If Player(Index).HookShotDir = DIR_UP Then
            Do While BY >= Player(Index).HookShotToY
                If BY >= 0 Then
                    Call DD_BackBuffer.BltFast((BX - NewPlayerX) * PIC_X + sx - NewXOffset, (BY - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                BY = BY - 1
            Loop
        End If

        If Player(Index).HookShotDir = DIR_RIGHT Then
            Do While BX <= Player(Index).HookShotToX
                If BX <= MAX_MAPX Then
                    Call DD_BackBuffer.BltFast((BX - NewPlayerX) * PIC_X + sx - NewXOffset, (BY - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                BX = BX + 1
            Loop
        End If

        If Player(Index).HookShotDir = DIR_LEFT Then
            Do While BX >= Player(Index).HookShotToX
                If BX >= 0 Then
                    Call DD_BackBuffer.BltFast((BX - NewPlayerX) * PIC_X + sx - NewXOffset, (BY - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                BX = BX - 1
            Loop
        End If
    End If
End Sub

Sub BltArrow(ByVal Index As Long)
    Dim rec As DXVBLib.RECT
    Dim X As Long
    Dim y As Long
    Dim i As Long
    Dim z As Long

    For z = 1 To MAX_PLAYER_ARROWS
        If Player(Index).Arrow(z).Arrow > 0 Then
            rec.Top = Player(Index).Arrow(z).ArrowAnim * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = Player(Index).Arrow(z).ArrowPosition * PIC_X
            rec.Right = rec.Left + PIC_X

            If GetTickCount > Player(Index).Arrow(z).ArrowTime + 30 Then
                Player(Index).Arrow(z).ArrowTime = GetTickCount
                Player(Index).Arrow(z).ArrowVarX = Player(Index).Arrow(z).ArrowVarX + 10
                Player(Index).Arrow(z).ArrowVarY = Player(Index).Arrow(z).ArrowVarY + 10
            End If

            If Player(Index).Arrow(z).ArrowPosition = 0 Then
                X = Player(Index).Arrow(z).ArrowX
                y = Player(Index).Arrow(z).ArrowY + Int(Player(Index).Arrow(z).ArrowVarY / 32)

                If y > Player(Index).Arrow(z).ArrowY + Arrows(Player(Index).Arrow(z).ArrowNum).Range - 2 Then
                    Player(Index).Arrow(z).Arrow = 0
                End If

                If y <= MAX_MAPY Then
                    Call DD_BackBuffer.BltFast((Player(Index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset, (Player(Index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset + Player(Index).Arrow(z).ArrowVarY, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            End If

            If Player(Index).Arrow(z).ArrowPosition = 1 Then
                X = Player(Index).Arrow(z).ArrowX
                y = Player(Index).Arrow(z).ArrowY - Int(Player(Index).Arrow(z).ArrowVarY / 32)

                If y < Player(Index).Arrow(z).ArrowY - Arrows(Player(Index).Arrow(z).ArrowNum).Range + 2 Then
                    Player(Index).Arrow(z).Arrow = 0
                End If

                If y >= 0 Then
                    Call DD_BackBuffer.BltFast((Player(Index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset, (Player(Index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset - Player(Index).Arrow(z).ArrowVarY, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            End If

            If Player(Index).Arrow(z).ArrowPosition = 2 Then
                X = Player(Index).Arrow(z).ArrowX + Int(Player(Index).Arrow(z).ArrowVarX / 32)
                y = Player(Index).Arrow(z).ArrowY

                If X > Player(Index).Arrow(z).ArrowX + Arrows(Player(Index).Arrow(z).ArrowNum).Range - 2 Then
                    Player(Index).Arrow(z).Arrow = 0
                End If

                If X <= MAX_MAPX Then
                    Call DD_BackBuffer.BltFast((Player(Index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset + Player(Index).Arrow(z).ArrowVarX, (Player(Index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            End If

            If Player(Index).Arrow(z).ArrowPosition = 3 Then
                X = Player(Index).Arrow(z).ArrowX - Int(Player(Index).Arrow(z).ArrowVarX / 32)
                y = Player(Index).Arrow(z).ArrowY

                If X < Player(Index).Arrow(z).ArrowX - Arrows(Player(Index).Arrow(z).ArrowNum).Range + 2 Then
                    Player(Index).Arrow(z).Arrow = 0
                End If

                If X >= 0 Then
                    Call DD_BackBuffer.BltFast((Player(Index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset - Player(Index).Arrow(z).ArrowVarX, (Player(Index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            End If

            If X >= 0 And X <= MAX_MAPX Then
                If y >= 0 And y <= MAX_MAPY Then
                    If Map(GetPlayerMap(MyIndex)).Tile(X, y).Type = TILE_TYPE_BLOCKED Then
                        Player(Index).Arrow(z).Arrow = 0
                    End If
                End If
            End If

            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) Then
                    If GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                        If GetPlayerX(i) = X Then
                            If GetPlayerY(i) = y Then
                                If Index = MyIndex Then
                                    Call SendData(POut.ArrowHit & SEP_CHAR & 0 & SEP_CHAR & i & SEP_CHAR & X & SEP_CHAR & y & END_CHAR)
                                End If

                                If Index <> i Then
                                    Player(Index).Arrow(z).Arrow = 0
                                End If

                                Exit Sub
                            End If
                        End If
                    End If
                End If
            Next i

            For i = 1 To MAX_MAP_NPCS
                If MapNpc(i).num > 0 Then
                    If MapNpc(i).X = X Then
                        If MapNpc(i).y = y Then
                            If Index = MyIndex Then
                                Call SendData(POut.ArrowHit & SEP_CHAR & 1 & SEP_CHAR & i & SEP_CHAR & X & SEP_CHAR & y & END_CHAR)
                            End If

                            Player(Index).Arrow(z).Arrow = 0

                            Exit Sub
                        End If
                    End If
                End If
            Next i
        End If
    Next z
End Sub

Sub BltLevelUp(ByVal Index As Long)
    Dim rec As RECT
    Dim X As Integer
    Dim y As Integer

    If Player(Index).LevelUpT + 3000 > GetTickCount Then
        If GetPlayerMap(Index) = GetPlayerMap(MyIndex) Then
            rec.Top = PIC_Y * 2
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = PIC_X * 4
            rec.Right = rec.Left + 96

            X = GetPlayerX(Index) * PIC_X + Player(Index).xOffset + sx
            y = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset + sx

            Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - PIC_X - NewXOffset, y - (NewPlayerY * PIC_Y) - PIC_Y - NewYOffset - 8, DD_TileSurf(10), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

            If Player(Index).LevelUp >= 3 Then
                Player(Index).LevelUp = Player(Index).LevelUp - 1
            ElseIf Player(Index).LevelUp >= 1 Then
                Player(Index).LevelUp = Player(Index).LevelUp + 1
            End If
        Else
            Player(Index).LevelUpT = 0
        End If
    End If
End Sub

Sub BltSpriteChange(ByVal X As Long, ByVal y As Long)
    Dim rec As DXVBLib.RECT

    If Map(GetPlayerMap(MyIndex)).Tile(X, y).Type = TILE_TYPE_SPRITE_CHANGE Then
        If SpriteSize = 0 Then
            rec.Top = Map(GetPlayerMap(MyIndex)).Tile(X, y).Data1 * PIC_Y + 16
            rec.Bottom = rec.Top + PIC_Y - 16
        Else
            rec.Top = Map(GetPlayerMap(MyIndex)).Tile(X, y).Data1 * 64 + 16
            rec.Bottom = rec.Top + 64 - 16
        End If
        
        rec.Left = 96
        rec.Right = rec.Left + PIC_X

        X = X * PIC_X + sx
        y = y * PIC_Y + (sx / 2) '- 16

        If y < 0 Then
            Exit Sub
        End If

        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
End Sub

' New Visual Inventory [Mellowz]
Public Sub Inv_BltItem(ByVal InvNum As Long)
    Dim sRECT As DXVBLib.RECT
    Dim dRECT As DXVBLib.RECT
    Dim ItemNum As Long

    ItemNum = Player(MyIndex).Inv(InvNum).num

    If ItemNum = 0 Then
        frmMirage.picInv(InvNum).Picture = LoadPicture()
    Else
        sRECT.Top = Int(Item(ItemNum).Pic / 6) * PIC_Y
        sRECT.Bottom = sRECT.Top + PIC_Y
        sRECT.Left = (Item(ItemNum).Pic - Int(Item(ItemNum).Pic / 6) * 6) * PIC_X
        sRECT.Right = sRECT.Left + PIC_X

        dRECT.Top = 0
        dRECT.Bottom = PIC_Y
        dRECT.Left = 0
        dRECT.Right = PIC_X

        Call DD_ItemSurf.BltToDC(frmMirage.picInv(InvNum).hDC, sRECT, dRECT)
    End If
    
    frmMirage.picInv(InvNum).Refresh
End Sub

