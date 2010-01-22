Attribute VB_Name = "modSound"
Option Explicit

Public Sub MapMusic(ByVal Song As String)
    If Not Map(GetPlayerMap(MyIndex)).music = CurrentSong Then
        Call PlayBGM(Map(GetPlayerMap(MyIndex)).music)
    End If
End Sub

Public Sub PlayBGM(Song As String)
    If ReadINI("CONFIG", "Music", App.Path & "\config.ini") = 1 Then
        If FileExists("\Music\" & Song) Then
            If Not LenB(Song) = 0 Then
                If Not Left$(Song, 7) = "http://" Then
                    Call MusicPlay(Song)

                    CurrentSong = Song
                Else
                    ' Currently, we don't support music that can be downloaded.
                End If
            End If
        Else
            Call AddText(Song & " does not exist!", BRIGHTRED)
        End If
    End If
End Sub

Public Sub PlaySound(Sound As String)
    If ReadINI("CONFIG", "Sound", App.Path & "\config.ini") = 1 Then
        If FileExists("\SFX\" & Sound) Then
            'Call frmMirage.SoundPlayer.PlayMedia(App.Path & "\SFX\" & Sound, False)
            Call SoundPlay(Sound)
        Else
            Call AddText(Sound & " does not exist!", BRIGHTRED)
        End If
    End If
End Sub

Public Sub PlayBGS(Sound As String)
    If ReadINI("CONFIG", "Sound", App.Path & "\config.ini") = 1 Then
        If FileExists("\SFX\" & Sound) Then
            Call SoundPlay(Sound)
            'Call frmMirage.BGSPlayer.PlayMedia(App.Path & "\BGS\" & Sound, True)
        Else
            Call AddText(Sound & " does not exist!", BRIGHTRED)
        End If
    End If
End Sub

Public Sub StopBGM()
    Call MusicStop
    CurrentSong = vbNullString
End Sub

Public Sub StopSound()
    Dim I As Long
    
    For I = 1 To MAX_SOUNDS
        Call SoundStop(I)
        Call SoundRemove(I)
    Next I
End Sub
