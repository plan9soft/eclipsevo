Attribute VB_Name = "modDirectMusic"
Option Explicit

' The master performance object.
Public DMPerf As DirectMusicPerformance

' The object that loads music into buffers.
Public DMLoader As DirectMusicLoader

' Stores the music in memory.
Public DMSegment As DirectMusicSegment

' Stores information about our segment.
Public DMState As DirectMusicSegmentState

' Defines our music folder.
Public Const MUSIC_DIR As String = "\Music\"

' Is the music engine loaded?
Public MEngineIsLoaded As Boolean

Public Sub DirectMusic_Init()
    ' Create the performance object.
    Set DMPerf = DX.DirectMusicPerformanceCreate()
    
    ' Create the loader object.
    Set DMLoader = DX.DirectMusicLoaderCreate()

    ' Check if any errors were raised when loading DirectMusic.
    If Err.Number <> 0 Then
        Call MsgBox("Failed to load the DirectMusic object. Music is disabled.")
        Exit Sub
    End If
        
    ' Initiate the audio device with the specified arguments.
    Call DMPerf.Init(Nothing, 0)

    ' Allow DirectMusic to download files streamed from the internet.
    Call DMPerf.SetPort(-1, 80)
    Call DMPerf.SetMasterAutoDownload(True)

    ' Enable the sound engine.
    MEngineIsLoaded = True
End Sub

Public Sub MusicLoad(ByVal FileName As String, Optional ByVal FilePath As String = vbNullString)
    ' Load the audio file into memory.
    Set DMSegment = DMLoader.LoadSegment(App.Path & MUSIC_DIR & FileName)

    ' Set the standard file type.
    DMSegment.SetStandardMidiFile
End Sub

Public Sub MusicPlay(ByVal FileName As String, Optional ByVal oVolume As Long = 100, Optional ByVal oTempo As Long = 100, Optional ByVal oGroove As Long = 0, Optional ByVal oLoop As Boolean = True)
    ' Check if the music engine is loaded.
    If Not MEngineIsLoaded Then Exit Sub
    
    ' Check to see if the file exists.
    If Not FileExistsNew(App.Path & MUSIC_DIR & FileName) Then
        Call AddText("Warning: Couldn't find 'BGM\" & FileName & "'!", BRIGHTRED)
        Exit Sub
    End If

    ' Load the music file into memory.
    Call MusicLoad(FileName)
    If oLoop Then
        ' Play the music track over and over.
        Call DMSegment.SetRepeats(-1)
    Else
        ' Player the music track once.
        Call DMSegment.SetRepeats(0)
    End If

    ' Set the volume level.
    Call SetTrackVolume(oVolume)

    ' Set the tempo level.
    Call SetTrackTempo(oTempo)

    ' Set the groove level.
    Call SetTrackGroove(oGroove)

    ' Play the segment in memory.
    Set DMState = DMPerf.PlaySegment(DMSegment, 0, 0)
End Sub

Public Sub MusicStop()
    ' Check if he music engine is loaded.
    If Not MEngineIsLoaded Then Exit Sub

    ' Stop the music currently playing.
    If Not DMSegment Is Nothing Then
        Call DMPerf.Stop(DMSegment, DMState, 0, 0)
    End If
End Sub

Public Sub SetTrackVolume(ByVal Level As Long)
    ' Check if the music engine is loaded.
    If Not MEngineIsLoaded Then Exit Sub

    ' Set the new volume level.
    Call DMPerf.SetMasterVolume(Level)
End Sub

Public Sub SetTrackTempo(ByVal Level As Long)
    ' Check if the music engine is loaded.
    If Not MEngineIsLoaded Then Exit Sub

    ' Set the new tempo level.
    Call DMPerf.SetMasterTempo(Level / 100)
End Sub

Public Sub SetTrackGroove(ByVal Level As Long)
    ' Check if the music engine is loaded.
    If Not MEngineIsLoaded Then Exit Sub

    ' Set the new groove level.
    Call DMPerf.SetMasterGrooveLevel(Level)
End Sub

Public Sub DestroyDirectMusic()
    ' Destroy the segment object.
    If ObjPtr(DMSegment) Then Set DMSegment = Nothing

    ' Destroy the loader object.
    If ObjPtr(DMLoader) Then Set DMLoader = Nothing

    ' Destroy the performance control.
    If Not (DMPerf Is Nothing) Then
        DMPerf.CloseDown
        Set DMPerf = Nothing
    End If

    ' Disable any music from being played.
    MEngineIsLoaded = False
End Sub


