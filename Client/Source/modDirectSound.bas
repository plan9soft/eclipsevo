Attribute VB_Name = "modDirectSound"
Option Explicit

' The main DirectSound object.
Public DS As DirectSound

' The maximum amount of sounds.
Public Const MAX_SOUNDS As Long = 20

' Constant that holds the sound folder.
Public Const SOUND_DIR As String = "\SFX\"

' Type that defines the buffers capabilities.
Private Type BufferCaps
    Volume As Boolean
    Frequency As Boolean
    Pan As Boolean
End Type

' Type that defines the sound file.
Private Type SoundArray
    DSBuffer As DirectSoundBuffer
    DSCaps As BufferCaps
    DSFileName As String
End Type

' Contains all the data needed for sound manipulation.
Private Sound(1 To MAX_SOUNDS) As SoundArray

' Contains the current sound index.
Private SoundIndex As Long

' Is the sound engine initiated?
Private SEngineIsLoaded As Boolean


Public Sub DirectSound_Init()
    ' Create the DirectSound object.
    Set DS = DX.DirectSoundCreate(vbNullString)

    ' Check if any errors were raised when creating the object.
    If Err.Number <> 0 Then
        Call MsgBox("Failed to create the DirectSound object. Sound is disabled.")
        Exit Sub
    End If

    ' Set the cooperative level for the object.
    DS.SetCooperativeLevel frmMirage.hWnd, DSSCL_PRIORITY

    ' Successfully initiated the sound engine.
    SEngineIsLoaded = True
End Sub

Public Sub SoundLoad(ByVal FileName As String)
    Dim DSWave As WAVEFORMATEX
    Dim DSBufferDescription As DSBUFFERDESC

    ' Set the sound index one higher for each sound.
    SoundIndex = SoundIndex + 1
    
    ' Reset the sound array if the array height is reached.
    If SoundIndex > UBound(Sound) Then
        SoundIndex = 1
    End If

    ' Remove the sound if it exists (needed for re-loop).
    If SoundInMemory(SoundIndex) Then
        Call SoundStop(SoundIndex)
        Call SoundRemove(SoundIndex)
    End If

    ' Load the sound array with the data given.
    With Sound(SoundIndex)
        ' The file name that we want to load into memory.
        .DSFileName = FileName

        ' Is this sound capable of altered volume capabilities?
        .DSCaps.Volume = True

        ' Is this sound capable of altered frequency capabilities?
        .DSCaps.Frequency = True

        ' Is this sound capable of altered panning capabilities?
        .DSCaps.Pan = True
    End With

    DSWave.nFormatTag = WAVE_FORMAT_PCM 'Sound Must be PCM otherwise we get errors
    DSWave.nChannels = 2    '1= Mono, 2 = Stereo
    DSWave.lSamplesPerSec = 22050
    DSWave.nBitsPerSample = 16 '16 =16bit, 8=8bit
    DSWave.nBlockAlign = DSWave.nBitsPerSample / 8 * DSWave.nChannels
    DSWave.lAvgBytesPerSec = DSWave.lSamplesPerSec * DSWave.nBlockAlign

    ' Set the buffer description according to the data provided.
    With DSBufferDescription
        If Sound(SoundIndex).DSCaps.Frequency Then
            .lFlags = .lFlags Or DSBCAPS_CTRLFREQUENCY
        End If
        If Sound(SoundIndex).DSCaps.Pan Then
            .lFlags = .lFlags Or DSBCAPS_CTRLPAN
        End If
        If Sound(SoundIndex).DSCaps.Volume Then
            .lFlags = .lFlags Or DSBCAPS_CTRLVOLUME
        End If
    End With

    ' Load the sound file into the buffer.
    Set Sound(SoundIndex).DSBuffer = DS.CreateSoundBufferFromFile(App.Path & SOUND_DIR & Sound(SoundIndex).DSFileName, DSBufferDescription, DSWave)
End Sub

Public Sub SoundRemove(ByVal Index As Integer)
    'Reset all the variables in the sound array
    With Sound(Index)
        Set .DSBuffer = Nothing

        .DSCaps.Frequency = False
        .DSCaps.Pan = False
        .DSCaps.Volume = False
        .DSFileName = vbNullString
    End With
End Sub

Public Sub SoundPlay(ByVal FileName As String, Optional ByVal Volume As Long = 100, Optional ByVal Pan As Long = 50, Optional ByVal Frequency As Long = 11)
    ' Check to see if DirectSound was successfully initalized.
    If Not SEngineIsLoaded Then Exit Sub

    ' Check to see if the file exists.
    If Not FileExists("SFX\" & FileName) Then
        Call AddText("Warning: Couldn't find 'SFX\" & FileName & "'!", BRIGHTRED)
        Exit Sub
    End If

    ' Loads our sound into memory.
    Call SoundLoad(FileName)

    ' Sets the volume for the sound.
    Call SetVolume(SoundIndex, Volume)
    
    ' Sets the pan for the sound.
    Call SetPan(SoundIndex, Pan)
    
    ' Sets the frequency for the sound.
    Call SetFrequency(SoundIndex, Frequency)

    ' Play the sound.
    Sound(SoundIndex).DSBuffer.Play DSBPLAY_DEFAULT
End Sub

Public Sub SoundStop(ByVal Index As Integer)
    ' Stop the buffer and reset to the beginning
    Sound(Index).DSBuffer.Stop
    Sound(Index).DSBuffer.SetCurrentPosition 0
End Sub

Public Sub SoundPause(ByVal Index As Integer)
    ' Stop the buffer
    Sound(Index).DSBuffer.Stop
End Sub

Private Function SoundInMemory(ByVal Index As Integer) As Boolean
    ' Check if a sound is in memory.
    If Not Sound(Index).DSBuffer Is Nothing Then
        SoundInMemory = True
    End If
End Function

Public Sub SetFrequency(ByVal Index As Integer, ByVal Level As Long)
    ' Check to make sure that the buffer can alter its frequency.
    If Not Sound(Index).DSCaps.Frequency Then Exit Sub

    ' Alter the frequency according to the frequency provided.
    Select Case Level
        Case 0
            Sound(Index).DSBuffer.SetFrequency (DSBFREQUENCY_MIN)
        Case 100
            Sound(Index).DSBuffer.SetFrequency (DSBFREQUENCY_MAX)
        Case Else
            Sound(Index).DSBuffer.SetFrequency (Level * 1000)
    End Select
End Sub

Public Sub SetVolume(ByVal Index As Integer, ByVal Level As Long)
    ' Check to make sure that the buffer can alter its volume.
    If Not Sound(Index).DSCaps.Volume Then Exit Sub

    ' Alter the volume according to the volume provided.
    If Level > 0 Then
        Sound(Index).DSBuffer.SetVolume ((60 * Level) - 6000)
    Else
        Sound(Index).DSBuffer.SetVolume (-6000)
    End If
End Sub

Public Sub SetPan(ByVal Index As Integer, ByVal Level As Long)
    ' Check to make sure that the buffer can alter its pan.
    If Not Sound(Index).DSCaps.Pan Then Exit Sub

    ' Alter the pan according to the pan provided.
    Select Case Level
        Case 0
            Sound(Index).DSBuffer.SetPan -10000
        Case 100
            Sound(Index).DSBuffer.SetPan 10000
        Case Else
            Sound(Index).DSBuffer.SetPan (100 * Level) - 5000
    End Select
End Sub

Public Function GetFrequency(ByVal Index As Integer) As Long
    ' Check to make sure that the buffer can alter its frequency.
    If Not Sound(Index).DSCaps.Frequency Then Exit Function
    
    ' Return the frequency value.
    GetFrequency = Sound(Index).DSBuffer.GetFrequency()
End Function

Public Function GetVolume(ByVal Index As Integer) As Long
    ' Check to make sure that the buffer can alter its volume.
    If Not Sound(Index).DSCaps.Volume Then Exit Function
    
    ' Return the volume value.
    GetVolume = Sound(Index).DSBuffer.GetVolume()
End Function

Public Function GetPan(ByVal Index As Integer) As Long
    ' Check to make sure that the buffer can alter its pan.
    If Not Sound(Index).DSCaps.Pan Then Exit Function
    
    ' Return the pan value.
    GetPan = Sound(Index).DSBuffer.GetPan()
End Function

Public Function GetState(ByVal Index As Integer) As String
    ' Returns the current state of the given sound.
    GetState = Sound(Index).DSBuffer.GetStatus
End Function

Public Sub DestroyDirectSound()
    Dim I As Long

    ' Delete all of the sounds created.
    For I = 1 To UBound(Sound)
        If Not Sound(I).DSBuffer Is Nothing Then
            Call SoundStop(I)
            Call SoundRemove(I)
        End If
    Next I

    ' Disable any sounds from being played.
    SEngineIsLoaded = False
End Sub
