Attribute VB_Name = "mPlayer"
Option Explicit
'Public variables for playback
'Const for channels
Private Const conNumChan As Long = 64
Private lngNumChan As Long
Private StreamHandle(conNumChan) As Long, ChannelHandle(conNumChan) As Long
Public lCurrentChannel As Long
Public bSlider As Boolean             '// arrastrando slider posbar
Private Enum pePlayer
    NotLoaded = 0
    Stopped = 1
    Playing = 2
    Paused = 3
End Enum

Private ePlayerState(conNumChan) As pePlayer

'Private variables for the spectrum
Private blnSpectrumOn As Boolean

'Private variables for FX
Private lngEQ(9) As Long, lngFX(8) As Long

'public variables for playback
Public strCurrentFile As String

'FMOD Functions
Public Sub FMOD_Initialize(Optional BufferSize As Long = 50, Optional MixerRate As Long = 44100, Optional MaxChannels As Long = 16, Optional Flags As FSOUND_INITMODES = 4, Optional Driver As FSOUND_OUTPUTTYPES = 2, Optional MixerType As FSOUND_MIXERTYPES = 4, Optional Device As Long = 0)
    On Error GoTo hell

    'Pre-initialize
    'These must be called before the initit
    FSOUND_SetDriver (Device)
    FSOUND_SetOutput (Driver)
    FSOUND_SetMixer (MixerType)

    FSOUND_SetBufferSize (100)

    If (MaxChannels > conNumChan) Then
        MsgBox "You can only have a maximum of " & conNumChan & " channels.", vbCritical + vbOKOnly, "Error"
        Exit Sub
    End If

    lngNumChan = MaxChannels
    'Initiate fmod
    If (FSOUND_Init(MixerRate, MaxChannels, Flags) = 0) Then    '1 or 2 or 4
        MsgBox "FMOD Cannot initialize because of an error." & _
               vbCrLf & (FSOUND_GetErrorString(FSOUND_GetError)), vbCritical + vbApplicationModal + vbOKOnly
    End If

    FSOUND_Stream_SetBufferSize (BufferSize)

    Exit Sub
hell:
    MsgBox "MAHESH MP3 Cannot initialize because of an error." & vbCrLf & err.Description, vbCritical + vbApplicationModal + vbOKOnly
End Sub


Public Function FMOD_GetCPU() As Single
    On Error Resume Next
    FMOD_GetCPU = FSOUND_GetCPUUsage
End Function

Public Function FMOD_GetNumChannels() As Long
    On Error Resume Next
    FMOD_GetNumChannels = lngNumChan
End Function

Public Function FMOD_GetMaxChannels() As Long
    On Error Resume Next
    FMOD_GetMaxChannels = conNumChan
End Function

Public Function FMOD_GetStatus(Channel As Long) As Boolean
    On Error Resume Next
    'Detect end of file
    If (Stream_GetPosition(Channel) = Stream_GetDuration(Channel)) Then
        FMOD_GetStatus = False
    Else
        FMOD_GetStatus = True
    End If

End Function

Public Function FMOD_Version() As Single
    On Error Resume Next
    FMOD_Version = FSOUND_GetVersion
End Function

Public Sub FMOD_Terminate()
    On Error Resume Next
    'Force FMOD to close
    FSOUND_Close
End Sub

Public Function Stream_GetState(Channel As Long) As Long
    On Error Resume Next
    'Check for channel bounds
    If (Channel > lngNumChan) Then Exit Function

    Stream_GetState = ePlayerState(Channel)
End Function

Public Sub Stream_SetFrequency(Channel As Long, Value As Long)
    On Error Resume Next
    'Check for channel bounds
    If (Channel > lngNumChan) Then Exit Sub

    FSOUND_SetFrequency Channel, Value
End Sub

Public Sub Stream_SetMute(Channel As Long, Value As Boolean)
    On Error Resume Next
    'Check for channel bounds
    If (Channel > lngNumChan) Then Exit Sub

    FSOUND_SetMute Channel, Value
End Sub

Public Sub Stream_Open(File As String, Optional Mode As FSOUND_MODES = 16, Optional Channel As Long = 0, Optional Play As Boolean = False, Optional Volume As Long = 255)
    On Error Resume Next
    'Check for channel bounds
    If (Channel > lngNumChan) Then Exit Sub

    strCurrentFile = File

    'Stop the current song
    Stream_Stop Channel

    'Open the new stream
    StreamHandle(Channel) = FSOUND_Stream_Open(File, Mode, 0, 0)

    'Change the playstate
    'Test to make sure the file is loaded
    If (StreamHandle(Channel) <> 0) Then
        ePlayerState(Channel) = Stopped
    End If

    If Play = True Then
        Stream_Play Channel
        Stream_SetVolume Channel, Volume
    End If
End Sub

Public Sub Stream_Play(Channel As Long)
    On Error Resume Next
    'Check for channel bounds
    If (Channel > lngNumChan) Then Exit Sub

    If (FSOUND_GetPaused(StreamHandle(Channel)) = 1) Then Stream_Pause Channel
    'Play the stream in a channel
    ChannelHandle(Channel) = FSOUND_Stream_Play(Channel, StreamHandle(Channel))

    'Change the playstate
    If (ChannelHandle(Channel) <> 0) Then
        ePlayerState(Channel) = Playing
    End If
End Sub

Public Sub Stream_Pause(Channel As Long)
    On Error Resume Next
    'Check for channel bounds
    If (Channel > lngNumChan) Then Exit Sub

    'Pause the stream
    If (FSOUND_GetPaused(ChannelHandle(Channel)) = False) Then
        FSOUND_SetPaused ChannelHandle(Channel), True

        'Stream is paused
        If (StreamHandle(Channel) <> 0) Then
            ePlayerState(Channel) = Paused
        End If
    Else

        FSOUND_SetPaused ChannelHandle(Channel), False

        'Stream is playing
        If (StreamHandle(Channel) <> 0) Then
            ePlayerState(Channel) = Playing
        End If
    End If


End Sub

Public Sub Stream_Stop(Channel As Long)
    On Error Resume Next
    'Check for channel bounds
    If (Channel > lngNumChan) Then Exit Sub

    If (StreamHandle(Channel) <> 0) Then

        'stop all sounds
        FSOUND_Stream_Stop StreamHandle(Channel)

        StreamHandle(Channel) = 0

        ePlayerState(Channel) = Stopped

    End If

End Sub

Public Sub Stream_SetVolume(Channel As Long, Value As Long)
    On Error Resume Next
    'Check for channel bounds
    If (Channel > lngNumChan) Then Exit Sub
    '// volumen 0 -> 255
    FSOUND_SetVolume ChannelHandle(Channel), Value

End Sub

Public Function Stream_GetVolume(Channel As Long) As Long
    On Error Resume Next
    If (Channel > lngNumChan) Then Exit Function
    Stream_GetVolume = FSOUND_GetVolume(Channel)
End Function

Public Sub Stream_SetBalance(Channel As Long, Value As Long)
    On Error Resume Next
    'Check for channel bounds
    If (Channel > lngNumChan) Then Exit Sub
    'Set balance
    FSOUND_SetPan ChannelHandle(Channel), Value

End Sub

Public Function Stream_GetPosition(Channel As Long) As Long
    On Error Resume Next
    'Check for channel bounds
    If (Channel > lngNumChan) Then Exit Function
    'Current position
    Stream_GetPosition = FSOUND_Stream_GetTime(StreamHandle(Channel)) / 1000

End Function

Public Sub Stream_SetPosition(Channel As Long, PosValue As Long)
    On Error Resume Next
    'Check for channel bounds
    If (Channel > lngNumChan) Then Exit Sub
    'Current position
    FSOUND_Stream_SetTime (StreamHandle(Channel)), PosValue * 1000

End Sub

Public Function Stream_GetDuration(Channel As Long) As Long
    On Error Resume Next
    'Check for channel bounds
    If (Channel > lngNumChan) Then Exit Function
    'Length
    Stream_GetDuration = FSOUND_Stream_GetLengthMs(StreamHandle(Channel)) / 1000

End Function


Public Sub FX_SetChorus(sngWetDryMix As Single, sngDepth As Single, sngFeedback As Single, sngFrequency As Single, lngWaveform As Long, sngDelay As Single, lngPhase As Long)
    On Error Resume Next
    Call FSOUND_FX_SetChorus(lngFX(0), sngWetDryMix, sngDepth, sngFeedback, sngFrequency, lngWaveform, sngDelay, lngPhase)
End Sub

Public Sub FX_SetCompressor(sngGain As Single, sngAttack As Single, sngRelease As Single, sngThreshold As Single, sngRatio As Single, sngPreDelay As Single)
    On Error Resume Next
    Call FSOUND_FX_SetCompressor(lngFX(1), sngGain, sngAttack, sngRelease, sngThreshold, sngRatio, sngPreDelay)
End Sub

Public Sub FX_SetDistortion(sngGain As Single, sngEdge As Single, sngPostEQCenterFrequency As Single, sngPostEQBandwidth As Single, sngPreLowpassCuttoff As Single)
    On Error Resume Next
    Call FSOUND_FX_SetDistortion(lngFX(2), sngGain, sngEdge, sngPostEQCenterFrequency, sngPostEQBandwidth, sngPreLowpassCuttoff)
End Sub

Public Sub FX_SetEcho(sngWetDryMix As Single, sngFeedback As Single, sngLeftDelay As Single, sngRightDelay As Single, sngPanDelay As Long)
    On Error Resume Next
    Call FSOUND_FX_SetEcho(lngFX(3), sngWetDryMix, sngFeedback, sngLeftDelay, sngRightDelay, sngPanDelay)
End Sub

Public Sub FX_SetFlanger(sngWetDryMix As Single, sngDepth As Single, sngFeedback As Single, sngFrequency As Single, lngWaveform As Long, sngDelay As Single, lngPhase As Long)
    On Error Resume Next
    Call FSOUND_FX_SetFlanger(lngFX(4), sngWetDryMix, sngDepth, sngFeedback, sngFrequency, lngWaveform, sngDelay, lngPhase)
End Sub

Public Sub FX_SetGargle(lngRateHz As Long, lngWaveShape As Long)
    On Error Resume Next
    Call FSOUND_FX_SetGargle(lngFX(5), lngRateHz, lngWaveShape)
End Sub

Public Sub FX_SetI3DL2Reverb(lngRoom As Long, lngRoomHF As Long, sngRoomRolloffFactor As Single, sngDecayTime As Single, sngDecayHFRatio As Single, lngReflections As Long, sngReflectionsDecay As Single, lngReverb As Long, sngReverbDelay As Single, sngDiffusion As Single, sngDensity As Single, sngHFReference As Single)
    On Error Resume Next
    Call FSOUND_FX_SetI3DL2Reverb(lngFX(6), lngRoom, lngRoomHF, sngRoomRolloffFactor, sngDecayTime, sngDecayHFRatio, lngReflections, sngReflectionsDecay, lngReverb, sngReverbDelay, sngDiffusion, sngDensity, sngHFReference)
End Sub

Public Sub FX_SetEQ(lngIndex As Long, sngValue As Single)
    On Error Resume Next

    'Define variables
    Dim sngGain As Single, sngCenter As Single

    Select Case lngIndex
    Case 0: sngCenter = 80
    Case 1: sngCenter = 170
    Case 2: sngCenter = 310
    Case 3: sngCenter = 600
    Case 4: sngCenter = 1000
    Case 5: sngCenter = 3000
    Case 6: sngCenter = 6000
    Case 7: sngCenter = 12000
    Case 8: sngCenter = 14000
    Case 9: sngCenter = 16000
    End Select
    'Get gain
    sngGain = sngValue
    'Set EQ
    Call FSOUND_FX_SetParamEQ(lngEQ(lngIndex), sngCenter, 18, sngGain)
    '    Call FSOUND_FX_SetParamEQ(lngFX(7), sngCenter, 18, sngGain)

End Sub

Public Sub FX_SetWavesReverb(sngInGain As Single, sngReverbMix As Single, sngReverbTime As Single, sngHighFrequencyRTRatio As Single)
    On Error Resume Next
    Call FSOUND_FX_SetWavesReverb(lngFX(8), sngInGain, sngReverbMix, sngReverbTime, sngHighFrequencyRTRatio)
End Sub

Public Sub FX_Enable(fX As FSOUND_FX_MODES, Optional Channel As Long = -1000)
    On Error Resume Next
    Dim intX As Integer

    If (Channel > lngNumChan) Then Exit Sub

    'The channel must be paused before we can change it
    FSOUND_SetPaused Channel, True

    'Set the Eq FX
    If (fX = FSOUND_FX_PARAMEQ) Then
        For intX = 0 To 9
            lngEQ(intX) = FSOUND_FX_Enable(Channel, FSOUND_FX_PARAMEQ)
        Next intX
    Else
        'Set up the FX
        lngFX(fX) = FSOUND_FX_Enable(Channel, fX)
    End If


    'Unpause
    FSOUND_SetPaused Channel, False
End Sub

Public Sub FX_Disable(Optional Channel As Long = -1000)
    On Error Resume Next
    If (Channel > lngNumChan) Then Exit Sub

    'The channel must be paused before we can change it
    FSOUND_SetPaused Channel, True
    FSOUND_FX_Disable Channel
    FSOUND_SetPaused Channel, False

End Sub

Public Function Spectrum_GetLeft(Channel As Long) As Single
    On Error Resume Next
    Dim sngVolLeft As Single, sngVolRight As Single

    'Check for channel bounds
    If (Channel < 0) Or (Channel > lngNumChan) Then Exit Function
    FSOUND_GetCurrentLevels Channel, sngVolLeft, sngVolRight

    Spectrum_GetLeft = sngVolLeft
End Function

Public Function Spectrum_GetRight(Channel As Long) As Single
    On Error Resume Next
    Dim sngVolLeft As Single, sngVolRight As Single

    'Check for channel bounds
    If (Channel < 0) Or (Channel > lngNumChan) Then Exit Function
    FSOUND_GetCurrentLevels Channel, sngVolLeft, sngVolRight

    Spectrum_GetRight = sngVolRight
End Function

Public Sub Spectrum_Enable(Value As Boolean)
    On Error Resume Next
    'Spectrum
    FSOUND_DSP_SetActive FSOUND_DSP_GetFFTUnit, Value

    'Reset the display
    blnSpectrumOn = Value
End Sub

Public Function Spectrum_GetData(Channel As Long, sngData() As Single)
    On Error Resume Next
    'Retreive spectrum data on if playing
    'Check for channel bounds
    If (Channel > lngNumChan) Then Exit Function

    'If blnSpectrumOn = False Then Exit Function

    If (ChannelHandle(Channel) <> 0) Then
        GetSpectrum sngData
    End If

End Function


'Driver functions
Public Function Driver_GetNum() As Long
    Driver_GetNum = FSOUND_GetNumDrivers
End Function

Public Function Driver_GetName(intDriverNum As Integer) As String
    Driver_GetName = GetStringFromPointer(FSOUND_GetDriverName(intDriverNum))
End Function

