'================================================================================
'   File:         Player.cls
'   Project:      SuperiorWaveFX32
'   Author:       Johnny Walker (y0_Superior)
'   Date:         2004
'   Description:  This class encapsulates all audio playback functionality. It
'                 receives compressed TrueSpeech data, decompresses it back
'                 to raw PCM in real-time, and plays it through the speakers.
'================================================================================
Option Explicit

' Event raised when the playback buffer is empty and all sounds have finished playing.
Event onSoundComplete()

' Implements the event interface from the core DLL to handle stream events.
Implements superiorWavefx32.StreamEvent

' Define the DSP Group's TrueSpeech codec format tag.
Private Const TRUESPEECH = &H22

' Core engine objects from the custom DLL.
Private wfx As New superiorWavefx32.Wavefx
Private fSoundBuffer As New superiorWavefx32.StreamIO ' A custom buffer for incoming sound data.

' WAVEFORMATEX structures to define the audio formats for playback (PCM)
' and the source compressed format (TrueSpeech).
Private fSoundFormat As superiorWavefx32.WAVEFORMATEX
Private fTrueSpeechFormat As superiorWavefx32.WAVEFORMATEX

' Core objects for managing the audio stream.
Private fSoundStream As superiorWavefx32.SoundStream
Private fStreamConverter As superiorWavefx32.StreamConverter


'--------------------------------------------------------------------------------
' Sub: Class_Initialize
' Desc: Constructor for the Player class. Sets up the required WAVEFORMATEX
'       structure for the TrueSpeech codec.
'--------------------------------------------------------------------------------
Private Sub Class_Initialize()
    ' Configure the TrueSpeech format attributes.
    With fTrueSpeechFormat
        .FormatTag = TRUESPEECH
        .channels = 1
        .SamplesPerSec = 8000
        .BitsPerSample = 1
        .BlockAlign = 32
        .AvgBytesPerSec = 1067
        .cbSize = 32
        
        ' Mandatory extra bytes for the codec header.
        .extraBytes(0) = &H1
        .extraBytes(2) = &HF0
    End With
End Sub

'--------------------------------------------------------------------------------
' Sub: Initalize
' Desc: Prepares the player for receiving audio by setting up the sound stream
'       and the format converter.
'--------------------------------------------------------------------------------
Sub Initalize()
    ' Create the main sound stream object.
    Set fSoundStream = New superiorWavefx32.SoundStream
    
    ' Define the target audio format for playback: 16-bit, 8kHz, Mono PCM.
    fSoundFormat = wfx.createFormat(1, 8000, 16)
    
    ' Configure the sound stream for playback.
    fSoundStream.setSoundDescription fSoundFormat, Me, 1440
    
    ' Create the stream converter and open it to convert from TrueSpeech (input)
    ' back to PCM (output).
    Set fStreamConverter = New superiorWavefx32.StreamConverter
    fStreamConverter.streamOpen fTrueSpeechFormat, fSoundFormat
End Sub

'--------------------------------------------------------------------------------
' Sub: PlayWave
' Desc: Receives a string of compressed TrueSpeech data, adds it to the buffer,
'       and begins the playback process if it's not already running.
'--------------------------------------------------------------------------------
Sub PlayWave(ByVal trueSpeechData As String)
    Dim wavData() As Byte
    ' Convert the incoming string data back to a byte array.
    wavData = StrConv(trueSpeechData, vbFromUnicode)
    
    ' Write the compressed data into our custom buffer.
    fSoundBuffer.Write_ wavData(), UBound(wavData) - LBound(wavData) + 1, 0
    
    ' If playback isn't already active, begin writing to the sound device.
    If (fSoundBuffer.chunkSize = 4) Then fSoundStream.beginWrite
End Sub

'--------------------------------------------------------------------------------
' Sub: Class_Terminate
' Desc: Destructor for the Player class. Ensures the sound stream is closed.
'--------------------------------------------------------------------------------
Private Sub Class_Terminate()
    fSoundStream.closeSound
    Set fSoundStream = Nothing
End Sub

'--------------------------------------------------------------------------------
' Sub: StreamEvent_onCapture (Implements StreamEvent)
' Desc: Not implemented in this class, as it only handles playback.
'--------------------------------------------------------------------------------
Private Sub StreamEvent_onCapture(waveBuffer() As Byte, lBytesCaptured As Long)
    ' Required by the interface, but not used here.
End Sub

'--------------------------------------------------------------------------------
' Function: StreamEvent_onWrite (Implements StreamEvent)
' Desc: This event is called automatically by fSoundStream when it needs more
'       PCM data to play. This function pulls compressed data from our buffer,
'       converts it, and supplies it to the playback stream.
'--------------------------------------------------------------------------------
Private Function StreamEvent_onWrite(waveBuffer() As Byte, lBufferBytes As Long) As Long
    Dim tsWavData() As Byte
    Dim waveLength As Long
    
    ' Check if there is any data left in our buffer to play.
    If (fSoundBuffer.chunkSize < 1) Then
        ' If not, clear the buffer, signal completion, and tell the stream
        ' there is no more data (return 0).
        fSoundBuffer.Clear
        RaiseEvent onSoundComplete
        StreamEvent_onWrite = 0
    Else
        ' If there is data, read the next chunk of compressed data from our buffer.
        waveLength = fSoundBuffer.Read_(tsWavData(), 0)
        ' Convert the TrueSpeech chunk back to a PCM byte array.
        waveBuffer = fStreamConverter.Convert(tsWavData(), waveLength)
        
        ' Update the byte count for the playback stream.
        lBufferBytes = UBound(waveBuffer) - LBound(waveBuffer) + 1
        
        ' Tell the sound stream that we have successfully provided data (return 1).
        StreamEvent_onWrite = 1
    End If
End Function
