'================================================================================
'   File:         Recorder.cls
'   Project:      SuperiorWaveFX32
'   Author:       Johnny Walker (y0_Superior)
'   Date:         2004
'   Description:  This class encapsulates all audio recording functionality. It
'                 captures raw audio from the microphone, converts it to the
'                 TrueSpeech format in real-time, and raises an event with
'                 the compressed data.
'================================================================================
Option Explicit

' Implements the event interface from the core DLL to handle stream events.
Implements superiorWavefx32.StreamEvent

' Define the DSP Group's TrueSpeech codec format tag.
Private Const TRUESPEECH = 34

' Core engine object from the custom DLL.
Private wfx As New superiorWavefx32.Wavefx

' WAVEFORMATEX structures to define the audio formats for capture (PCM)
' and the target compressed format (TrueSpeech).
Private fCaptureFormat As superiorWavefx32.WAVEFORMATEX
Private fTrueSpeechFormat As superiorWavefx32.WAVEFORMATEX

' Core objects for managing the audio stream.
Private fCaptureStream As superiorWavefx32.CaptureStream
Private fStreamConverter As superiorWavefx32.StreamConverter

' Event that will be raised when a chunk of audio has been successfully compressed.
Event onSoundCompressed(ByVal trueSpeechData As String, ByVal lBufferBytes As Long)


'--------------------------------------------------------------------------------
' Sub: Class_Initialize
' Desc: Constructor for the Recorder class. Sets up the required WAVEFORMATEX
'       structure for the TrueSpeech codec.
'--------------------------------------------------------------------------------
Private Sub Class_Initialize()
    ' Configure the TrueSpeech format attributes. These specific values are
    ' required by the DSP Group's codec to function correctly.
    With fTrueSpeechFormat
        .FormatTag = TRUESPEECH
        .channels = 1 'Mono
        .SamplesPerSec = 8000 '8khz sample rate
        .BitsPerSample = 1 'bit rate
        .BlockAlign = 32
        .AvgBytesPerSec = 1067
        .cbSize = 32

        ' The extraBytes are mandatory for the TrueSpeech ACM codec header.
        .extraBytes(0) = &H1
        .extraBytes(2) = &HF0
    End With
End Sub

'--------------------------------------------------------------------------------
' Sub: Record
' Desc: Initializes and starts the audio capture and conversion process.
'--------------------------------------------------------------------------------
Sub Record()
    ' Create a new capture stream object.
    Set fCaptureStream = New superiorWavefx32.CaptureStream
    
    ' Define the source audio format: 16-bit, 8kHz, Mono PCM.
    fCaptureFormat = wfx.createFormat(1, 8000, 16)
    
    ' Create and open the stream converter, specifying the input (PCM) and
    ' output (TrueSpeech) formats.
    Set fStreamConverter = New superiorWavefx32.StreamConverter
    fStreamConverter.streamOpen fCaptureFormat, fTrueSpeechFormat
    
    ' Configure and start the capture stream.
    With fCaptureStream
        ' Tell the stream which format to use, which object will handle its events (Me),
        ' and the buffer size.
        .setCaptureDescription fCaptureFormat, Me, 1440
        ' Start capturing from the default audio device (-1).
        Call .startCapture((-1))
    End With
End Sub

'--------------------------------------------------------------------------------
' Sub: EndRecord
' Desc: Stops the capture stream and releases all associated objects.
'--------------------------------------------------------------------------------
Sub EndRecord()
    Call fCaptureStream.stopCapture
    Set fCaptureStream = Nothing
    
    Call fStreamConverter.streamClose
    Set fStreamConverter = Nothing
End Sub

'--------------------------------------------------------------------------------
' Sub: StreamEvent_onCapture (Implements StreamEvent)
' Desc: This event handler is called automatically by the fCaptureStream object
'       whenever a new buffer of raw audio data is ready.
'--------------------------------------------------------------------------------
Private Sub StreamEvent_onCapture(waveBuffer() As Byte, lBytesCaptured As Long)
    Dim length As Long
    Dim wavData() As Byte
    
    ' Convert the raw PCM waveBuffer into the compressed TrueSpeech format.
    wavData = fStreamConverter.Convert(waveBuffer(), lBytesCaptured)
    length = UBound(wavData) - LBound(wavData)
    
    ' If the conversion produced data, raise our public event to send it
    ' to the main form.
    If (length > 0) Then RaiseEvent onSoundCompressed(StrConv(wavData, vbUnicode), length)
End Sub

'--------------------------------------------------------------------------------
' Function: StreamEvent_onWrite (Implements StreamEvent)
' Desc: Not implemented in this class, as this class only handles recording.
'--------------------------------------------------------------------------------
Private Function StreamEvent_onWrite(waveBuffer() As Byte, lBufferBytes As Long) As Long
    ' This function is required by the StreamEvent interface but is not used here.
End Function
