Attribute VB_Name = "mSound"
Option Explicit

'19 apr 2006 UMG
'added overlap capability
'pre-calculated a few values
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'15 apr 2006 UMG
'changed to double sound buffering because FFT is still not fast enough for high ADC sample rates
'added code for checking device caps

Private Declare Function waveInGetNumDevs Lib "winmm.dll" () As Long
Private Declare Function waveInGetDevCaps Lib "winmm.dll" Alias "waveInGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As WAVEInCaps, ByVal uSize As Long) As Long
Private Declare Function waveInReset Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Private Declare Function waveInOpen Lib "winmm.dll" (ByRef lphWaveIn As Long, ByVal uDeviceID As Long, lpFormat As WAVEFORMAT, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Private Declare Function waveInClose Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Private Declare Function waveInStart Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Private Declare Function waveInStop Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
Private Declare Function waveInAddBuffer Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare Function waveInPrepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare Function waveInUnprepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, lpWaveInHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare Function waveInGetErrorText Lib "winmm.dll" Alias "waveInGetErrorTextA" (ByVal nErr As Long, ByVal lpText As String, ByVal uSize As Long) As Long
Private hWaveIn                 As Long
Private WaveErrMsg              As String
Private Const WHDR_DONE         As Long = 1
Private Const MMSYSERR_NOERROR  As Long = 0
Private Const WAVE_FORMAT_PCM   As Long = 1

Private Const MAXPNAMELEN       As Long = 32
Private Type WAVEInCaps
    ManufacturerID As Integer
    ProductID As Integer
    DriverVersion As Long
    ProductName(1 To MAXPNAMELEN) As Byte
    Formats As Long
    Channels As Integer
    Reserved As Integer
End Type
Private WvCaps                  As WAVEInCaps

Private DvcId                   As Long

Private Enum WvFormats
    WAVE_FORMAT_1M08 = &H1    '11.025 kHz  Mono     8-bit
    WAVE_FORMAT_1S08 = &H2    '11.025 kHz  Stereo   8-bit
    WAVE_FORMAT_1M16 = &H4    '11.025 kHz  Mono    16-bit
    WAVE_FORMAT_1S16 = &H8    '11.025 kHz  Stereo  16-bit
    WAVE_FORMAT_2M08 = &H10   '22.05  kHz  Mono     8-bit
    WAVE_FORMAT_2S08 = &H20   '22.05  kHz  Stereo   8-bit
    WAVE_FORMAT_2M16 = &H40   '22.05  kHz  Mono    16-bit
    WAVE_FORMAT_2S16 = &H80   '22.05  kHz  Stereo  16-bit
    WAVE_FORMAT_4M08 = &H100  '44.1   kHz  Mono     8-bit
    WAVE_FORMAT_4S08 = &H200  '44.1   kHz  Stereo   8-bit
    WAVE_FORMAT_4M16 = &H400  '44.1   kHz  Mono    16-bit
    WAVE_FORMAT_4S16 = &H800  '44.1   kHz  Stereo  16-bit
End Enum
#If False Then
Private WAVE_FORMAT_1M08, WAVE_FORMAT_1S08, WAVE_FORMAT_1M16, WAVE_FORMAT_1S16, WAVE_FORMAT_2M08, WAVE_FORMAT_2S08, WAVE_FORMAT_2M16, WAVE_FORMAT_2S16, _
        WAVE_FORMAT_4M08, WAVE_FORMAT_4S08, WAVE_FORMAT_4M16, WAVE_FORMAT_4S16
#End If
Private Const WhatWeNeed        As Long = WAVE_FORMAT_1M16 Or _
                                          WAVE_FORMAT_2M16 Or _
                                          WAVE_FORMAT_4M16
Private Type WAVEHDR
    lpData          As Long
    dwBufferLength  As Long
    dwBytesRecorded As Long
    dwUser          As Long
    dwFlags         As Long
    dwLoops         As Long
    lpNext          As Long
    Reserved        As Long
End Type
Private WvHdr(0 To 1)           As WAVEHDR

Private Type WAVEFORMAT
    wFormatTag      As Integer
    nChannels       As Integer
    nSamplesPerSec  As Long
    nAvgBytesPerSec As Long
    nBlockAlign     As Integer
    wBitsPerSample  As Integer
    cbSize          As Integer
End Type
Private WvFmt                   As WAVEFORMAT

Private Const sSI   As String = "Sound Input"
Private Const sUn   As String = "Unfortunately you have no "
Private Const sDv   As String = " Device"

Private Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private hMem(0 To 1)            As Long

Public PtrOverlap               As Long
Private NumSamples              As Long
Private n                       As Long
Private BufferSelector          As Long
Private aFirst                  As Long
Private aSecond                 As Long
Private LenHdr                  As Long
Private Corr                    As Double
Private Sum                     As Double
Private SoundSample             As Double

Private mySoundSamples()        As Integer

Public Function SoundBufferIsReady() As Boolean

    With WvHdr(BufferSelector)
        If .dwFlags And WHDR_DONE Then 'this buffer was filled and dequeued by winmm
            MemCopy aFirst, aSecond, .dwBufferLength 'copy my second to my first buffer
            MemCopy aSecond, .lpData, .dwBufferLength 'get sound data from filled buffer into my second buffer
            waveInAddBuffer hWaveIn, WvHdr(BufferSelector), LenHdr 'enqueue "this" buffer again
            BufferSelector = BufferSelector Xor 1 'next time "other" buffer
            PtrOverlap = PtrOverlap - NumSamples 'adjust pointer
            If PtrOverlap < 1 Then 'this can happen if overlap stride was too small
                PtrOverlap = 1
            End If
        End If
    End With 'WVHDR(BUFFERSELECTOR)
    If PtrOverlap <= NumSamples Then
        SoundBufferIsReady = True
        Corr = Corr + Sum / NumSamples 'ADC zero offset correction
        Sum = 0
    End If

End Function

Public Function SoundCheckDevice() As Boolean

    n = waveInGetNumDevs
    If n < 1 Then
        MsgBox sUn & sSI & sDv, vbCritical, sSI
      Else 'NOT N...
        With WvCaps
            Do
                n = n - 1
                waveInGetDevCaps n, WvCaps, Len(WvCaps)
                If .Formats And WhatWeNeed Then
                    MsgBox "Going to use " & StrConv(.ProductName, vbUnicode), vbInformation, sSI
                    DvcId = n
                    SoundCheckDevice = True
                    Exit Do 'loop 
                End If
            Loop While n
        End With 'WVCAPS
        If Not SoundCheckDevice Then
            MsgBox sUn & "suitable " & sSI & sDv, vbCritical, sSI
        End If
    End If

End Function

Public Function SoundGetSample(ByVal SampleNum As Long, ByVal Weight As Long) As Double

    SoundSample = mySoundSamples(SampleNum) - Corr 'ADC zero offset correction
    Sum = Sum + SoundSample
    SoundGetSample = SoundSample / Weight

End Function

Public Function SoundStartRecording(ByVal BufferSize As Long, ByVal SamplingRate As Long) As Boolean

    If hMem(0) Then
        SoundStopRecording
    End If

    NumSamples = BufferSize

    With WvFmt
        .wFormatTag = WAVE_FORMAT_PCM
        .nChannels = 1
        .wBitsPerSample = 16
        .nSamplesPerSec = SamplingRate
        .nBlockAlign = .nChannels * .wBitsPerSample / 8
        .nAvgBytesPerSec = .nSamplesPerSec * .nBlockAlign
        .cbSize = Len(WvFmt)
    End With 'WVFMT

    'create two buffer headers and allocate buffer memory
    For n = 0 To 1
        With WvHdr(n)
            .dwBufferLength = NumSamples * WvFmt.nBlockAlign
            hMem(n) = GlobalAlloc(0, .dwBufferLength) 'get buffer memory
            .lpData = GlobalLock(hMem(n))
            .dwFlags = 0
            .dwLoops = 0
        End With 'WVHDR(N)
    Next n

    ReDim mySoundSamples(1 To NumSamples * 2) 'my own two buffers for overlap
    BufferSelector = 0
    aFirst = VarPtr(mySoundSamples(1))
    aSecond = VarPtr(mySoundSamples(NumSamples + 1))
    PtrOverlap = NumSamples * 2 + 1

    LenHdr = Len(WvHdr(0))
    n = waveInOpen(hWaveIn, DvcId, WvFmt, 0, 0, 0) 'open sound
    If n = MMSYSERR_NOERROR Then
        n = waveInPrepareHeader(hWaveIn, WvHdr(0), LenHdr) 'prepare 1st buffer
        If n = MMSYSERR_NOERROR Then
            n = waveInAddBuffer(hWaveIn, WvHdr(0), LenHdr) 'enqueue 1st buffer
            If n = MMSYSERR_NOERROR Then
                n = waveInPrepareHeader(hWaveIn, WvHdr(1), LenHdr) 'prepare 2nd buffer
                If n = MMSYSERR_NOERROR Then
                    n = waveInAddBuffer(hWaveIn, WvHdr(1), LenHdr) 'enqueue 2nd buffer
                    If n = MMSYSERR_NOERROR Then
                        n = waveInStart(hWaveIn) 'get going
                    End If
                End If
            End If
        End If
    End If
    If n = MMSYSERR_NOERROR Then
        SoundStartRecording = True
        Sum = 0
      Else 'NOT n...
        WaveErrMsg = Space$(256)
        waveInGetErrorText n, WaveErrMsg, Len(WaveErrMsg)
        SoundStopRecording 'oops - something wrong
        MsgBox WaveErrMsg, vbCritical, sSI
    End If

End Function

Public Sub SoundStopRecording()

    waveInReset hWaveIn 'stop and dequeue buffers
    For n = 0 To 1
        waveInUnprepareHeader hWaveIn, WvHdr(n), LenHdr 'discard buffers
        If hMem(n) Then
            GlobalFree hMem(n) 'free buffer memory
            hMem(n) = 0
        End If
    Next n
    waveInClose hWaveIn 'close sound
    hWaveIn = 0

End Sub

':) Ulli's VB Code Formatter V2.21.6 (2006-Apr-24 11:50)  Decl: 107  Code: 137  Total: 244 Lines
':) CommentOnly: 8 (3,3%)  Commented: 40 (16,4%)  Empty: 30 (12,3%)  Max Logic Depth: 6
