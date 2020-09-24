VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWavConcatenate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'**********************************
' Wave concatenation class
'
' [rm_code] 2005
'**********************************

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" ( _
        Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function acmStreamPrepareHeader Lib "msacm32" (ByVal has As Long, _
        pash As ACMSTREAMHEADER, ByVal fdwPrepare As Long) As Long

Private Declare Function acmStreamUnprepareHeader Lib "msacm32" (ByVal has As Long, _
        pash As ACMSTREAMHEADER, ByVal fdwUnprepare As Long) As Long

Private Declare Function acmStreamOpen Lib "msacm32" (ByRef phas As Long, _
        ByVal had As Long, pwfxSrc As Any, pwfxDst As Any, ByVal pwfltr As Long, _
        ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long

Private Declare Function acmStreamSize Lib "msacm32" (ByVal has As Long, ByVal cbInput As Long, _
        pdwOutputBytes As Long, ByVal fdwSize As Long) As Long

Private Declare Function acmStreamConvert Lib "msacm32" (ByVal has As Long, _
        pash As ACMSTREAMHEADER, ByVal fdwConvert As Long) As Long

Private Declare Function acmStreamReset Lib "msacm32.dll" (ByVal has As Long, _
        ByVal fdwReset As Long) As Long

Private Declare Function acmStreamClose Lib "msacm32" (ByVal has As Long, _
        ByVal fdwClose As Long) As Long

Private Declare Function acmMetrics Lib "msacm32" (ByVal hao As Long, _
        ByVal uMetric As Integer, pMetric As Any) As Long

Private Declare Function mmioClose Lib "winmm.dll" ( _
        ByVal hMMIO As Long, ByVal uFlags As Long) As Long

Private Declare Function mmioDescend Lib "winmm.dll" ( _
        ByVal hMMIO As Long, lpck As MMCKINFO, _
        lpckParent As MMCKINFO, ByVal uFlags As Long) As Long

Private Declare Function mmioDescendParent Lib "winmm.dll" _
        Alias "mmioDescend" (ByVal hMMIO As Long, lpck As MMCKINFO, _
        ByVal x As Long, ByVal uFlags As Long) As Long

Private Declare Function mmioOpen Lib "winmm.dll" _
        Alias "mmioOpenA" (ByVal szFileName As String, _
        lpmmioinfo As mmioinfo, ByVal dwOpenFlags As Long) As Long

Private Declare Function mmioSeek Lib "winmm.dll" ( _
        ByVal hMMIO As Long, ByVal lOffset As Long, _
        ByVal iOrigin As Long) As Long

Private Declare Function mmioStringToFOURCC Lib "winmm.dll" _
        Alias "mmioStringToFOURCCA" (ByVal sz As String, _
        ByVal uFlags As Long) As Long

Private Type mmioinfo
   dwFlags                      As Long
   fccIOProc                    As Long
   pIOProc                      As Long
   wErrorRet                    As Long
   htask                        As Long
   cchBuffer                    As Long
   pchBuffer                    As String
   pchNext                      As String
   pchEndRead                   As String
   pchEndWrite                  As String
   lBufOffset                   As Long
   lDiskOffset                  As Long
   adwInfo(4)                   As Long
   dwReserved1                  As Long
   dwReserved2                  As Long
   hMMIO                        As Long
End Type

Private Type MMCKINFO
   ckid                         As Long
   ckSize                       As Long
   fccType                      As Long
   dwDataOffset                 As Long
   dwFlags                      As Long
End Type

Private Type ACMSTREAMHEADER
    cbStruct                    As Long
    fdwStatus                   As Long
    dwUser                      As Long
    pbSrc                       As Long
    cbSrcLength                 As Long
    cbSrcLengthUsed             As Long
    dwSrcUser                   As Long
    pbDst                       As Long
    cbDstLength                 As Long
    cbDstLengthUsed             As Long
    dwDstUser                   As Long
    dwReservedDriver(9)         As Long
End Type

Private Type WAVEFORMATEX
    wFormatTag                  As Integer
    nChannels                   As Integer
    nSamplesPerSec              As Long
    nAvgBytesPerSec             As Long
    nBlockAlign                 As Integer
    wBitsPerSample              As Integer
    cbSize                      As Integer
End Type

Private Type WaveFileHeader
    RiffID                      As Long
    len                         As Long
    WaveID                      As Long
    FmtID                       As Long
    HdrLen                      As Long
    WaveFormat                  As WAVEFORMATEX
End Type

Private Type WaveFile
    file                        As String
    datastart                   As Long
    DataLen                     As Long
    wfx                         As WAVEFORMATEX
End Type

Private Type WaveFileHeaderA
    riff                        As String * 4         ' RIFF Header
    len                         As Long               ' file length
    WavFmt                      As String * 8         ' Wave Format
    HdrLen                      As Long               ' header length
    format                      As Integer            ' format
    NumChannels                 As Integer            ' channels
    SampleRate                  As Long               ' frequency
    BytesPerSec                 As Long               ' Bytes/second
    BlockAlign                  As Integer            ' Block Align
    BitsPerSample               As Integer            ' Bits/Sample
    Data                        As String * 4         ' Data Chunk
    DataLen                     As Long               ' datalength w/o header
End Type

Private Const MMIO_READ                     As Long = &H0&
Private Const MMIO_FINDCHUNK                As Long = &H10&
Private Const MMIO_FINDRIFF                 As Long = &H20&
Private Const MMSYSERR_NOERROR              As Long = &H0&
Private Const SEEK_CUR                      As Long = &H1&

Private Const ACM_METRIC_MAX_SIZE_FORMAT    As Long = &H32

Private Const ACM_STREAMCONVERTF_BLOCKALIGN As Long = &H4
Private Const ACM_STREAMCONVERTF_END        As Long = &H20

Private Const ACM_STREAMSIZEF_SOURCE        As Long = &H0&

Private Enum ACMDRIVERDETAILS_SUPPORTF
    ACMDRIVERDETAILS_SUPPORTF_CODEC = &H1
    ACMDRIVERDETAILS_SUPPORTF_CONVERTER = &H2&
    ACMDRIVERDETAILS_SUPPORTF_FILTER = &H4&
    ACMDRIVERDETAILS_SUPPORTF_HARDWARE = &H8&
    ACMDRIVERDETAILS_SUPPORTF_ASYNC = &H10&
    ACMDRIVERDETAILS_SUPPORTF_DISABLED = &H80000000
End Enum

Private Enum ACM_FORMATENUMF
    ACM_FORMATENUMF_WFORMATTAG = &H10000
    ACM_FORMATENUMF_NCHANNELS = &H20000
    ACM_FORMATENUMF_NSAMPLESPERSEC = &H40000
    ACM_FORMATENUMF_WBITSPERSAMPLE = &H80000
    ACM_FORMATENUMF_CONVERT = &H100000
    ACM_FORMATENUMF_SUGGEST = &H200000
    ACM_FORMATENUMF_HARDWARE = &H400000
    ACM_FORMATENUMF_INPUT = &H800000
    ACM_FORMATENUMF_OUTPUT = &H1000000
End Enum

Private WaveFiles()             As WaveFile

Private wfxOutput               As WAVEFORMATEX

Public Event Progress(percent As Integer)
Public Event FileChanged(file As String)

Public Property Get OutputFormat() As Long
    OutputFormat = VarPtr(wfxOutput)
End Property

Public Property Let OutputFormat(wfx As Long)
    CopyMem wfxOutput, ByVal wfx, Len(wfxOutput)
End Property

Public Sub ResetQueue()
    Class_Initialize
End Sub

Public Function WaveAdd(file As String) As Boolean

    '   File Handle
    Dim fh                  As Integer
    '   MMIO Handle
    Dim hMMIO               As Long
    '   catch returned values
    Dim ret                 As Long

    '   Wave File Header
    Dim udtWavHdr           As WaveFileHeader
    '   Collection Item
    Dim udtWavFile          As WaveFile


    '   MMIO info structures
    Dim mmioinf             As mmioinfo
    Dim mmckinfoParentIn    As MMCKINFO
    Dim mmckinfoSubchunkIn  As MMCKINFO

    ' file exists?
    If Not FileExists(file) Then Exit Function

    fh = FreeFile
    Open file For Binary Access Read As #fh

        ' read Wave Header
        Get #fh, , udtWavHdr

    Close #fh

    udtWavFile.file = file
    udtWavFile.wfx = udtWavHdr.WaveFormat

    ' get the data chunk`s length
    hMMIO = mmioOpen(file, mmioinf, MMIO_READ)
    If (hMMIO = 0) Then
        Exit Function
    End If

    ' try to find a WAVE chunk to make sure
    ' it is a wave file
    mmckinfoParentIn.fccType = mmioStringToFOURCC("WAVE", 0)
    ret = mmioDescendParent(hMMIO, mmckinfoParentIn, 0, MMIO_FINDRIFF)
    If Not (ret = MMSYSERR_NOERROR) Then
        Exit Function
    End If

    ' try to find the DATA chunk
    mmckinfoSubchunkIn.ckid = mmioStringToFOURCC("data", 0)
    ret = mmioDescend(hMMIO, mmckinfoSubchunkIn, mmckinfoParentIn, MMIO_FINDCHUNK)
    If Not (ret = MMSYSERR_NOERROR) Then
        Exit Function
    End If

    ' start offset of the DATA chunk
    udtWavFile.datastart = mmioSeek(hMMIO, 0, SEEK_CUR) + 1
    ' length of the DATA chunk
    udtWavFile.DataLen = mmckinfoSubchunkIn.ckSize

    mmioClose hMMIO, 0

    ' add wave file to queue
    WaveFiles(UBound(WaveFiles)) = udtWavFile
    ReDim Preserve WaveFiles(UBound(WaveFiles) + 1)

    WaveAdd = True

End Function

Public Function WaveConcatenate(output As String, _
    Optional WriteWaveHeader As Boolean = True) As Boolean

    ' file handles
    Dim fhIN            As Integer
    Dim fhOUT           As Integer

    ' conversion buffer
    Dim btBuffer()      As Byte
    Dim lngBufferSize   As Long
    ' target buffer
    Dim btTarget()      As Byte
    Dim lngTargetSize   As Long
    ' out buffer
    Dim btCompressed()  As Byte
    Dim lngBytesOut     As Long

    ' counter
    Dim i               As Long

    ' ACM driver handle
    Dim hDriver         As Long
    ' Stream handle
    Dim hStream         As Long
    ' stream header
    Dim udtStreamheader As ACMSTREAMHEADER

    ' catch returned values
    Dim ret             As Long

    ' current WAVEFORMATEX
    Dim curWFX          As WAVEFORMATEX

    ' output wave header
    Dim whdr            As WaveFileHeaderA

    ' estimated output size
    Dim lngOutputSize   As Long

    ' files added?
    If WaveFiles(0).file = "" Then
        Exit Function
    End If

    fhIN = FreeFile
    fhOUT = FreeFile(fhIN)

    Open output For Binary Access Write As #fhOUT
    If WriteWaveHeader Then Seek #fhOUT, 45

    For i = LBound(WaveFiles) To UBound(WaveFiles)
        lngOutputSize = lngOutputSize + WaveFiles(i).DataLen
    Next

    For i = LBound(WaveFiles) To UBound(WaveFiles) - 1

        RaiseEvent FileChanged(WaveFiles(i).file)

        curWFX = WaveFiles(i).wfx

        ' open the stream
        ret = acmStreamOpen(hStream, 0, curWFX, wfxOutput, 0, 0, 0, 0)
        If Not (ret = MMSYSERR_NOERROR) Then
            Exit Function
        End If

        ' input buffer size (10 secs)
        lngBufferSize = curWFX.nBlockAlign _
            * curWFX.nSamplesPerSec * 10

        ReDim btBuffer(lngBufferSize - 1) As Byte

        ' get the output buffer size
        ret = acmStreamSize(hStream, lngBufferSize, lngTargetSize, ACM_STREAMSIZEF_SOURCE)
        If Not (ret = MMSYSERR_NOERROR) Then
            acmStreamClose hStream, 0
            Exit Function
        End If
        ReDim btTarget(lngTargetSize - 1) As Byte

        ' stream header data
        With udtStreamheader
            .cbStruct = LenB(udtStreamheader)
            .pbSrc = VarPtr(btBuffer(0))
            .cbSrcLength = lngBufferSize
            .pbDst = VarPtr(btTarget(0))
            .cbDstLength = lngTargetSize
        End With

        ' prepare a stream header
        ret = acmStreamPrepareHeader(hStream, udtStreamheader, 0)
        If Not (ret = MMSYSERR_NOERROR) Then
            acmStreamClose hStream, 0
            Exit Function
        End If

        Open WaveFiles(i).file For Binary Access Read As #fhIN

        Seek #fhIN, WaveFiles(i).datastart

        Do

            If Seek(fhIN) + lngBufferSize > WaveFiles(i).DataLen + WaveFiles(i).datastart Then

                udtStreamheader.cbSrcLength = WaveFiles(i).DataLen + WaveFiles(i).datastart - Seek(fhIN)
                ' read 10 secs of wave data
                Get #fhIN, , btBuffer

                ' convert the data
                acmStreamConvert hStream, udtStreamheader, ACM_STREAMCONVERTF_END
                ' amount of converted data
                lngBytesOut = udtStreamheader.cbDstLengthUsed
    
                If lngBytesOut > 0 Then
                    ReDim btCompressed(lngBytesOut - 1) As Byte
                    CopyMem btCompressed(0), btTarget(0), lngBytesOut
                    Put #fhOUT, , btCompressed
                End If

                Exit Do

            Else

                udtStreamheader.cbSrcLength = lngBufferSize
                ' read 10 secs of wave data
                Get #fhIN, , btBuffer

                ' convert the data
                acmStreamConvert hStream, udtStreamheader, ACM_STREAMCONVERTF_BLOCKALIGN
                ' amount of converted data
                lngBytesOut = udtStreamheader.cbDstLengthUsed
    
                If lngBytesOut > 0 Then
                    ReDim btCompressed(lngBytesOut - 1) As Byte
                    CopyMem btCompressed(0), btTarget(0), lngBytesOut
                    Put #fhOUT, , btCompressed
                End If

            End If

            RaiseEvent Progress(Seek(fhOUT) / lngOutputSize * 100)

            DoEvents

        Loop

        Close #fhIN

        ' reset the ACM stuff
        acmStreamReset hStream, 0
        acmStreamUnprepareHeader hStream, udtStreamheader, 0
        acmStreamClose hStream, 0

    Next

    ' write the wave header
    If WriteWaveHeader Then

        Seek #fhOUT, 1
        With wfxOutput
    
            WriteWavHdr fhOUT, LOF(fhOUT) - 45, _
                        .nSamplesPerSec, _
                        .nSamplesPerSec * 4, _
                        .nChannels, _
                        .wBitsPerSample, _
                        .nBlockAlign
    
        End With

    End If

    Close #fhOUT

    WaveConcatenate = True

End Function

Private Sub Class_Initialize()
    ReDim WaveFiles(0) As WaveFile
End Sub

Private Function FileExists(strPath As String) As Boolean
    On Error Resume Next
    FileExists = ((GetAttr(strPath) And (vbDirectory Or vbVolume)) = 0)
End Function

Private Sub WriteWavHdr(ByVal FileHandle As Integer, ByVal F_Len As Long, _
                        ByVal SampleRate As Long, ByVal BytesPerSec As Long, _
                        ByVal Channels As Long, ByVal BitsPerSample As Long, _
                        ByVal BlockAlign As Long)

    Dim wavhdr As WaveFileHeaderA

    With wavhdr
        .WavFmt = "WAVEfmt "            ' WAVEfmt
        .Data = "data"                  ' data chunk
        .riff = "RIFF"                  ' RIFF chunk

        .len = F_Len + 44 - 8           ' length of the file
        .DataLen = F_Len                ' data length without header length
        .HdrLen = 16                    ' header length

        .SampleRate = SampleRate        ' sample rate (44.1kHz)
        .BytesPerSec = BytesPerSec      ' Bytes Per Second
        .BitsPerSample = BitsPerSample  ' Bits per sample
        .NumChannels = Channels         ' number of channels (stereo)
        .BlockAlign = BlockAlign        ' Block Align

        .format = 1                     ' format (PCM WAV)
    End With

    Put #FileHandle, , wavhdr

End Sub
