Imports System.IO
Imports System.Threading
Imports Microsoft.DirectX
Imports Microsoft.DirectX.DirectSound

Public Class SoundRecord
    ''' <summary>
    ''' 为了正常使用 DirectX 录制音频，需要做到以下：
    ''' 1.App.config 文件里"startup"标签内后加" useLegacyV2RuntimeActivationPolicy="true""
    ''' 2.并把 [\项目名称\.exe.config] 文件与 EXE 可执行程序在同一目录
    ''' 如果出现 [LoaderLock] 错误：按下 "Ctrl+Alt+E"(调试\窗口\异常设置) 打开 异常设置，找到 "Managed Debuggin Assistants-" 取消勾选"LoaderLock"。
    ''' 或者直接 "继续"
    ''' </summary>

#Region "成员数据"
    Private CaptureDevice As Capture = Nothing ' 音频捕捉设备
    Private CaptureBuffer As CaptureBuffer = Nothing     ' 缓冲区对象
    Private WaveFormat As WaveFormat             ' 录音的格式
    Private NextCaptureOffset As Integer = 0         ' 该次录音缓冲区的起始点
    Private SampleCount As Integer = 0               ' 录制的样本数目
    Private NotifyObject As Notify = Nothing               ' 消息通知对象
    Public Const NotifyCount As Integer = 16           ' 通知的个数
    Private NotifySize As Integer = 0                ' 每次通知大小
    Private BufferSize As Integer = 0                ' 缓冲队列大小
    Private NotifyThread As Thread = Nothing                 ' 处理缓冲区消息的线程
    Private NotificationEvent As AutoResetEvent = Nothing    ' 通知事件
    Private FileName As String = String.Empty     ' 文件保存路径
    Private WaveFile As FileStream = Nothing         ' 文件流
    Private Writer As BinaryWriter = Nothing         ' 写文件
#End Region

#Region "对外操作函数"
    ''' <summary>
    ''' 构造函数,设定录音设备,设定录音格式.
    ''' </summary>
    Public Sub New()
        InitCaptureDevice()        ' 初始化音频捕捉设备
        WaveFormat = CreateWaveFormat() ' 设定录音格式
    End Sub

    ''' <summary>
    ''' 创建录音格式,此处使用16bit,16KHz,Mono的录音格式
    ''' </summary>
    Private Function CreateWaveFormat() As WaveFormat
        Dim format As WaveFormat = New WaveFormat()
        format.FormatTag = WaveFormatTag.Pcm   ' PCM
        format.SamplesPerSecond = 16000        ' 采样率：16KHz，人耳可以识别最高 48KHz
        format.BitsPerSample = 16              ' 采样位数：16Bit
        format.Channels = 1                    ' 声道：Mono
        format.BlockAlign = Convert.ToInt16(format.Channels * (format.BitsPerSample / 8))  ' 单位采样点的字节数 
        format.AverageBytesPerSecond = format.BlockAlign * format.SamplesPerSecond
        Return format
        ' 按照以上采样规格，可知采样1秒钟的字节数为 16000*2=32000B 约为31K
    End Function

    ''' <summary>
    ''' 设定录音结束后保存的文件,包括路径
    ''' </summary>
    ''' <param name="FileName">保存wav文件的路径名</param>
    Public Sub SetFileName(ByVal FileName As String)
        Me.FileName = FileName
    End Sub

    ''' <summary>
    ''' 开始录音
    ''' </summary>
    Public Sub RecordStart()
        ' 创建录音文件
        CreateSoundFile()
        ' 创建一个录音缓冲区，并开始录音
        CreateCaptureBuffer()
        ' 建立通知消息,当缓冲区满的时候处理方法
        InitNotifications()
        CaptureBuffer.Start(True)
    End Sub

    ''' <summary>
    ''' 停止录音
    ''' </summary>
    Public Sub RecordStop()
        CaptureBuffer.Stop()      ' 调用缓冲区的停止方法，停止采集声音
        If (NotificationEvent IsNot Nothing) Then NotificationEvent.Set()       '关闭通知
        NotifyThread.Abort()  '结束线程
        RecordCapturedData()   ' 将缓冲区最后一部分数据写入到文件中
        ' 写WAV文件尾
        Writer.Seek(4, SeekOrigin.Begin)
        Writer.Write(SampleCount + 36)   ' 写文件长度
        Writer.Seek(40, SeekOrigin.Begin)
        Writer.Write(SampleCount)                ' 写数据长度
        Writer.Close()
        WaveFile.Close()
        Writer = Nothing
        WaveFile = Nothing
    End Sub
#End Region

#Region "对内操作函数"
    ''' <summary>
    ''' 初始化录音设备,此处使用主录音设备.
    ''' </summary>
    ''' <returns>调用成功返回true,否则返回false</returns>
    Private Function InitCaptureDevice() As Boolean
        Dim Devices As CaptureDevicesCollection = New CaptureDevicesCollection()  ' 枚举音频捕捉设备
        Dim DeviceGuid As Guid = Guid.Empty
        '使用默认音频捕捉设备
        If (Devices.Count > 0) Then
            DeviceGuid = Devices(0).DriverGuid
        Else
            Debug.Print("系统中没有音频捕捉设备")
            Return False
        End If
        ' 用指定的捕捉设备创建Capture对象
        Try
            CaptureDevice = New Capture(DeviceGuid)
        Catch e As DirectXException
            Debug.Print("错误：" & e.ToString())
            Return False
        End Try
        Return True
    End Function

    ''' <summary>
    ''' 创建录音使用的缓冲区
    ''' </summary>
    Private Sub CreateCaptureBuffer()
        ' 缓冲区的描述对象
        Dim BufferDescription As CaptureBufferDescription = New CaptureBufferDescription()
        If (NotifyObject IsNot Nothing) Then
            NotifyObject.Dispose()
            NotifyObject = Nothing
        End If
        If (CaptureBuffer IsNot Nothing) Then
            CaptureBuffer.Dispose()
            CaptureBuffer = Nothing
        End If
        '设定通知的大小,默认为1s
        NotifySize = IIf((1024 > WaveFormat.AverageBytesPerSecond / 8), 1024, (WaveFormat.AverageBytesPerSecond / 8))
        NotifySize -= NotifySize Mod WaveFormat.BlockAlign
        '设定缓冲区大小
        BufferSize = NotifySize * NotifyCount
        '创建缓冲区描述
        BufferDescription.BufferBytes = BufferSize
        BufferDescription.Format = WaveFormat           ' 录音格式
        ' 创建缓冲区
        CaptureBuffer = New CaptureBuffer(BufferDescription, CaptureDevice)
        NextCaptureOffset = 0
    End Sub

    ''' <summary>
    ''' 初始化通知事件,将原缓冲区分成16个缓冲队列,在每个缓冲队列的结束点设定通知点.
    ''' </summary>
    ''' <returns>是否成功</returns>
    Private Function InitNotifications() As Boolean
        If (CaptureBuffer Is Nothing) Then
            Debug.Print("未创建录音缓冲区")
            Return False
        End If
        ' 创建一个通知事件,当缓冲队列满了就激发该事件.
        NotificationEvent = New AutoResetEvent(False)
        ' 创建一个线程管理缓冲区事件
        If (NotifyThread Is Nothing) Then
            NotifyThread = New Thread(New ThreadStart(AddressOf WaitThread))
            NotifyThread.Start()
        End If
        ' 设定通知的位置
        Dim PositionNotify(NotifyCount) As BufferPositionNotify
        For Index As Integer = 0 To NotifyCount - 1
            PositionNotify(Index) = New BufferPositionNotify
            PositionNotify(Index).Offset = (NotifySize * Index) + NotifySize - 1
            PositionNotify(Index).EventNotifyHandle = NotificationEvent.SafeWaitHandle.DangerousGetHandle()
        Next
        NotifyObject = New Notify(CaptureBuffer)
        NotifyObject.SetNotificationPositions(PositionNotify, NotifyCount)
        Return True
    End Function

    ''' <summary>
    ''' 接收缓冲区满消息的处理线程
    ''' </summary>
    Private Sub WaitThread()
        Do While True
            ' 等待缓冲区的通知消息
            NotificationEvent.WaitOne(Timeout.Infinite, True)
            ' 录制数据
            RecordCapturedData()
        Loop
    End Sub

    ''' <summary>
    ''' 将录制的数据写入wav文件
    ''' </summary>
    Private Sub RecordCapturedData()
        Dim CaptureData As Byte() = Nothing
        Dim ReadPosition As Integer = 0, CapturePos As Integer = 0, LockSize As Integer = 0
        CaptureBuffer.GetCurrentPosition(CapturePos, ReadPosition) '此处两个参数需要传入地址
        LockSize = ReadPosition - NextCaptureOffset
        ' 因为是循环的使用缓冲区，所以有一种情况下为负：当文以载读指针回到第一个通知点，而Ibuffeoffset还在最后一个通知处
        If (LockSize < 0) Then LockSize += BufferSize
        LockSize -= (LockSize Mod NotifySize)   ' 对齐缓冲区边界,实际上由于开始设定完整,这个操作是多余的.
        If (LockSize = 0) Then Exit Sub

        ' 读取缓冲区内的数据
        CaptureData = CaptureBuffer.Read(NextCaptureOffset, GetType(Byte), LockFlag.None, LockSize)
        ' 写入Wav文件
        Writer.Write(CaptureData, 0, CaptureData.Length)
        ' 更新已经录制的数据长度.
        SampleCount += CaptureData.Length
        ' 移动录制数据的起始点,通知消息只负责指示产生消息的位置,并不记录上次录制的位置
        NextCaptureOffset += CaptureData.Length
        NextCaptureOffset = NextCaptureOffset Mod BufferSize ' Circular buffer
    End Sub

    ''' <summary>
    ''' 创建保存的波形文件,并写入必要的文件头.
    ''' </summary>
    Private Sub CreateSoundFile()
        ' Open up the wave file for writing.
        WaveFile = New FileStream(FileName, FileMode.Create)
        Writer = New BinaryWriter(WaveFile)
        '************************************************************************** 
        'Here Is where the file will be created. A 
        'wave File Is a RIFF file, which has chunks 
        'of data that describe what the file contains. 
        'A wave RIFF file Is put together Like this
        'The 12 byte RIFF chunk Is constructed Like this 
        'Bytes 0 - 3   'R' 'I' 'F' 'F' 
        'Bytes 4 - 7   Length Of file, minus the first 8 bytes of the RIFF description. 
        '(4 bytes for "WAVE" + 24 bytes for format chunk length + 
        '8 bytes for data chunk description + actual sample data size.) 
        'Bytes 8 - 11 'W' 'A' 'V' 'E' 
        'The 24 byte FORMAT chunk Is constructed Like this 
        'Bytes 0 - 3  'f' 'm' 't' ' ' 
        'Bytes 4 - 7 : The Format chunk length. This Is always 16. 
        'Bytes 8 - 9 : File Padding.Always 1. 
        'Bytes 10 - 11: Number of channels. Either 1 for mono,  Or 2 for stereo. 
        'Bytes 12 - 15: Sample Rate.
        'Bytes 16- 19: Number of bytes per second. 
        'Bytes 20 - 21: Bytes per sample. 1 For 8 bit mono, 2 For 8 bit stereo Or 
        '16 bit mono, 4 for 16 bit stereo. 
        'Bytes 22 - 23: Number of bits per sample. 
        'The Data chunk Is constructed Like this
        'Bytes 0 - 3  'd' 'a' 't' 'a' 
        'Bytes 4 - 7  Length Of data, in bytes. 
        'Bytes 8 -: Actual sample data. 
        '***************************************************************************
        ' Set up file with RIFF chunk info.
        Dim ChunkRIFF As Char() = {"R", "I", "F", "F"}
        Dim ChunkTYPE As Char() = {"W", "A", "V", "E"}
        Dim ChunkFMT As Char() = {"f", "m", "t", " "}
        Dim ChunkDATA As Char() = {"d", "a", "t", "a"}

        Dim FilePadding As Short = 1                ' File padding
        Dim FormatChunkLength As Integer = &H10  ' Format chunk length.
        Dim FileLength As Integer = 0                ' File length, minus first 8 bytes Of RIFF description. This will be filled In later.
        Dim BytesPerSample As Short = 0     ' Bytes per sample.

        ' 一个样本点的字节数目
        If (WaveFormat.BitsPerSample = 8 And WaveFormat.Channels = 1) Then
            BytesPerSample = 1
        ElseIf ((WaveFormat.BitsPerSample = 8 And WaveFormat.Channels = 2) Or (WaveFormat.BitsPerSample = 16 And WaveFormat.Channels = 1)) Then
            BytesPerSample = 2
        ElseIf (WaveFormat.BitsPerSample = 16 And WaveFormat.Channels = 2) Then
            BytesPerSample = 4
        End If
        ' RIFF 块
        Writer.Write(ChunkRIFF)
        Writer.Write(FileLength)
        Writer.Write(ChunkTYPE)

        ' WAVE块
        Writer.Write(ChunkFMT)
        Writer.Write(FormatChunkLength)
        Writer.Write(FilePadding)
        Writer.Write(WaveFormat.Channels)
        Writer.Write(WaveFormat.SamplesPerSecond)
        Writer.Write(WaveFormat.AverageBytesPerSecond)
        Writer.Write(BytesPerSample)
        Writer.Write(WaveFormat.BitsPerSample)

        ' 数据块
        Writer.Write(ChunkData)
        Writer.Write(0)   ' The sample length will be written in later.
    End Sub
#End Region

End Class
