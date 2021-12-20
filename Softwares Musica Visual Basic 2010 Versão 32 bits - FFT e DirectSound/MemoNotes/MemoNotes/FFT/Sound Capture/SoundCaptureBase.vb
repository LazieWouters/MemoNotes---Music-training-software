Imports System.Threading
Imports Microsoft.DirectX.DirectSound
Imports Microsoft.Win32.SafeHandles


''' <summary>
''' Base class to capture audio samples.
''' </summary>
Public MustInherit Class SoundCaptureBase
	Implements IDisposable
    Dim BufferSeconds As Integer = My.Settings.NovoValorBufferSeconds '3
    Dim NotifyPointsInSecond As Integer = My.Settings.NovoValorNotifyPointsInSecond ' indica quantas atualizações por segundo. 15 é o máximo que consegui colocar sem dar erro.


    'It depends on your hardware: smaller buffer gives less latency but more CPU usage.
    'I recommend to experiment with NotifyPointsInSecond, increasing this amount will give you smaller amount of data that need to processed by the FFT.



	' change in next two will require also code change
    Const BitsPerSample As Integer = 16 '16
    Private ChannelCount As Integer = My.Settings.NovoValorLineInput(2) '1

    Dim m_sampleRate As Integer = My.Settings.NovoValorSampleRate '44100
	Private m_isCapturing As Boolean = False
	Private disposed As Boolean = False

	Public ReadOnly Property IsCapturing() As Boolean
		Get
			Return m_isCapturing
		End Get
	End Property

	Public Property SampleRate() As Integer
		Get
			Return m_sampleRate
		End Get
		Set
			If m_sampleRate <= 0 Then
				Throw New ArgumentOutOfRangeException()
			End If

			EnsureIdle()

			m_sampleRate = value
		End Set
	End Property

	Private capture As Capture
	Private buffer As CaptureBuffer
	Private notify As Notify
	Private bufferLength As Integer
	Private positionEvent As AutoResetEvent
	Private positionEventHandle As SafeWaitHandle
	Private terminated As ManualResetEvent
	Private thread As Thread
	Private device As SoundCaptureDevice

	Public Sub New()

		Me.New(SoundCaptureDevice.GetDefaultDevice())
	End Sub

	Public Sub New(device As SoundCaptureDevice)
		Me.device = device

		positionEvent = New AutoResetEvent(False)
		positionEventHandle = positionEvent.SafeWaitHandle
		terminated = New ManualResetEvent(True)
	End Sub

	Private Sub EnsureIdle()
		If IsCapturing Then
			Throw New SoundCaptureException("Capture is in process")
		End If
	End Sub

	''' <summary>
	''' Starts capture process.
	''' </summary>
    Public Sub Start()

        Try

            EnsureIdle()

            m_isCapturing = True

            Dim format As New WaveFormat()
            format.Channels = ChannelCount
            format.BitsPerSample = BitsPerSample
            format.SamplesPerSecond = SampleRate
            format.FormatTag = WaveFormatTag.Pcm
            format.BlockAlign = CShort((format.Channels * format.BitsPerSample + 7) \ 8)
            format.AverageBytesPerSecond = format.BlockAlign * format.SamplesPerSecond

            bufferLength = format.AverageBytesPerSecond * BufferSeconds
            Dim desciption As New CaptureBufferDescription()
            desciption.Format = format
            desciption.BufferBytes = bufferLength

            capture = New Capture(device.Id)
            buffer = New CaptureBuffer(desciption, capture)

            Dim waitHandleCount As Integer = BufferSeconds * NotifyPointsInSecond
            Dim positions As BufferPositionNotify() = New BufferPositionNotify(waitHandleCount - 1) {}
            For i As Integer = 0 To waitHandleCount - 1
                Dim position As New BufferPositionNotify()
                position.Offset = (i + 1) * bufferLength \ positions.Length - 1
                position.EventNotifyHandle = positionEventHandle.DangerousGetHandle()
                positions(i) = position
            Next

            notify = New Notify(buffer)
            notify.SetNotificationPositions(positions)

            terminated.Reset()
            thread = New Thread(New ThreadStart(AddressOf ThreadLoop))
            thread.Name = "Sound capture"
            thread.Start()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

	Private Sub ThreadLoop()
		buffer.Start(True)
		Try
			Dim nextCapturePosition As Integer = 0
			Dim [handles] As WaitHandle() = New WaitHandle() {terminated, positionEvent}
			While WaitHandle.WaitAny([handles]) > 0
				Dim capturePosition As Integer, readPosition As Integer
				buffer.GetCurrentPosition(capturePosition, readPosition)

				Dim lockSize As Integer = readPosition - nextCapturePosition
				If lockSize < 0 Then
					lockSize += bufferLength
				End If
				If (lockSize And 1) <> 0 Then
					lockSize -= 1
				End If

				Dim itemsCount As Integer = lockSize >> 1

				Dim data As Short() = DirectCast(buffer.Read(nextCapturePosition, GetType(Short), LockFlag.None, itemsCount), Short())
				ProcessData(data)
				nextCapturePosition = (nextCapturePosition + lockSize) Mod bufferLength
            End While

        Catch ex As Exception
            Exit Sub
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui


		Finally
			buffer.[Stop]()
		End Try
	End Sub

	''' <summary>
	''' Processes the captured data.
	''' </summary>
	''' <param name="data">Captured data</param>
	Protected MustOverride Sub ProcessData(data As Short())

	''' <summary>
	''' Stops capture process.
	''' </summary>
    Public Sub [Stop]()

        Try

            If m_isCapturing Then
                m_isCapturing = False

                terminated.[Set]()
                thread.Join()

                notify.Dispose()
                buffer.Dispose()
                capture.Dispose()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

	Private Sub IDisposable_Dispose() Implements IDisposable.Dispose
		Dispose(True)
	End Sub

	Protected Overrides Sub Finalize()
		Try
			Dispose(False)
		Finally
			MyBase.Finalize()
		End Try
	End Sub

    Private Sub Dispose(ByVal disposing As Boolean)

        Try

            If disposed Then
                Return
            End If

            disposed = True
            GC.SuppressFinalize(Me)
            If IsCapturing Then
                [Stop]()
            End If
            positionEventHandle.Dispose()
            positionEvent.Close()
            terminated.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub
End Class
