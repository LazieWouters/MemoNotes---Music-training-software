Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports SoundCapture
Imports SoundAnalysis

Class SoundFrequencyInfoSource
	Inherits FrequencyInfoSource
	Private device As SoundCaptureDevice
	Private adapter As _Adapter

	Friend Sub New(device As SoundCaptureDevice)
		Me.device = device
	End Sub

	Public Overrides Sub Listen()
		adapter = New _Adapter(Me, device)
		adapter.Start()
	End Sub

	Public Overrides Sub [Stop]()
		adapter.[Stop]()
	End Sub

	Private Class _Adapter
		Inherits SoundCaptureBase
		Private owner As SoundFrequencyInfoSource

        Const MinFreq As Double = 40 '60
        Private MaxFreq As Double = My.Settings.NovoValorMaxFreq '3951.1 'última nota Si do teclado, acima deste valor está dando erro index out of bounds                  '1300

		Friend Sub New(owner As SoundFrequencyInfoSource, device As SoundCaptureDevice)
			MyBase.New(device)
			Me.owner = owner
		End Sub

        Protected Overrides Sub ProcessData(ByVal data As Short())

            Dim x As Double() = New Double(data.Length - 1) {}

            Try

                For i As Integer = 0 To x.Length - 1
                    x(i) = data(i) / 32768.0
                Next

                Dim freq As Double = FrequencyUtils.FindFundamentalFrequency(x, SampleRate, MinFreq, MaxFreq)
                owner.OnFrequencyDetected(New FrequencyDetectedEventArgs(freq))

            Catch ex As Exception


                'x(1) = 5000 'por algum motivo isto está evitando aquela mensagem chata "index was out of range"


                'MsgBox(ex.Message, MsgBoxStyle.Critical)
                'Exit Sub
                'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
            Finally
                'esta parte do bloco é executada independentemente se algum erro acontecer ou não

            End Try

        End Sub
	End Class
End Class
