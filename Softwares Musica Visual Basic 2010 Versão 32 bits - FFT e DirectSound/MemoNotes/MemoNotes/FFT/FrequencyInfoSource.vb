Imports System.Collections.Generic
Imports System.Linq
Imports System.Threading

Public MustInherit Class FrequencyInfoSource
	Public MustOverride Sub Listen()
	Public MustOverride Sub [Stop]()

	Public Event FrequencyDetected As EventHandler(Of FrequencyDetectedEventArgs)

	Protected Sub OnFrequencyDetected(e As FrequencyDetectedEventArgs)
		RaiseEvent FrequencyDetected(Me, e)
	End Sub
End Class

Public Class FrequencyDetectedEventArgs
	Inherits EventArgs
	Private m_frequency As Double

	Public ReadOnly Property Frequency() As Double
		Get
			Return m_frequency
		End Get
	End Property

	Public Sub New(frequency As Double)
		Me.m_frequency = frequency
	End Sub
End Class
