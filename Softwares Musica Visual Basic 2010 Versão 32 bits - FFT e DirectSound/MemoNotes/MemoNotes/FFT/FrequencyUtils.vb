Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.Text

Imports SoundAnalysis

''' <summary>
''' Utils that helps to detect the fundumental frequency.
''' </summary> 
NotInheritable Class FrequencyUtils
	Private Sub New()
	End Sub
	''' <summary>
	''' Finds fundamental frequency: calculates spectrogram, finds peaks, analyzes
	''' and refines frequency by diff sample values.
	''' </summary>
	''' <param name="x">The sounds samples data</param>
	''' <param name="sampleRate">The sound sample rate</param>
	''' <param name="minFreq">The min useful frequency</param>
	''' <param name="maxFreq">The max useful frequency</param>
	''' <returns>Found frequency, 0 - otherwise</returns>
    Friend Shared Function FindFundamentalFrequency(ByVal x As Double(), ByVal sampleRate As Integer, ByVal minFreq As Double, ByVal maxFreq As Double) As Double
        Try
            Dim spectr As Double() = FftAlgorithm.Calculate(x)

            Dim usefullMinSpectr As Integer = Math.Max(0, CInt(Math.Truncate(minFreq * spectr.Length / sampleRate)))
            Dim usefullMaxSpectr As Integer = Math.Min(spectr.Length, CInt(Math.Truncate(maxFreq * spectr.Length / sampleRate)) + 1)

            ' find peaks in the FFT frequency bins 
            Const PeaksCount As Integer = 5
            Dim peakIndices As Integer()
            peakIndices = FindPeaks(spectr, usefullMinSpectr, usefullMaxSpectr - usefullMinSpectr, PeaksCount)

            If Array.IndexOf(peakIndices, usefullMinSpectr) >= 0 Then
                ' lowest usefull frequency bin shows active
                ' looks like is no detectable sound, return 0
                Return 0
            End If

            ' select fragment to check peak values: data offset
            Const verifyFragmentOffset As Integer = 0
            ' ... and half length of data
            Dim verifyFragmentLength As Integer = CInt(Math.Truncate(sampleRate / minFreq))

            ' trying all peaks to find one with smaller difference value
            Dim minPeakValue As Double = [Double].PositiveInfinity
            Dim minPeakIndex As Integer = 0
            Dim minOptimalInterval As Integer = 0
            For i As Integer = 0 To peakIndices.Length - 1
                Dim index As Integer = peakIndices(i)
                Dim binIntervalStart As Integer = spectr.Length \ (index + 1), binIntervalEnd As Integer = spectr.Length \ index
                Dim interval As Integer
                Dim peakValue As Double
                ' scan bins frequencies/intervals
                ScanSignalIntervals(x, verifyFragmentOffset, verifyFragmentLength, binIntervalStart, binIntervalEnd, _
                    interval, peakValue)

                If peakValue < minPeakValue Then
                    minPeakValue = peakValue
                    minPeakIndex = index
                    minOptimalInterval = interval
                End If
            Next

            Return CDbl(sampleRate) / minOptimalInterval
        Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxStyle.Critical)
            Exit Function
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try


    End Function

    Private Shared Sub ScanSignalIntervals(ByVal x As Double(), ByVal index As Integer, ByVal length As Integer, _
     ByVal intervalMin As Integer, ByVal intervalMax As Integer, _
     <Out()> ByRef optimalInterval As Integer, <Out()> ByRef optimalValue As Double)

        Try

            optimalValue = [Double].PositiveInfinity
            optimalInterval = 0

            ' distance between min and max range value can be big
            ' limiting it to the fixed value
            Const MaxAmountOfSteps As Integer = 30
            Dim steps As Integer = intervalMax - intervalMin
            If steps > MaxAmountOfSteps Then
                steps = MaxAmountOfSteps
            ElseIf steps <= 0 Then
                steps = 1
            End If

            ' trying all intervals in the range to find one with
            ' smaller difference in signal waves
            For i As Integer = 0 To steps - 1
                Dim interval As Integer = intervalMin + (intervalMax - intervalMin) * i \ steps

                Dim sum As Double = 0
                For j As Integer = 0 To length - 1
                    Dim diff As Double = x(index + j) - x(index + j + interval)
                    sum += diff * diff
                Next
                If optimalValue > sum Then
                    optimalValue = sum
                    optimalInterval = interval
                End If
            Next

        Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxStyle.Critical)
            Exit Sub
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
       
    End Sub

	Private Shared Function FindPeaks(values As Double(), index As Integer, length As Integer, peaksCount As Integer) As Integer()
        Dim peakValues As Double() = New Double(peaksCount - 1) {}
        Dim peakIndices As Integer() = New Integer(peaksCount - 1) {}

        Try



            For i As Integer = 0 To peaksCount - 1
                peakValues(i) = values(InlineAssignHelper(peakIndices(i), i + index))
            Next

            ' find min peaked value
            Dim minStoredPeak As Double = peakValues(0)
            Dim minIndex As Integer = 0
            For i As Integer = 1 To peaksCount - 1
                If minStoredPeak > peakValues(i) Then
                    minStoredPeak = peakValues(InlineAssignHelper(minIndex, i))
                End If
            Next

            For i As Integer = peaksCount To length - 1
                If minStoredPeak < values(i + index) Then
                    ' replace the min peaked value with bigger one
                    peakValues(minIndex) = values(InlineAssignHelper(peakIndices(minIndex), i + index))

                    ' and find min peaked value again
                    minStoredPeak = peakValues(InlineAssignHelper(minIndex, 0))
                    For j As Integer = 1 To peaksCount - 1
                        If minStoredPeak > peakValues(j) Then
                            minStoredPeak = peakValues(InlineAssignHelper(minIndex, j))
                        End If
                    Next
                End If
            Next

            Return peakIndices

        Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxStyle.Critical)
            Exit Function
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Function
	Private Shared Function InlineAssignHelper(Of T)(ByRef target As T, value As T) As T
		target = value
		Return value
	End Function
End Class
