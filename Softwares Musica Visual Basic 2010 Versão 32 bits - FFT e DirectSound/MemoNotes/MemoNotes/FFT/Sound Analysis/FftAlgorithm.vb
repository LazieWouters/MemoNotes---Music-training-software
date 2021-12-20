
''' <summary>
''' Cooley-Tukey FFT algorithm.
''' </summary>
Public NotInheritable Class FftAlgorithm
	Private Sub New()
	End Sub
	''' <summary>
	''' Calculates FFT using Cooley-Tukey FFT algorithm.
	''' </summary>
	''' <param name="x">input data</param>
	''' <returns>spectrogram of the data</returns>
	''' <remarks>
	''' If amount of data items not equal a power of 2, then algorithm
	''' automatically pad with 0s to the lowest amount of power of 2.
	''' </remarks>
	Public Shared Function Calculate(x As Double()) As Double()
		Dim length As Integer
		Dim bitsInLength As Integer
		If IsPowerOfTwo(x.Length) Then
			length = x.Length
			bitsInLength = Log2(length) - 1
		Else
			bitsInLength = Log2(x.Length)
				' the items will be pad with zeros
			length = 1 << bitsInLength
		End If

		' bit reversal
		Dim data As ComplexNumber() = New ComplexNumber(length - 1) {}
		For i As Integer = 0 To x.Length - 1
			Dim j As Integer = ReverseBits(i, bitsInLength)
			data(j) = New ComplexNumber(x(i))
		Next

		' Cooley-Tukey 
		For i As Integer = 0 To bitsInLength - 1
			Dim m As Integer = 1 << i
			Dim n As Integer = m * 2
			Dim alpha As Double = -(2 * Math.PI / n)

			For k As Integer = 0 To m - 1
				' e^(-2*pi/N*k)
				Dim oddPartMultiplier As ComplexNumber = New ComplexNumber(0, alpha * k).PoweredE()

				Dim j As Integer = k
				While j < length
					Dim evenPart As ComplexNumber = data(j)
					Dim oddPart As ComplexNumber = oddPartMultiplier * data(j + m)
					data(j) = evenPart + oddPart
					data(j + m) = evenPart - oddPart
					j += n
				End While
			Next
		Next

		' calculate spectrogram
		Dim spectrogram As Double() = New Double(length - 1) {}
		For i As Integer = 0 To spectrogram.Length - 1
			spectrogram(i) = data(i).AbsPower2()
		Next
		Return spectrogram
	End Function

	''' <summary>
	''' Gets number of significat bytes.
	''' </summary>
	''' <param name="n">Number</param>
	''' <returns>Amount of minimal bits to store the number.</returns>
	Private Shared Function Log2(n As Integer) As Integer
		Dim i As Integer = 0
		While n > 0
			i += 1
			n >>= 1
		End While
		Return i
	End Function

	''' <summary>
	''' Reverses bits in the number.
	''' </summary>
	''' <param name="n">Number</param>
	''' <param name="bitsCount">Significant bits in the number.</param>
	''' <returns>Reversed binary number.</returns>
	Private Shared Function ReverseBits(n As Integer, bitsCount As Integer) As Integer
		Dim reversed As Integer = 0
		For i As Integer = 0 To bitsCount - 1
			Dim nextBit As Integer = n And 1
			n >>= 1

			reversed <<= 1
			reversed = reversed Or nextBit
		Next
		Return reversed
	End Function

	''' <summary>
	''' Checks if number is power of 2.
	''' </summary>
	''' <param name="n">number</param>
	''' <returns>true if n=2^k and k is positive integer</returns>
	Private Shared Function IsPowerOfTwo(n As Integer) As Boolean
		Return n > 1 AndAlso (n And (n - 1)) = 0
	End Function
End Class


