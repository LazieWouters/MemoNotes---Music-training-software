
''' <summary>
''' Complex number.
''' </summary>
Structure ComplexNumber
	Public Re As Double
	Public Im As Double

	Public Sub New(re As Double)
		Me.Re = re
		Me.Im = 0
	End Sub

	Public Sub New(re As Double, im As Double)
		Me.Re = re
		Me.Im = im
	End Sub

	Public Shared Operator *(n1 As ComplexNumber, n2 As ComplexNumber) As ComplexNumber
		Return New ComplexNumber(n1.Re * n2.Re - n1.Im * n2.Im, n1.Im * n2.Re + n1.Re * n2.Im)
	End Operator

	Public Shared Operator +(n1 As ComplexNumber, n2 As ComplexNumber) As ComplexNumber
		Return New ComplexNumber(n1.Re + n2.Re, n1.Im + n2.Im)
	End Operator

	Public Shared Operator -(n1 As ComplexNumber, n2 As ComplexNumber) As ComplexNumber
		Return New ComplexNumber(n1.Re - n2.Re, n1.Im - n2.Im)
	End Operator

	Public Shared Operator -(n As ComplexNumber) As ComplexNumber
		Return New ComplexNumber(-n.Re, -n.Im)
	End Operator

	Public Shared Widening Operator CType(n As Double) As ComplexNumber
		Return New ComplexNumber(n, 0)
	End Operator

	Public Function PoweredE() As ComplexNumber
		Dim e As Double = Math.Exp(Re)
		Return New ComplexNumber(e * Math.Cos(Im), e * Math.Sin(Im))
	End Function

	Public Function Power2() As Double
		Return Re * Re - Im * Im
	End Function

	Public Function AbsPower2() As Double
		Return Re * Re + Im * Im
	End Function

	Public Overrides Function ToString() As String
		Return [String].Format("{0}+i*{1}", Re, Im)
	End Function
End Structure
