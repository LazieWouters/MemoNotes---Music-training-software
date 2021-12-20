Public Class IndicadorDeFrequencia

    Private Sub IndicadorDeFrequencia_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        FrequenciesScale1.Size = Me.ClientRectangle.Size
        FrequenciesScale1.GerarImagemDeFundo()
    End Sub

End Class