Public Class OpçõesTelaAcordes

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        My.Settings.NovoValorNumeraçãoTrastes = Color.FromArgb(150, 150, 150)
        My.Settings.NovoValorLinhasDiagramaAcordes = Color.FromArgb(200, 200, 200)
        My.Settings.NovoValorLinhasDiagramaAcordes = Color.FromArgb(200, 200, 200)
        My.Settings.NovoValorAcordesMaisUsados = Color.FromArgb(153, 217, 234)
        My.Settings.NovoValorBolinhaAcordes = Color.FromArgb(0, 0, 0)
        My.Settings.NovoValorNumeraçãoDedilhados1 = Color.FromArgb(0, 0, 0)
        My.Settings.NovoValorNumeraçãoDedilhados2 = Color.FromArgb(255, 255, 255)
        My.Settings.NovoValorNotaDeReferênciaAcorde = Color.FromArgb(255, 0, 0)
        My.Settings.NovoValorCorPestanas = Color.FromArgb(0, 0, 0)
        My.Settings.NovoValorCorPestanas = Color.FromArgb(0, 0, 0)
        My.Settings.NovoValorCorDasCifras = Color.FromArgb(0, 0, 0)
        AtualizaCoresNosLabels()
        Acordes.Desenhar()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub OpçõesTelaAcordes_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        AtualizaCoresNosLabels()
    End Sub

    Private Sub AtualizaCoresNosLabels()
        Label7.BackColor = My.Settings.NovoValorNumeraçãoTrastes
        Label8.BackColor = My.Settings.NovoValorLinhasDiagramaAcordes
        Label9.BackColor = My.Settings.NovoValorAcordesMaisUsados
        Label10.BackColor = My.Settings.NovoValorBolinhaAcordes
        Label11.BackColor = My.Settings.NovoValorNumeraçãoDedilhados1
        Label13.BackColor = My.Settings.NovoValorNumeraçãoDedilhados2
        Label12.BackColor = My.Settings.NovoValorNotaDeReferênciaAcorde
        Label14.BackColor = My.Settings.NovoValorCorPestanas
        Label16.BackColor = My.Settings.NovoValorCorDasCifras
    End Sub

    Private Sub Label7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label7.Click
        ColorDialog1.Color = My.Settings.NovoValorNumeraçãoTrastes
        ColorDialog1.ShowDialog()
        My.Settings.NovoValorNumeraçãoTrastes = ColorDialog1.Color
        AtualizaCoresNosLabels() : Acordes.Desenhar()
    End Sub

    Private Sub Label8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label8.Click
        ColorDialog1.Color = My.Settings.NovoValorLinhasDiagramaAcordes
        ColorDialog1.ShowDialog()
        My.Settings.NovoValorLinhasDiagramaAcordes = ColorDialog1.Color
        AtualizaCoresNosLabels() : Acordes.Desenhar()
    End Sub

    Private Sub Label9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label9.Click
        ColorDialog1.Color = My.Settings.NovoValorAcordesMaisUsados
        ColorDialog1.ShowDialog()
        My.Settings.NovoValorAcordesMaisUsados = ColorDialog1.Color
        AtualizaCoresNosLabels() : Acordes.Desenhar()
    End Sub

    Private Sub Label10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label10.Click
        ColorDialog1.Color = My.Settings.NovoValorBolinhaAcordes
        ColorDialog1.ShowDialog()
        My.Settings.NovoValorBolinhaAcordes = ColorDialog1.Color
        AtualizaCoresNosLabels() : Acordes.Desenhar()
    End Sub

    Private Sub Label11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label11.Click
        ColorDialog1.Color = My.Settings.NovoValorNumeraçãoDedilhados1
        ColorDialog1.ShowDialog()
        My.Settings.NovoValorNumeraçãoDedilhados1 = ColorDialog1.Color
        AtualizaCoresNosLabels() : Acordes.Desenhar()
    End Sub

    Private Sub Label13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label13.Click
        ColorDialog1.Color = My.Settings.NovoValorNumeraçãoDedilhados2
        ColorDialog1.ShowDialog()
        My.Settings.NovoValorNumeraçãoDedilhados2 = ColorDialog1.Color
        AtualizaCoresNosLabels() : Acordes.Desenhar()
    End Sub

    Private Sub Label12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label12.Click
        ColorDialog1.Color = My.Settings.NovoValorNotaDeReferênciaAcorde
        ColorDialog1.ShowDialog()
        My.Settings.NovoValorNotaDeReferênciaAcorde = ColorDialog1.Color
        AtualizaCoresNosLabels() : Acordes.Desenhar()
    End Sub

    Private Sub Label14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label14.Click
        ColorDialog1.Color = My.Settings.NovoValorCorPestanas
        ColorDialog1.ShowDialog()
        My.Settings.NovoValorCorPestanas = ColorDialog1.Color
        AtualizaCoresNosLabels() : Acordes.Desenhar()
    End Sub

    Private Sub Label16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label16.Click
        ColorDialog1.Color = My.Settings.NovoValorCorDasCifras
        ColorDialog1.ShowDialog()
        My.Settings.NovoValorCorDasCifras = ColorDialog1.Color
        AtualizaCoresNosLabels() : Acordes.Desenhar()
    End Sub
End Class