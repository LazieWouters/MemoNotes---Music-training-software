Public Class Escolha_uma_escala

    Private Sub ListBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox2.SelectedIndexChanged

        Try

            VGoo = sender.Text
            NomeEscala = sender.Text
            Escalas.GeraEscalas()
            Me.BringToFront()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub
End Class