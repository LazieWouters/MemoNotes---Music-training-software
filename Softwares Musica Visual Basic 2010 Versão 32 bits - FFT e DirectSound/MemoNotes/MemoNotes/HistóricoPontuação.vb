Public Class HistóricoPontuação

    Dim a As Integer = 1

    Private Sub DataGridView1_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting

        Try

            If DataGridView1.Columns(e.ColumnIndex).Index = 4 Then

                If e.Value < 95 AndAlso e.Value >= 90 Then
                    e.CellStyle.BackColor = Color.FromArgb(10, 213, 10)
                ElseIf e.Value < 90 AndAlso e.Value >= 80 Then
                    e.CellStyle.BackColor = Color.FromArgb(136, 255, 136)
                ElseIf e.Value < 80 AndAlso e.Value >= 70 Then
                    e.CellStyle.BackColor = Color.FromArgb(255, 255, 132)
                ElseIf e.Value < 70 AndAlso e.Value >= 60 Then
                    e.CellStyle.BackColor = Color.FromArgb(255, 255, 10)
                ElseIf e.Value < 60 AndAlso e.Value >= 40 Then
                    e.CellStyle.BackColor = Color.FromArgb(255, 80, 80)
                ElseIf e.Value < 40 AndAlso e.Value >= 20 Then
                    e.CellStyle.BackColor = Color.FromArgb(255, 40, 40)
                ElseIf e.Value < 20 AndAlso e.Value >= 0 Then
                    e.CellStyle.BackColor = Color.FromArgb(255, 10, 10)
                Else
                    e.CellStyle.BackColor = Color.FromArgb(10, 166, 10)
                End If


                e.Value = FormatNumber(e.Value, 3) & "%"
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub ComboBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.TextChanged

        Try

            If a = 1 Then
                a = 2
            Else
                MemoNotes.GerarGráfico()
            End If
            My.Settings.NovoValorQtdeRegistrosGráfico = ComboBox1.Text

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

End Class