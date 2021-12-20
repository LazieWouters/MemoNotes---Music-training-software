Public Class OpcoesMemonotes

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try

            SalvaSettings()
            Me.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub SalvaSettings()
        Try
            For i = 0 To 40
                My.Settings.NovoValorCifras(i) = CifrasSelecionadas.GetItemChecked(i)
            Next

            For i = 0 To 6
                My.Settings.NovoValorExibirNotas(i) = NotasSelecionadas.GetItemChecked(i)
            Next

            For i = 0 To 6
                My.Settings.NovoValorIntervalosPauta(i + 2) = IntervalosSelecionados.GetItemChecked(i)
            Next

            For i = 0 To 4
                My.Settings.NovoValorTransposição(i) = TransposiçãoSelecionada.GetItemChecked(i)
            Next

            My.Settings.NovoValorExibirNumeraçãoDasNotas = OpçõesDiversas.GetItemChecked(0)
            My.Settings.NovoValorExibirNomeDaNota = OpçõesDiversas.GetItemChecked(1)
            My.Settings.NovoValorCoresNotas = OpçõesDiversas.GetItemChecked(2)
            My.Settings.NovoValorCoresArmadura = OpçõesDiversas.GetItemChecked(3)
            My.Settings.NovoValorIdentificarLinhas = OpçõesDiversas.GetItemChecked(4)
            My.Settings.NovoValorIdentificarPosiçãoDóCentral = OpçõesDiversas.GetItemChecked(5)
            My.Settings.NovoValorIdentificarPosiçãoNotasSOLeFÁ = OpçõesDiversas.GetItemChecked(6)
            My.Settings.NovoValorIdentificarPosiçãoTodasNotasDó = OpçõesDiversas.GetItemChecked(7)
            My.Settings.NovoValorExibirNumeraçãoEspaçosLinhasSuplementares = OpçõesDiversas.GetItemChecked(8)
            My.Settings.NovoValorLinhaSuplementar = OpçõesDiversas.GetItemChecked(9)

            My.Settings.NovoValorNotifyPointsInSecond = CInt(NumericUpDown1.Value)
            My.Settings.NovoValorBufferSeconds = CInt(NumericUpDown2.Value)
            My.Settings.NovoValorSampleRate = CInt(ComboBox1.Text)

            My.Settings.NovoValorTimers(12) = NumericUpDown3.Value
            My.Settings.NovoValorTimers(13) = NumericUpDown4.Value

            MemoNotes.ExibirLigadoDesligadoNomeDaNota()

            'MemoNotes.AtualizaRegiões()
            MemoNotes.ReligarCaptaçãoSom()
            MemoNotes.GerarImagemDeFundo()
            MemoNotes.GerarImagemDeFundo2()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub OpcoesDeCifras_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            For i = 0 To 40
                CifrasSelecionadas.SetItemChecked(i, My.Settings.NovoValorCifras(i))
            Next

            For i = 0 To 6
                NotasSelecionadas.SetItemChecked(i, My.Settings.NovoValorExibirNotas(i))
            Next

            For i = 0 To 6
                IntervalosSelecionados.SetItemChecked(i, My.Settings.NovoValorIntervalosPauta(i + 2))
            Next

            For i = 0 To 4
                TransposiçãoSelecionada.SetItemChecked(i, My.Settings.NovoValorTransposição(i))
            Next

            OpçõesDiversas.SetItemChecked(0, My.Settings.NovoValorExibirNumeraçãoDasNotas)
            OpçõesDiversas.SetItemChecked(1, My.Settings.NovoValorExibirNomeDaNota)
            OpçõesDiversas.SetItemChecked(2, My.Settings.NovoValorCoresNotas)
            OpçõesDiversas.SetItemChecked(3, My.Settings.NovoValorCoresArmadura)
            OpçõesDiversas.SetItemChecked(4, My.Settings.NovoValorIdentificarLinhas)
            OpçõesDiversas.SetItemChecked(5, My.Settings.NovoValorIdentificarPosiçãoDóCentral)
            OpçõesDiversas.SetItemChecked(6, My.Settings.NovoValorIdentificarPosiçãoNotasSOLeFÁ)
            OpçõesDiversas.SetItemChecked(7, My.Settings.NovoValorIdentificarPosiçãoTodasNotasDó)
            OpçõesDiversas.SetItemChecked(8, My.Settings.NovoValorExibirNumeraçãoEspaçosLinhasSuplementares)
            OpçõesDiversas.SetItemChecked(9, My.Settings.NovoValorLinhaSuplementar)

            NumericUpDown1.Value = My.Settings.NovoValorNotifyPointsInSecond
            NumericUpDown2.Value = My.Settings.NovoValorBufferSeconds

            NumericUpDown3.Value = My.Settings.NovoValorTimers(12)
            NumericUpDown4.Value = My.Settings.NovoValorTimers(13)

            ComboBox1.Text = CStr(My.Settings.NovoValorSampleRate)


            If My.Settings.NovoValorFonteCaptaçãoÁudio(1) = "True" Then
                TransposiçãoSelecionada.Enabled = False 'não será permitido alterar transposição se opção de captação de áudio for violão, isso é necessário pois a pauta do violão é escrita uma oitava acima do que soam as notas
            Else
                TransposiçãoSelecionada.Enabled = True
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub NotasSelecionadas_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotasSelecionadas.MouseUp
        Try

            If NotasSelecionadas.GetItemChecked(0) = False AndAlso NotasSelecionadas.GetItemChecked(1) = False AndAlso NotasSelecionadas.GetItemChecked(2) = False AndAlso NotasSelecionadas.GetItemChecked(3) = False _
                AndAlso NotasSelecionadas.GetItemChecked(4) = False AndAlso NotasSelecionadas.GetItemChecked(5) = False AndAlso NotasSelecionadas.GetItemChecked(6) = False Then
                NotasSelecionadas.SetItemChecked(0, True)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub TransposiçãoSelecionada_ItemCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles TransposiçãoSelecionada.ItemCheck
        Try
            'permite a seleção de apenas 1 item por vez
            If e.NewValue = CheckState.Checked Then
                For Each i In TransposiçãoSelecionada.CheckedIndices
                    TransposiçãoSelecionada.SetItemChecked(i, False)
                Next
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try

            SalvaSettings()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Try

            Me.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Public device As SoundCaptureDevice = Nothing
    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles SelecionarDispositivoAudio.Click
        Try

            Using form As New SelectDeviceForm()
                If form.ShowDialog() = DialogResult.OK Then
                    device = form.SelectedDevice
                Else
                    device = Nothing
                End If
            End Using

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub
End Class