Public Class TelaAcordesModoJogo

    Private Sub TelaAcordesModoJogo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try

            For i = 0 To 12
                CheckedListBox1.SetItemChecked(i, My.Settings.NovoValorModoJogoTelaAcordesCifras(i))
            Next

            For i = 0 To 14
                CheckedListBox2.SetItemChecked(i, My.Settings.NovoValorModoJogoTelaAcordesCifras(i + 13))
            Next

            For i = 0 To 23
                CheckedListBox3.SetItemChecked(i, My.Settings.NovoValorModoJogoTelaAcordesCifras(i + 29))
            Next

            CheckedListBox4.SetItemChecked(0, My.Settings.NovoValorModoJogoTelaAcordesCifras(28))
            CheckedListBox4.SetItemChecked(1, My.Settings.NovoValorModoJogoTelaAcordesCifras(53))

            For i = 0 To 6
                CheckedListBox5.SetItemChecked(i, My.Settings.NovoValorModoJogoTelaAcordesTom(i))
            Next

            CheckBox1.Checked = My.Settings.NovoValorExibirAcordesInvertidos
            CheckBox2.Checked = My.Settings.NovoValorExibirAcordesMaisUsadosModoJogo

            TrocarAcordesNoIntervaloDeTempo.Checked = My.Settings.NovoValorTrocarAcordesNoIntervaloDeTempo

            ExibiçãoDosBotões()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub CheckedListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckedListBox1.SelectedIndexChanged
        Try
            For i = 0 To 12
                My.Settings.NovoValorModoJogoTelaAcordesCifras(i) = CheckedListBox1.GetItemChecked(i)
            Next

            ExibiçãoDosBotões()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub CheckedListBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckedListBox2.SelectedIndexChanged
        Try
            For i = 0 To 14
                My.Settings.NovoValorModoJogoTelaAcordesCifras(i + 13) = CheckedListBox2.GetItemChecked(i)
            Next

            ExibiçãoDosBotões()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub CheckedListBox4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckedListBox4.SelectedIndexChanged
        Try
            My.Settings.NovoValorModoJogoTelaAcordesCifras(28) = CheckedListBox4.GetItemChecked(0)
            My.Settings.NovoValorModoJogoTelaAcordesCifras(53) = CheckedListBox4.GetItemChecked(1)

            ExibiçãoDosBotões()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub CheckedListBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckedListBox3.SelectedIndexChanged
        Try
            For i = 0 To 23
                My.Settings.NovoValorModoJogoTelaAcordesCifras(i + 29) = CheckedListBox3.GetItemChecked(i)
            Next

            ExibiçãoDosBotões()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub CheckedListBox5_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckedListBox5.SelectedIndexChanged
        Try
            For i = 0 To 6
                My.Settings.NovoValorModoJogoTelaAcordesTom(i) = CheckedListBox5.GetItemChecked(i)
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Button54_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click, Button8.Click, Button7.Click, Button6.Click, Button54.Click, Button53.Click, Button52.Click, Button51.Click, Button50.Click, Button5.Click, Button49.Click, Button48.Click, Button47.Click, Button46.Click, Button45.Click, Button44.Click, Button43.Click, Button42.Click, Button41.Click, Button40.Click, Button4.Click, Button39.Click, Button38.Click, Button37.Click, Button36.Click, Button35.Click, Button34.Click, Button33.Click, Button32.Click, Button31.Click, Button30.Click, Button3.Click, Button29.Click, Button28.Click, Button27.Click, Button26.Click, Button25.Click, Button24.Click, Button23.Click, Button22.Click, Button21.Click, Button20.Click, Button2.Click, Button19.Click, Button18.Click, Button17.Click, Button16.Click, Button15.Click, Button14.Click, Button13.Click, Button12.Click, Button11.Click, Button10.Click, Button1.Click
        Try
            Dim Botão As Button = sender
            If Botão.Tag = NumeraçãoFamiliaAcorde Then
                Label1.Text = "Correto"
            Else
                Label1.Text = "Incorreto"
            End If

            Acordes.ExibeTodosOsAcordesParaResposta()

            Timer1.Enabled = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try
            Label1.Text = ""
            Acordes.GerarNovoAcordeJogo()
            Timer1.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub ExibiçãoDosBotões()
        Try
            Panel1.AutoScroll = False

            Dim BotãoTopo As Integer = -20

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(0) = "True" Then
                BotãoTopo += 24
                Button1.Top = BotãoTopo
                Button1.Visible = True
            Else
                Button1.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(1) = "True" Then
                BotãoTopo += 24
                Button2.Top = BotãoTopo
                Button2.Visible = True
            Else
                Button2.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(2) = "True" Then
                BotãoTopo += 24
                Button3.Top = BotãoTopo
                Button3.Visible = True
            Else
                Button3.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(3) = "True" Then
                BotãoTopo += 24
                Button4.Top = BotãoTopo
                Button4.Visible = True
            Else
                Button4.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(4) = "True" Then
                BotãoTopo += 24
                Button5.Top = BotãoTopo
                Button5.Visible = True
            Else
                Button5.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(5) = "True" Then
                BotãoTopo += 24
                Button6.Top = BotãoTopo
                Button6.Visible = True
            Else
                Button6.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(6) = "True" Then
                BotãoTopo += 24
                Button7.Top = BotãoTopo
                Button7.Visible = True
            Else
                Button7.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(7) = "True" Then
                BotãoTopo += 24
                Button8.Top = BotãoTopo
                Button8.Visible = True
            Else
                Button8.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(8) = "True" Then
                BotãoTopo += 24
                Button9.Top = BotãoTopo
                Button9.Visible = True
            Else
                Button9.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(9) = "True" Then
                BotãoTopo += 24
                Button10.Top = BotãoTopo
                Button10.Visible = True
            Else
                Button10.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(10) = "True" Then
                BotãoTopo += 24
                Button11.Top = BotãoTopo
                Button11.Visible = True
            Else
                Button11.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(11) = "True" Then
                BotãoTopo += 24
                Button12.Top = BotãoTopo
                Button12.Visible = True
            Else
                Button12.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(12) = "True" Then
                BotãoTopo += 24
                Button13.Top = BotãoTopo
                Button13.Visible = True
            Else
                Button13.Visible = False
            End If

            BotãoTopo = -20
            If My.Settings.NovoValorModoJogoTelaAcordesCifras(13) = "True" Then
                BotãoTopo += 24
                Button14.Top = BotãoTopo
                Button14.Visible = True
            Else
                Button14.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(14) = "True" Then
                BotãoTopo += 24
                Button15.Top = BotãoTopo
                Button15.Visible = True
            Else
                Button15.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(15) = "True" Then
                BotãoTopo += 24
                Button16.Top = BotãoTopo
                Button16.Visible = True
            Else
                Button16.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(16) = "True" Then
                BotãoTopo += 24
                Button17.Top = BotãoTopo
                Button17.Visible = True
            Else
                Button17.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(17) = "True" Then
                BotãoTopo += 24
                Button18.Top = BotãoTopo
                Button18.Visible = True
            Else
                Button18.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(18) = "True" Then
                BotãoTopo += 24
                Button19.Top = BotãoTopo
                Button19.Visible = True
            Else
                Button19.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(19) = "True" Then
                BotãoTopo += 24
                Button20.Top = BotãoTopo
                Button20.Visible = True
            Else
                Button20.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(20) = "True" Then
                BotãoTopo += 24
                Button21.Top = BotãoTopo
                Button21.Visible = True
            Else
                Button21.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(21) = "True" Then
                BotãoTopo += 24
                Button22.Top = BotãoTopo
                Button22.Visible = True
            Else
                Button22.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(22) = "True" Then
                BotãoTopo += 24
                Button23.Top = BotãoTopo
                Button23.Visible = True
            Else
                Button23.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(23) = "True" Then
                BotãoTopo += 24
                Button24.Top = BotãoTopo
                Button24.Visible = True
            Else
                Button24.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(24) = "True" Then
                BotãoTopo += 24
                Button25.Top = BotãoTopo
                Button25.Visible = True
            Else
                Button25.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(25) = "True" Then
                BotãoTopo += 24
                Button26.Top = BotãoTopo
                Button26.Visible = True
            Else
                Button26.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(26) = "True" Then
                BotãoTopo += 24
                Button27.Top = BotãoTopo
                Button27.Visible = True
            Else
                Button27.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(27) = "True" Then
                BotãoTopo += 24
                Button28.Top = BotãoTopo
                Button28.Visible = True
            Else
                Button28.Visible = False
            End If

            BotãoTopo = -20
            If My.Settings.NovoValorModoJogoTelaAcordesCifras(29) = "True" Then
                BotãoTopo += 24
                Button30.Top = BotãoTopo
                Button30.Visible = True
            Else
                Button30.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(30) = "True" Then
                BotãoTopo += 24
                Button31.Top = BotãoTopo
                Button31.Visible = True
            Else
                Button31.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(31) = "True" Then
                BotãoTopo += 24
                Button32.Top = BotãoTopo
                Button32.Visible = True
            Else
                Button32.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(32) = "True" Then
                BotãoTopo += 24
                Button33.Top = BotãoTopo
                Button33.Visible = True
            Else
                Button33.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(33) = "True" Then
                BotãoTopo += 24
                Button34.Top = BotãoTopo
                Button34.Visible = True
            Else
                Button34.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(34) = "True" Then
                BotãoTopo += 24
                Button35.Top = BotãoTopo
                Button35.Visible = True
            Else
                Button35.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(35) = "True" Then
                BotãoTopo += 24
                Button36.Top = BotãoTopo
                Button36.Visible = True
            Else
                Button36.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(36) = "True" Then
                BotãoTopo += 24
                Button37.Top = BotãoTopo
                Button37.Visible = True
            Else
                Button37.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(37) = "True" Then
                BotãoTopo += 24
                Button38.Top = BotãoTopo
                Button38.Visible = True
            Else
                Button38.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(38) = "True" Then
                BotãoTopo += 24
                Button39.Top = BotãoTopo
                Button39.Visible = True
            Else
                Button39.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(39) = "True" Then
                BotãoTopo += 24
                Button40.Top = BotãoTopo
                Button40.Visible = True
            Else
                Button40.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(40) = "True" Then
                BotãoTopo += 24
                Button41.Top = BotãoTopo
                Button41.Visible = True
            Else
                Button41.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(41) = "True" Then
                BotãoTopo += 24
                Button42.Top = BotãoTopo
                Button42.Visible = True
            Else
                Button42.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(42) = "True" Then
                BotãoTopo += 24
                Button43.Top = BotãoTopo
                Button43.Visible = True
            Else
                Button43.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(43) = "True" Then
                BotãoTopo += 24
                Button44.Top = BotãoTopo
                Button44.Visible = True
            Else
                Button44.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(44) = "True" Then
                BotãoTopo += 24
                Button45.Top = BotãoTopo
                Button45.Visible = True
            Else
                Button45.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(45) = "True" Then
                BotãoTopo += 24
                Button46.Top = BotãoTopo
                Button46.Visible = True
            Else
                Button46.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(46) = "True" Then
                BotãoTopo += 24
                Button47.Top = BotãoTopo
                Button47.Visible = True
            Else
                Button47.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(47) = "True" Then
                BotãoTopo += 24
                Button48.Top = BotãoTopo
                Button48.Visible = True
            Else
                Button48.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(48) = "True" Then
                BotãoTopo += 24
                Button49.Top = BotãoTopo
                Button49.Visible = True
            Else
                Button49.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(49) = "True" Then
                BotãoTopo += 24
                Button50.Top = BotãoTopo
                Button50.Visible = True
            Else
                Button50.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(50) = "True" Then
                BotãoTopo += 24
                Button51.Top = BotãoTopo
                Button51.Visible = True
            Else
                Button51.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(51) = "True" Then
                BotãoTopo += 24
                Button52.Top = BotãoTopo
                Button52.Visible = True
            Else
                Button52.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(52) = "True" Then
                BotãoTopo += 24
                Button53.Top = BotãoTopo
                Button53.Visible = True
            Else
                Button53.Visible = False
            End If

            BotãoTopo = -20
            If My.Settings.NovoValorModoJogoTelaAcordesCifras(28) = "True" Then
                BotãoTopo += 24
                Button29.Top = BotãoTopo
                Button29.Visible = True
            Else
                Button29.Visible = False
            End If

            If My.Settings.NovoValorModoJogoTelaAcordesCifras(53) = "True" Then
                BotãoTopo += 24
                Button54.Top = BotãoTopo
                Button54.Visible = True
            Else
                Button54.Visible = False
            End If

            Panel1.AutoScroll = True


            If Button2.Visible OrElse Button3.Visible OrElse Button15.Visible OrElse Button16.Visible OrElse _
                Button21.Visible OrElse Button31.Visible OrElse Button32.Visible OrElse Button33.Visible Then
                CheckBox1.Checked = True
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Button55_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button55.Click
        Try
            If Me.Width = 345 Then
                Me.Width = 760
                Me.Height = 550
            Else
                Me.Width = 345
                Me.Height = 385
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        Try
            My.Settings.NovoValorExibirAcordesInvertidos = CheckBox1.Checked
            If Not CheckBox1.Checked Then
                CheckedListBox1.SetItemChecked(1, False) : My.Settings.NovoValorModoJogoTelaAcordesCifras(1) = CheckedListBox1.GetItemChecked(1)
                CheckedListBox1.SetItemChecked(2, False) : My.Settings.NovoValorModoJogoTelaAcordesCifras(2) = CheckedListBox1.GetItemChecked(2)
                CheckedListBox2.SetItemChecked(1, False) : My.Settings.NovoValorModoJogoTelaAcordesCifras(14) = CheckedListBox2.GetItemChecked(1)
                CheckedListBox2.SetItemChecked(2, False) : My.Settings.NovoValorModoJogoTelaAcordesCifras(15) = CheckedListBox2.GetItemChecked(2)
                CheckedListBox2.SetItemChecked(7, False) : My.Settings.NovoValorModoJogoTelaAcordesCifras(20) = CheckedListBox2.GetItemChecked(7)
                CheckedListBox3.SetItemChecked(1, False) : My.Settings.NovoValorModoJogoTelaAcordesCifras(30) = CheckedListBox3.GetItemChecked(1)
                CheckedListBox3.SetItemChecked(2, False) : My.Settings.NovoValorModoJogoTelaAcordesCifras(31) = CheckedListBox3.GetItemChecked(2)
                CheckedListBox3.SetItemChecked(3, False) : My.Settings.NovoValorModoJogoTelaAcordesCifras(32) = CheckedListBox3.GetItemChecked(3)

                ExibiçãoDosBotões()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        Try

            My.Settings.NovoValorExibirAcordesMaisUsadosModoJogo = CheckBox2.Checked

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrocarAcordesNoIntervaloDeTempo.CheckedChanged
        Try

            My.Settings.NovoValorTrocarAcordesNoIntervaloDeTempo = TrocarAcordesNoIntervaloDeTempo.Checked

            If TrocarAcordesNoIntervaloDeTempo.Checked Then
                Timer2.Interval = My.Settings.NovoValorTimerTrocaDeAcorde
                Timer2.Enabled = True
            Else
                Timer2.Enabled = False
            End If

            Panel1.Enabled = Not Timer2.Enabled

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub NumericUpDown1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown1.ValueChanged
        Try

            My.Settings.NovoValorTimerTrocaDeAcorde = NumericUpDown1.Value
            Timer2.Interval = My.Settings.NovoValorTimerTrocaDeAcorde

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Try

            Acordes.GerarNovoAcordeJogo()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub
End Class