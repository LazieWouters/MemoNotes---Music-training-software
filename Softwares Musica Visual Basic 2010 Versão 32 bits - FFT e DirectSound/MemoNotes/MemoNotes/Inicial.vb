Public Class Inicial
    Inherits PerPixelAlphaForm 'para mecher no form no modo design é preciso mudar para 'Inherits Form'
    'e depois de ajustado voltar para 'Inherits PerPixelAlphaForm' e clicar para mudar o inherit
    Public TransAmount As Byte = 255

    Dim a, b As Integer
    Dim newPoint As New Point()
    Dim horainicial As Date = Date.Now
    Dim DataHora As String

    Private Sub Inicial_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            Process.GetCurrentProcess().PriorityClass = ProcessPriorityClass.RealTime
            Me.Top = -12
            ExibirEstaBarraSempreAcimaDasJanelasToolStripMenuItem.Checked = TopMost1.Checked
            If ExibirEstaBarraSempreAcimaDasJanelasToolStripMenuItem.Checked = True Then
                Me.TopMost = True
            Else
                Me.TopMost = False
            End If

            exibeControles()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub exibeControles() Handles Timer4.Tick, MenuStrip1.MouseHover

        Try

            Dim gr As Graphics
            Dim FaceBit As New Bitmap(My.Resources.BarraTelaInicial)
            gr = Graphics.FromImage(FaceBit)
            gr.SmoothingMode = SmoothingMode.AntiAlias

            For Each Control In Controls
                If Control.Visible Then
                    Dim Screenshot As New Bitmap(Me.Width, Me.Height)
                    Control.DrawToBitmap(Screenshot, New Rectangle(Control.Left, Control.Top, Control.Width, Control.Height))
                    gr.DrawImage(Screenshot, 0, 0)
                End If
            Next

            Dim CorFonte As SolidBrush = New SolidBrush(Color.FromArgb(0, 0, 0))
            Dim Fonte As New Font("Arial", 12, FontStyle.Bold, GraphicsUnit.Pixel)
            Dim MedidasTexto As SizeF = gr.MeasureString(DataHora, Fonte)
            Dim PosicaoX As Integer = (Me.Width / 2) - (MedidasTexto.Width / 2)
            Dim PosicaoY As Integer = 15

            gr.DrawString(DataHora, Fonte, CorFonte, PosicaoX, PosicaoY)

            Me.SetBitmap(FaceBit, TransAmount)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub Inicial_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        Try

            WinMM.midiInReset(VGhMidiIn) 'se não resetar o midiinclose não funcionará
            WinMM.midiInClose(VGhMidiIn)

            If MemoNotes.IsListenning Then
                MemoNotes.StopListenning()
            End If


            EstudoRitmico.microTimer.Enabled = False
            My.Settings.NovoValorTopMost = TopMost1.Checked
            'MemoNotes.SalvaSettings()
            'Escalas.SalvaSettings()
            'Teclado.SalvaSettings()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub ab_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown
        a = Me.MousePosition.X - Me.Location.X
        b = Me.MousePosition.Y - Me.Location.Y
    End Sub

    Private Sub ab_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove
        If e.Button = MouseButtons.Left Then
            newPoint = Me.MousePosition
            newPoint.X = newPoint.X - a
            newPoint.Y = newPoint.Y - b
            Me.Location = newPoint
        End If
    End Sub

    Private Sub FecharToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FecharToolStripMenuItem.Click
        Try

            exibeControles()
            Me.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        Try

            exibeControles()
            Me.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Timer4_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer4.Tick, Me.Load
        Try

            DataHora = FormatDateTime(Now, DateFormat.LongDate) & " - " & FormatDateTime(Now, DateFormat.LongTime)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub ExibirEstaBarraSempreAcimaDasJanelasToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExibirEstaBarraSempreAcimaDasJanelasToolStripMenuItem.Click
        Try

            exibeControles()
            ExibirEstaBarraSempreAcimaDasJanelasToolStripMenuItem.Checked = Not ExibirEstaBarraSempreAcimaDasJanelasToolStripMenuItem.Checked
            If ExibirEstaBarraSempreAcimaDasJanelasToolStripMenuItem.Checked = True Then
                Me.TopMost = True
            Else
                Me.TopMost = False
            End If
            TopMost1.Checked = ExibirEstaBarraSempreAcimaDasJanelasToolStripMenuItem.Checked

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub MenuToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuToolStripMenuItem.Click, MemorizaçãoDasNotasToolStripMenuItem.Click
        Try

            exibeControles()
            MemoNotes.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub EstudoDasEscalasToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EstudoDasEscalasToolStripMenuItem.Click, EstudoDasEscalasToolStripMenuItem1.Click
        Try

            exibeControles()
            Escalas.Show()

            Teclado.Top = Escalas.Top + 674
            Teclado.Left = Escalas.Left
            Teclado.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub SobreToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SobreToolStripMenuItem.Click, SobreToolStripMenuItem1.Click
        Try

            exibeControles()
            AboutBox1.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Inicial_FormClosing(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
        MemoNotes.RadioButton1.Checked = True
    End Sub

    Private Sub EstudoRitmicoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EstudoRitmicoToolStripMenuItem.Click, EstudoRítmicoToolStripMenuItem.Click
        Try

            exibeControles()
            EstudoRitmico.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub TreinamentoDinâmicaMãosEsquerdaEDireitaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TreinamentoDinâmicaMãosEsquerdaEDireitaToolStripMenuItem.Click, TreinamentoDinâmicaMãosEsquerdaEDireitaToolStripMenuItem1.Click
        Try

            exibeControles()
            TreinamentoDinâmicaMãoEsquerdaDireita.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub AcordesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AcordesToolStripMenuItem.Click, AcordesToolStripMenuItem1.Click
        Try

            exibeControles()
            Acordes.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub ReconhecimentoDasArmadurasDeClaveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReconhecimentoDasArmadurasDeClaveToolStripMenuItem.Click, ReconhecimentoDasArmadurasDeClaveToolStripMenuItem1.Click
        Try

            exibeControles()
            ReconhecimentoArmaduraClave.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub ReconhecimentoDosAcordesDeUmaTonalidadeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReconhecimentoDosAcordesDeUmaTonalidadeToolStripMenuItem.Click, ReconhecimentoDosAcordesDeUmaTonalidadeToolStripMenuItem1.Click
        Try

            exibeControles()
            ReconhecimentoAcordesTonalidade.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub TreinamentoLocalizaçãoGrausNaEscalaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TreinamentoLocalizaçãoGrausNaEscalaToolStripMenuItem.Click, TreinamentoLocalizaçãoDosGrausNaEscalaToolStripMenuItem.Click
        Try

            exibeControles()
            TreinamentoGrausEscala.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub MemorizaçãoDasNotasNoViolãoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MemorizaçãoDasNotasNoViolãoToolStripMenuItem.Click, MemorizaçãoDasNotasNoViolãoToolStripMenuItem1.Click
        Try

            exibeControles()
            MemorizacaoNotasViolao.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub TreinamentoDeIntervalosToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TreinamentoDeIntervalosToolStripMenuItem.Click, TreinamentoDeIntervalosToolStripMenuItem1.Click
        Try

            exibeControles()
            TreinamentoIntervalos.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub ExercícioParaMemorizarEscalasToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExercícioParaMemorizarEscalasToolStripMenuItem.Click, ExercícioParaMemorizarEscalasToolStripMenuItem1.Click
        Try

            exibeControles()
            ExercicioParaMemorizarEscalas.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub EstudosDosIntervalosDosAcordesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EstudosDosIntervalosDosAcordesToolStripMenuItem.Click, EstudosDosIntervalosDosAcordesToolStripMenuItem1.Click
        Try

            exibeControles()
            EstudoIntervalosAcordes.Show()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub
End Class