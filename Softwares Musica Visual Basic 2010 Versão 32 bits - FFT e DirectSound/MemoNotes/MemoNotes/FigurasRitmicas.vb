Public Class FigurasRitmicas
    Dim Cor1 As SolidBrush = New SolidBrush(Color.FromArgb(50, 255, 0, 0))
    Dim Cor2 As SolidBrush = New SolidBrush(Color.FromArgb(50, 0, 0, 255))
    Dim Cor3 As SolidBrush = New SolidBrush(Color.FromArgb(50, 255, 255, 0))
    Dim Cor4 As SolidBrush = New SolidBrush(Color.FromArgb(150, 0, 128, 64))

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try

            SalvaSettings()
            EstudoRitmico.microTimer.Enabled = False
            EstudoRitmico.GerarExercícioRitmico()
            EstudoRitmico.AtualizaRegiãoDoTempo()
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

            My.Settings.NovoValorFiguraRitmicas(1) = CheckBox1.Checked
            My.Settings.NovoValorFiguraRitmicas(2) = CheckBox2.Checked
            My.Settings.NovoValorFiguraRitmicas(3) = CheckBox3.Checked
            My.Settings.NovoValorFiguraRitmicas(4) = CheckBox4.Checked
            My.Settings.NovoValorFiguraRitmicas(5) = CheckBox5.Checked
            My.Settings.NovoValorFiguraRitmicas(6) = CheckBox6.Checked
            My.Settings.NovoValorFiguraRitmicas(7) = CheckBox7.Checked
            My.Settings.NovoValorFiguraRitmicas(8) = CheckBox8.Checked
            My.Settings.NovoValorFiguraRitmicas(9) = CheckBox9.Checked
            My.Settings.NovoValorFiguraRitmicas(10) = CheckBox10.Checked
            My.Settings.NovoValorFiguraRitmicas(11) = CheckBox11.Checked
            My.Settings.NovoValorFiguraRitmicas(12) = CheckBox12.Checked
            My.Settings.NovoValorFiguraRitmicas(13) = CheckBox13.Checked
            My.Settings.NovoValorFiguraRitmicas(14) = CheckBox14.Checked
            My.Settings.NovoValorFiguraRitmicas(15) = CheckBox15.Checked

            My.Settings.NovoValorFiguraRitmicas(17) = CheckBox17.Checked
            My.Settings.NovoValorFiguraRitmicas(18) = CheckBox18.Checked
            My.Settings.NovoValorFiguraRitmicas(19) = CheckBox19.Checked
            My.Settings.NovoValorFiguraRitmicas(20) = CheckBox20.Checked
            My.Settings.NovoValorFiguraRitmicas(21) = CheckBox21.Checked
            My.Settings.NovoValorFiguraRitmicas(22) = CheckBox22.Checked
            My.Settings.NovoValorFiguraRitmicas(23) = CheckBox23.Checked
            My.Settings.NovoValorFiguraRitmicas(24) = CheckBox24.Checked
            My.Settings.NovoValorFiguraRitmicas(25) = CheckBox25.Checked
            My.Settings.NovoValorFiguraRitmicas(26) = CheckBox26.Checked
            My.Settings.NovoValorFiguraRitmicas(27) = CheckBox27.Checked
            My.Settings.NovoValorFiguraRitmicas(28) = CheckBox28.Checked
            My.Settings.NovoValorFiguraRitmicas(29) = CheckBox29.Checked
            My.Settings.NovoValorFiguraRitmicas(30) = CheckBox30.Checked
            My.Settings.NovoValorFiguraRitmicas(31) = CheckBox31.Checked


            My.Settings.NovoValorTempoInicial = NumericUpDown1.Value
            My.Settings.NovoValorMãoDireita = MãoDireita.Checked
            My.Settings.NovoValorMãoEsquerda = MãoEsquerda.Checked

            My.Settings.NovoValorPontoAumentoME = PontoAumentoME.Checked
            My.Settings.NovoValorPontoAumentoMD = PontoAumentoMD.Checked
            My.Settings.NovoValorDuploPontoAumentoME = DuploPontoAumentoME.Checked
            My.Settings.NovoValorDuploPontoAumentoMD = DuploPontoAumentoMD.Checked
            My.Settings.NovoValorTriploPontoAumentoME = TriploPontoAumentoME.Checked
            My.Settings.NovoValorTriploPontoAumentoMD = TriploPontoAumentoMD.Checked

            My.Settings.NovoValorLigadurasMD = LigadurasMD.Checked
            My.Settings.NovoValorLigadurasME = LigadurasME.Checked

            My.Settings.NovoValorSwingMD = SwingMD.Checked
            My.Settings.NovoValorSwingME = SwingME.Checked

            My.Settings.NovoValorDinamicaME = DinamicaME.Checked
            My.Settings.NovoValorDinamicaMD = DinamicaMD.Checked
            My.Settings.NovoValorPercentualDinamicaME = PercentualDinamicaME.Value
            My.Settings.NovoValorPercentualDinamicaMD = PercentualDinamicaMD.Value
            My.Settings.NovoValorPercentualLigadurasMD = PercentualLigadurasMD.Value
            My.Settings.NovoValorPercentualLigadurasME = PercentualLigadurasME.Value

            My.Settings.NovoValorLocalizaçãoCompassoAtual = LocalizaçãoCompassoAtual.Checked
            My.Settings.NovoValorLocalizaçãoSubdivisãoCompassoAtual = LocalizaçãoSubdivisaoCompassoAtual.Checked
            My.Settings.NovoValorLocalizaçãoMicroSubdivisãoCompassoAtual = LocalizaçãoMicroSubdivisão.Checked


            My.Settings.NovoValorCompassos(0) = Compasso2_2.Checked
            My.Settings.NovoValorCompassos(1) = Compasso3_2.Checked
            My.Settings.NovoValorCompassos(2) = Compasso4_2.Checked
            My.Settings.NovoValorCompassos(3) = Compasso2_4.Checked
            My.Settings.NovoValorCompassos(4) = Compasso3_4.Checked
            My.Settings.NovoValorCompassos(5) = Compasso4_4.Checked
            My.Settings.NovoValorCompassos(6) = Compasso5_4.Checked
            My.Settings.NovoValorCompassos(7) = Compasso6_4.Checked
            My.Settings.NovoValorCompassos(8) = Compasso9_4.Checked
            My.Settings.NovoValorCompassos(9) = Compasso12_4.Checked
            My.Settings.NovoValorCompassos(10) = Compasso3_8.Checked
            My.Settings.NovoValorCompassos(11) = Compasso4_8.Checked
            My.Settings.NovoValorCompassos(12) = Compasso5_8.Checked
            My.Settings.NovoValorCompassos(13) = Compasso6_8.Checked
            My.Settings.NovoValorCompassos(14) = Compasso7_8.Checked
            My.Settings.NovoValorCompassos(15) = Compasso9_8.Checked
            My.Settings.NovoValorCompassos(16) = Compasso12_8.Checked

            My.Settings.NovoValorForçaDinâmica(8) = CheckBox42.Checked
            My.Settings.NovoValorForçaDinâmica(9) = CheckBox41.Checked
            My.Settings.NovoValorForçaDinâmica(10) = CheckBox36.Checked
            My.Settings.NovoValorForçaDinâmica(11) = CheckBox37.Checked
            My.Settings.NovoValorForçaDinâmica(12) = CheckBox40.Checked
            My.Settings.NovoValorForçaDinâmica(13) = CheckBox39.Checked
            My.Settings.NovoValorForçaDinâmica(14) = CheckBox38.Checked
            My.Settings.NovoValorForçaDinâmica(15) = CheckBox35.Checked

            My.Settings.NovoValorForçaDinâmica(0) = CheckBox50.Checked
            My.Settings.NovoValorForçaDinâmica(1) = CheckBox49.Checked
            My.Settings.NovoValorForçaDinâmica(2) = CheckBox44.Checked
            My.Settings.NovoValorForçaDinâmica(3) = CheckBox45.Checked
            My.Settings.NovoValorForçaDinâmica(4) = CheckBox48.Checked
            My.Settings.NovoValorForçaDinâmica(5) = CheckBox47.Checked
            My.Settings.NovoValorForçaDinâmica(6) = CheckBox46.Checked
            My.Settings.NovoValorForçaDinâmica(7) = CheckBox43.Checked

            My.Settings.NovoValorCondescendente = Condescendente.Checked
            My.Settings.NovoValorNormal = Normal.Checked
            My.Settings.NovoValorSevera = Severa.Checked

            My.Settings.NovaTransparênciaCorLocalizaçãoDosCompassos = PercentualTransparenciaCor.Value

            My.Settings.NovoValorTocarSonsMetronomo = CheckBox33.Checked
            My.Settings.NovoValorMetrônomoTempoForte = ComboBox1.Text
            My.Settings.NovoValorMetrônomoTempoFraco = ComboBox2.Text
            My.Settings.NovoValorSomRitmo = ComboBox3.Text
            My.Settings.NovoValorVolumeTempoForte = NumericUpDown2.Value
            My.Settings.NovoValorVolumeTempoFraco = NumericUpDown3.Value
            My.Settings.NovoValorVolumeRitmo = NumericUpDown4.Value

            My.Settings.NovoValorTeclaMãoEsquerda = ComboBox4.Text
            My.Settings.NovoValorTeclaMãoDireita = ComboBox5.Text

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub FigurasRitmicas_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try

            CheckBox1.Checked = My.Settings.NovoValorFiguraRitmicas(1)
            CheckBox2.Checked = My.Settings.NovoValorFiguraRitmicas(2)
            CheckBox3.Checked = My.Settings.NovoValorFiguraRitmicas(3)
            CheckBox4.Checked = My.Settings.NovoValorFiguraRitmicas(4)
            CheckBox5.Checked = My.Settings.NovoValorFiguraRitmicas(5)
            CheckBox6.Checked = My.Settings.NovoValorFiguraRitmicas(6)
            CheckBox7.Checked = My.Settings.NovoValorFiguraRitmicas(7)
            CheckBox8.Checked = My.Settings.NovoValorFiguraRitmicas(8)
            CheckBox9.Checked = My.Settings.NovoValorFiguraRitmicas(9)
            CheckBox10.Checked = My.Settings.NovoValorFiguraRitmicas(10)
            CheckBox11.Checked = My.Settings.NovoValorFiguraRitmicas(11)
            CheckBox12.Checked = My.Settings.NovoValorFiguraRitmicas(12)
            CheckBox13.Checked = My.Settings.NovoValorFiguraRitmicas(13)
            CheckBox14.Checked = My.Settings.NovoValorFiguraRitmicas(14)
            CheckBox15.Checked = My.Settings.NovoValorFiguraRitmicas(15)


            CheckBox17.Checked = My.Settings.NovoValorFiguraRitmicas(17)
            CheckBox18.Checked = My.Settings.NovoValorFiguraRitmicas(18)
            CheckBox19.Checked = My.Settings.NovoValorFiguraRitmicas(19)
            CheckBox20.Checked = My.Settings.NovoValorFiguraRitmicas(20)
            CheckBox21.Checked = My.Settings.NovoValorFiguraRitmicas(21)
            CheckBox22.Checked = My.Settings.NovoValorFiguraRitmicas(22)
            CheckBox23.Checked = My.Settings.NovoValorFiguraRitmicas(23)
            CheckBox24.Checked = My.Settings.NovoValorFiguraRitmicas(24)
            CheckBox25.Checked = My.Settings.NovoValorFiguraRitmicas(25)
            CheckBox26.Checked = My.Settings.NovoValorFiguraRitmicas(26)
            CheckBox27.Checked = My.Settings.NovoValorFiguraRitmicas(27)
            CheckBox28.Checked = My.Settings.NovoValorFiguraRitmicas(28)
            CheckBox29.Checked = My.Settings.NovoValorFiguraRitmicas(29)
            CheckBox30.Checked = My.Settings.NovoValorFiguraRitmicas(30)
            CheckBox31.Checked = My.Settings.NovoValorFiguraRitmicas(31)


            NumericUpDown1.Value = My.Settings.NovoValorTempoInicial
            MãoDireita.Checked = My.Settings.NovoValorMãoDireita
            MãoEsquerda.Checked = My.Settings.NovoValorMãoEsquerda

            PontoAumentoME.Checked = My.Settings.NovoValorPontoAumentoME
            PontoAumentoMD.Checked = My.Settings.NovoValorPontoAumentoMD
            DuploPontoAumentoME.Checked = My.Settings.NovoValorDuploPontoAumentoME
            DuploPontoAumentoMD.Checked = My.Settings.NovoValorDuploPontoAumentoMD
            TriploPontoAumentoME.Checked = My.Settings.NovoValorTriploPontoAumentoME
            TriploPontoAumentoMD.Checked = My.Settings.NovoValorTriploPontoAumentoMD

            LigadurasMD.Checked = My.Settings.NovoValorLigadurasMD
            LigadurasME.Checked = My.Settings.NovoValorLigadurasME

            SwingMD.Checked = My.Settings.NovoValorSwingMD
            SwingME.Checked = My.Settings.NovoValorSwingME

            DinamicaME.Checked = My.Settings.NovoValorDinamicaME
            DinamicaMD.Checked = My.Settings.NovoValorDinamicaMD
            PercentualDinamicaME.Value = My.Settings.NovoValorPercentualDinamicaME
            PercentualDinamicaMD.Value = My.Settings.NovoValorPercentualDinamicaMD
            PercentualLigadurasMD.Value = My.Settings.NovoValorPercentualLigadurasMD
            PercentualLigadurasME.Value = My.Settings.NovoValorPercentualLigadurasME

            LocalizaçãoCompassoAtual.Checked = My.Settings.NovoValorLocalizaçãoCompassoAtual
            LocalizaçãoSubdivisaoCompassoAtual.Checked = My.Settings.NovoValorLocalizaçãoSubdivisãoCompassoAtual
            LocalizaçãoMicroSubdivisão.Checked = My.Settings.NovoValorLocalizaçãoMicroSubdivisãoCompassoAtual


            Compasso2_2.Checked = My.Settings.NovoValorCompassos(0)
            Compasso3_2.Checked = My.Settings.NovoValorCompassos(1)
            Compasso4_2.Checked = My.Settings.NovoValorCompassos(2)
            Compasso2_4.Checked = My.Settings.NovoValorCompassos(3)
            Compasso3_4.Checked = My.Settings.NovoValorCompassos(4)
            Compasso4_4.Checked = My.Settings.NovoValorCompassos(5)
            Compasso5_4.Checked = My.Settings.NovoValorCompassos(6)
            Compasso6_4.Checked = My.Settings.NovoValorCompassos(7)
            Compasso9_4.Checked = My.Settings.NovoValorCompassos(8)
            Compasso12_4.Checked = My.Settings.NovoValorCompassos(9)
            Compasso3_8.Checked = My.Settings.NovoValorCompassos(10)
            Compasso4_8.Checked = My.Settings.NovoValorCompassos(11)
            Compasso5_8.Checked = My.Settings.NovoValorCompassos(12)
            Compasso6_8.Checked = My.Settings.NovoValorCompassos(13)
            Compasso7_8.Checked = My.Settings.NovoValorCompassos(14)
            Compasso9_8.Checked = My.Settings.NovoValorCompassos(15)
            Compasso12_8.Checked = My.Settings.NovoValorCompassos(16)

            CheckBox42.Checked = My.Settings.NovoValorForçaDinâmica(8)
            CheckBox41.Checked = My.Settings.NovoValorForçaDinâmica(9)
            CheckBox36.Checked = My.Settings.NovoValorForçaDinâmica(10)
            CheckBox37.Checked = My.Settings.NovoValorForçaDinâmica(11)
            CheckBox40.Checked = My.Settings.NovoValorForçaDinâmica(12)
            CheckBox39.Checked = My.Settings.NovoValorForçaDinâmica(13)
            CheckBox38.Checked = My.Settings.NovoValorForçaDinâmica(14)
            CheckBox35.Checked = My.Settings.NovoValorForçaDinâmica(15)

            CheckBox50.Checked = My.Settings.NovoValorForçaDinâmica(0)
            CheckBox49.Checked = My.Settings.NovoValorForçaDinâmica(1)
            CheckBox44.Checked = My.Settings.NovoValorForçaDinâmica(2)
            CheckBox45.Checked = My.Settings.NovoValorForçaDinâmica(3)
            CheckBox48.Checked = My.Settings.NovoValorForçaDinâmica(4)
            CheckBox47.Checked = My.Settings.NovoValorForçaDinâmica(5)
            CheckBox46.Checked = My.Settings.NovoValorForçaDinâmica(6)
            CheckBox43.Checked = My.Settings.NovoValorForçaDinâmica(7)

            Condescendente.Checked = My.Settings.NovoValorCondescendente
            Normal.Checked = My.Settings.NovoValorNormal
            Severa.Checked = My.Settings.NovoValorSevera

            CorME.BackColor = My.Settings.NovaCorME
            CorMD.BackColor = My.Settings.NovaCorMD
            PictureBox82.BackColor = My.Settings.NovaCorLocalizaçãoDosCompassos
            PercentualTransparenciaCor.Value = My.Settings.NovaTransparênciaCorLocalizaçãoDosCompassos

            CheckBox33.Checked = My.Settings.NovoValorTocarSonsMetronomo
            NumericUpDown2.Value = My.Settings.NovoValorVolumeTempoForte
            NumericUpDown3.Value = My.Settings.NovoValorVolumeTempoFraco
            NumericUpDown4.Value = My.Settings.NovoValorVolumeRitmo


            PercentualDinMD()
            PercentualDinME()
            PercentualLigMD()
            PercentualLigME()

            ComboBox1.Items.Clear()
            ComboBox1.Items.Add("27 - High Q")
            ComboBox1.Items.Add("28 - Slap")
            ComboBox1.Items.Add("29 - Scratch Push")
            ComboBox1.Items.Add("30 - Scratch Pull")
            ComboBox1.Items.Add("31 - Sticks")
            ComboBox1.Items.Add("32 - Square Click")
            ComboBox1.Items.Add("33 - Metronome Click")
            ComboBox1.Items.Add("34 - Metronome Bell")
            ComboBox1.Items.Add("35 - Acoustic Bass Drum")
            ComboBox1.Items.Add("36 - Bass Drum 1")
            ComboBox1.Items.Add("37 - Side Stick")
            ComboBox1.Items.Add("38 - Acoustic Snare")
            ComboBox1.Items.Add("39 - Hand Clap")
            ComboBox1.Items.Add("40 - Electric Snare")
            ComboBox1.Items.Add("41 - Low Floor Tom")
            ComboBox1.Items.Add("42 - Closed Hi Hat")
            ComboBox1.Items.Add("43 - High Floor Tom")
            ComboBox1.Items.Add("44 - Pedal Hi Hat")
            ComboBox1.Items.Add("45 - Low Tom")
            ComboBox1.Items.Add("46 - Open Hi Hat")
            ComboBox1.Items.Add("47 - Low-Mid Tom")
            ComboBox1.Items.Add("48 - Hi-Mid Tom")
            ComboBox1.Items.Add("49 - Crash Cymbal 1")
            ComboBox1.Items.Add("50 - High Tom")
            ComboBox1.Items.Add("51 - Ride Cymbal 1")
            ComboBox1.Items.Add("52 - Chinese Cymbal")
            ComboBox1.Items.Add("53 - Ride Bell")
            ComboBox1.Items.Add("54 - Tambourine")
            ComboBox1.Items.Add("55 - Splash Cymbal")
            ComboBox1.Items.Add("56 - Cowbell")
            ComboBox1.Items.Add("57 - Crash Cymbal 2")
            ComboBox1.Items.Add("58 - Vibraslap")
            ComboBox1.Items.Add("59 - Ride Cymbal 2")
            ComboBox1.Items.Add("60 - Hi Bongo")
            ComboBox1.Items.Add("61 - Low Bongo")
            ComboBox1.Items.Add("62 - Mute Hi Conga")
            ComboBox1.Items.Add("63 - Open Hi Conga")
            ComboBox1.Items.Add("64 - Low Conga")
            ComboBox1.Items.Add("65 - High Timbale")
            ComboBox1.Items.Add("66 - Low Timbale")
            ComboBox1.Items.Add("67 - High Agogo")
            ComboBox1.Items.Add("68 - Low Agogo")
            ComboBox1.Items.Add("69 - Cabasa")
            ComboBox1.Items.Add("70 - Maracas")
            ComboBox1.Items.Add("71 - Short Whistle")
            ComboBox1.Items.Add("72 - Long Whistle")
            ComboBox1.Items.Add("73 - Short Guiro")
            ComboBox1.Items.Add("74 - Long Guiro")
            ComboBox1.Items.Add("75 - Claves")
            ComboBox1.Items.Add("76 - Hi Wood Block")
            ComboBox1.Items.Add("77 - Low Wood Block")
            ComboBox1.Items.Add("78 - Mute Cuica")
            ComboBox1.Items.Add("79 - Open Cuica")
            ComboBox1.Items.Add("80 - Mute Triangle")
            ComboBox1.Items.Add("81 - Open Triangle")
            ComboBox1.Items.Add("82 - Shaker")
            ComboBox1.Items.Add("83 - Jingle Bell")
            ComboBox1.Items.Add("84 - Carillon")
            ComboBox1.Items.Add("85 - Castanets")
            ComboBox1.Items.Add("86 - Mute Surdo")
            ComboBox1.Items.Add("87 - Open Surdo")
            ComboBox1.Text = My.Settings.NovoValorMetrônomoTempoForte

            ComboBox2.Items.Clear()
            ComboBox3.Items.Clear()
            For Each item In ComboBox1.Items
                ComboBox2.Items.Add(item)
                ComboBox3.Items.Add(item)
            Next
            ComboBox2.Text = My.Settings.NovoValorMetrônomoTempoFraco
            ComboBox3.Text = My.Settings.NovoValorSomRitmo

            ComboBox4.Items.Clear()
            ComboBox5.Items.Clear()
            ComboBox4.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "Ç", "´", "{", "}", "|", ",", ".", ";", "-", "=", "/", "*", "-", "+", "."})
            ComboBox5.Items.AddRange(New Object() {"1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "Ç", "´", "{", "}", "|", ",", ".", ";", "-", "=", "/", "*", "-", "+", "."})

            ComboBox4.Text = My.Settings.NovoValorTeclaMãoEsquerda
            ComboBox5.Text = My.Settings.NovoValorTeclaMãoDireita

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub PictureBox38_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox38.Paint
        If Compasso2_2.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox38_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox38.Click
        Compasso2_2.Checked = Not Compasso2_2.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox40_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox40.Paint
        If Compasso3_2.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox40_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox40.Click
        Compasso3_2.Checked = Not Compasso3_2.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox41_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox41.Paint
        If Compasso4_2.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox41_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox41.Click
        Compasso4_2.Checked = Not Compasso4_2.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox39_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox39.Paint
        If Compasso2_4.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox39_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox39.Click
        Compasso2_4.Checked = Not Compasso2_4.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox37_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox37.Paint
        If Compasso3_4.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox37_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox37.Click
        Compasso3_4.Checked = Not Compasso3_4.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox36_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox36.Paint
        If Compasso4_4.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox36_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox36.Click
        Compasso4_4.Checked = Not Compasso4_4.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox52_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox52.Paint
        If Compasso5_4.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox52_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox52.Click
        Compasso5_4.Checked = Not Compasso5_4.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox42_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox42.Paint
        If Compasso6_4.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox42_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox42.Click
        Compasso6_4.Checked = Not Compasso6_4.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox43_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox43.Paint
        If Compasso9_4.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox43_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox43.Click
        Compasso9_4.Checked = Not Compasso9_4.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox44_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox44.Paint
        If Compasso12_4.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox44_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox44.Click
        Compasso12_4.Checked = Not Compasso12_4.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox48_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox48.Paint
        If Compasso3_8.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox48_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox48.Click
        Compasso3_8.Checked = Not Compasso3_8.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox49_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox49.Paint
        If Compasso4_8.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox49_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox49.Click
        Compasso4_8.Checked = Not Compasso4_8.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox50_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox50.Paint
        If Compasso5_8.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox50_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox50.Click
        Compasso5_8.Checked = Not Compasso5_8.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox51_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox51.Paint
        If Compasso6_8.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox51_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox51.Click
        Compasso6_8.Checked = Not Compasso6_8.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox45_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox45.Paint
        If Compasso7_8.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox45_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox45.Click
        Compasso7_8.Checked = Not Compasso7_8.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox46_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox46.Paint
        If Compasso9_8.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox46_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox46.Click
        Compasso9_8.Checked = Not Compasso9_8.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox47_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox47.Paint
        If Compasso12_8.Checked Then e.Graphics.FillRectangle(Cor1, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox47_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox47.Click
        Compasso12_8.Checked = Not Compasso12_8.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub ChecarFigurasNecessárias()

        Try

            If Not Compasso2_2.Checked AndAlso Not Compasso3_2.Checked AndAlso _
            Not Compasso4_2.Checked AndAlso Not Compasso2_4.Checked AndAlso _
            Not Compasso3_4.Checked AndAlso Not Compasso4_4.Checked AndAlso _
            Not Compasso5_4.Checked AndAlso Not Compasso6_4.Checked AndAlso _
            Not Compasso9_4.Checked AndAlso Not Compasso12_4.Checked AndAlso _
            Not Compasso3_8.Checked AndAlso Not Compasso4_8.Checked AndAlso _
            Not Compasso5_8.Checked AndAlso Not Compasso6_8.Checked AndAlso _
            Not Compasso7_8.Checked AndAlso Not Compasso9_8.Checked AndAlso _
            Not Compasso12_8.Checked Then Compasso4_4.Checked = True : PictureBox36.Invalidate()

            If Not MãoEsquerda.Checked AndAlso Not MãoDireita.Checked Then MãoDireita.Checked = True : PictureBox34.Invalidate()

            If Compasso3_8.Checked OrElse Compasso5_8.Checked OrElse Compasso7_8.Checked OrElse Compasso9_8.Checked Then
                If MãoEsquerda.Checked Then
                    If Not CheckBox4.Checked AndAlso Not CheckBox10.Checked AndAlso _
                   Not CheckBox5.Checked AndAlso Not CheckBox11.Checked AndAlso _
                   Not CheckBox6.Checked AndAlso Not CheckBox12.Checked Then CheckBox4.Checked = True : PictureBox4.Invalidate()
                End If
                If MãoDireita.Checked Then
                    If Not CheckBox20.Checked AndAlso Not CheckBox26.Checked AndAlso _
                   Not CheckBox21.Checked AndAlso Not CheckBox27.Checked AndAlso _
                   Not CheckBox22.Checked AndAlso Not CheckBox28.Checked Then CheckBox20.Checked = True : PictureBox20.Invalidate()
                End If
            End If

            If Compasso3_4.Checked OrElse Compasso5_4.Checked OrElse Compasso6_8.Checked OrElse Compasso9_4.Checked Then
                If MãoEsquerda.Checked Then
                    If Not CheckBox3.Checked AndAlso Not CheckBox9.Checked AndAlso _
                   Not CheckBox4.Checked AndAlso Not CheckBox10.Checked AndAlso _
                   Not CheckBox5.Checked AndAlso Not CheckBox11.Checked AndAlso _
                   Not CheckBox6.Checked AndAlso Not CheckBox12.Checked Then CheckBox3.Checked = True : PictureBox3.Invalidate()
                End If
                If MãoDireita.Checked Then
                    If Not CheckBox19.Checked AndAlso Not CheckBox25.Checked AndAlso _
                   Not CheckBox20.Checked AndAlso Not CheckBox26.Checked AndAlso _
                   Not CheckBox21.Checked AndAlso Not CheckBox27.Checked AndAlso _
                   Not CheckBox22.Checked AndAlso Not CheckBox28.Checked Then CheckBox19.Checked = True : PictureBox21.Invalidate()
                End If
            End If

            If Compasso3_2.Checked OrElse Compasso2_4.Checked OrElse Compasso4_8.Checked OrElse Compasso12_8.Checked OrElse Compasso6_4.Checked Then
                If MãoEsquerda.Checked Then
                    If Not CheckBox2.Checked AndAlso Not CheckBox8.Checked AndAlso _
                   Not CheckBox3.Checked AndAlso Not CheckBox9.Checked AndAlso _
                   Not CheckBox4.Checked AndAlso Not CheckBox10.Checked AndAlso _
                   Not CheckBox5.Checked AndAlso Not CheckBox11.Checked AndAlso _
                   Not CheckBox6.Checked AndAlso Not CheckBox12.Checked Then CheckBox2.Checked = True : PictureBox2.Invalidate()
                End If
                If MãoDireita.Checked Then
                    If Not CheckBox18.Checked AndAlso Not CheckBox24.Checked AndAlso _
                   Not CheckBox19.Checked AndAlso Not CheckBox25.Checked AndAlso _
                   Not CheckBox20.Checked AndAlso Not CheckBox26.Checked AndAlso _
                   Not CheckBox21.Checked AndAlso Not CheckBox27.Checked AndAlso _
                   Not CheckBox22.Checked AndAlso Not CheckBox28.Checked Then CheckBox18.Checked = True : PictureBox22.Invalidate()
                End If
            End If

            If Compasso2_2.Checked OrElse Compasso4_4.Checked OrElse Compasso4_2.Checked OrElse Compasso12_4.Checked Then
                If MãoEsquerda.Checked Then
                    If Not CheckBox1.Checked AndAlso Not CheckBox7.Checked AndAlso _
                   Not CheckBox2.Checked AndAlso Not CheckBox8.Checked AndAlso _
                   Not CheckBox3.Checked AndAlso Not CheckBox9.Checked AndAlso _
                   Not CheckBox4.Checked AndAlso Not CheckBox10.Checked AndAlso _
                   Not CheckBox5.Checked AndAlso Not CheckBox11.Checked AndAlso _
                   Not CheckBox6.Checked AndAlso Not CheckBox12.Checked Then CheckBox1.Checked = True : PictureBox1.Invalidate()
                End If
                If MãoDireita.Checked Then
                    If Not CheckBox17.Checked AndAlso Not CheckBox23.Checked AndAlso _
                   Not CheckBox18.Checked AndAlso Not CheckBox24.Checked AndAlso _
                   Not CheckBox19.Checked AndAlso Not CheckBox25.Checked AndAlso _
                   Not CheckBox20.Checked AndAlso Not CheckBox26.Checked AndAlso _
                   Not CheckBox21.Checked AndAlso Not CheckBox27.Checked AndAlso _
                   Not CheckBox22.Checked AndAlso Not CheckBox28.Checked Then CheckBox17.Checked = True : PictureBox23.Invalidate()
                End If
            End If

            If Not CheckBox42.Checked AndAlso Not CheckBox41.Checked AndAlso _
           Not CheckBox36.Checked AndAlso Not CheckBox37.Checked AndAlso _
           Not CheckBox40.Checked AndAlso Not CheckBox39.Checked AndAlso _
           Not CheckBox38.Checked AndAlso Not CheckBox35.Checked Then CheckBox36.Checked = True : PictureBox63.Invalidate()

            If Not CheckBox50.Checked AndAlso Not CheckBox49.Checked AndAlso _
           Not CheckBox44.Checked AndAlso Not CheckBox45.Checked AndAlso _
           Not CheckBox48.Checked AndAlso Not CheckBox47.Checked AndAlso _
           Not CheckBox46.Checked AndAlso Not CheckBox43.Checked Then CheckBox44.Checked = True : PictureBox74.Invalidate()


            SalvaSettings()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub PictureBox1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox1.Paint
        If CheckBox1.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click
        CheckBox1.Checked = Not CheckBox1.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox2_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox2.Paint
        If CheckBox2.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox2.Click
        CheckBox2.Checked = Not CheckBox2.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox3_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox3.Paint
        If CheckBox3.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox3.Click
        CheckBox3.Checked = Not CheckBox3.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox4_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox4.Paint
        If CheckBox4.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox4.Click
        CheckBox4.Checked = Not CheckBox4.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox5_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox5.Paint
        If CheckBox5.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox5.Click
        CheckBox5.Checked = Not CheckBox5.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox6_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox6.Paint
        If CheckBox6.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox6.Click
        CheckBox6.Checked = Not CheckBox6.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox12_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox12.Paint
        If CheckBox7.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox12.Click
        CheckBox7.Checked = Not CheckBox7.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox24_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox24.Paint
        If CheckBox8.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox24.Click
        CheckBox8.Checked = Not CheckBox8.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox10_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox10.Paint
        If CheckBox9.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox10.Click
        CheckBox9.Checked = Not CheckBox9.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox9_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox9.Paint
        If CheckBox10.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox9.Click
        CheckBox10.Checked = Not CheckBox10.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox8_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox8.Paint
        If CheckBox11.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox8.Click
        CheckBox11.Checked = Not CheckBox11.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox7_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox7.Paint
        If CheckBox12.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox7.Click
        CheckBox12.Checked = Not CheckBox12.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox23_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox23.Paint
        If CheckBox17.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox23.Click
        CheckBox17.Checked = Not CheckBox17.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox22_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox22.Paint
        If CheckBox18.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox22.Click
        CheckBox18.Checked = Not CheckBox18.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox21_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox21.Paint
        If CheckBox19.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox21.Click
        CheckBox19.Checked = Not CheckBox19.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox20_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox20.Paint
        If CheckBox20.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox20.Click
        CheckBox20.Checked = Not CheckBox20.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox19_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox19.Paint
        If CheckBox21.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox19.Click
        CheckBox21.Checked = Not CheckBox21.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox18_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox18.Paint
        If CheckBox22.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox18.Click
        CheckBox22.Checked = Not CheckBox22.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox11_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox11.Paint
        If CheckBox23.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox11.Click
        CheckBox23.Checked = Not CheckBox23.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox25_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox25.Paint
        If CheckBox24.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox25.Click
        CheckBox24.Checked = Not CheckBox24.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox29_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox29.Paint
        If CheckBox25.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox29.Click
        CheckBox25.Checked = Not CheckBox25.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox28_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox28.Paint
        If CheckBox26.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox28.Click
        CheckBox26.Checked = Not CheckBox26.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox27_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox27.Paint
        If CheckBox27.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox27.Click
        CheckBox27.Checked = Not CheckBox27.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox26_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox26.Paint
        If CheckBox28.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox26.Click
        CheckBox28.Checked = Not CheckBox28.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox13_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox13.Paint
        If CheckBox13.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox13.Click
        CheckBox13.Checked = Not CheckBox13.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox14_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox14.Paint
        If CheckBox14.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox14.Click
        CheckBox14.Checked = Not CheckBox14.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox15_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox15.Paint
        If CheckBox15.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox15.Click
        CheckBox15.Checked = Not CheckBox15.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox16_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox16.Paint
        If CheckBox16.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox16.Click
        CheckBox16.Checked = Not CheckBox16.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox33_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox33.Paint
        If CheckBox29.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox33_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox33.Click
        CheckBox29.Checked = Not CheckBox29.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox32_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox32.Paint
        If CheckBox30.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox32_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox32.Click
        CheckBox30.Checked = Not CheckBox30.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox31_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox31.Paint
        If CheckBox31.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox31.Click
        CheckBox31.Checked = Not CheckBox31.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox30_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox30.Paint
        If CheckBox32.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox30.Click
        CheckBox32.Checked = Not CheckBox32.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox35_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox35.Paint
        If MãoEsquerda.Checked Then e.Graphics.FillRectangle(Cor3, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox35_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox35.Click
        MãoEsquerda.Checked = Not MãoEsquerda.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox34_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox34.Paint
        If MãoDireita.Checked Then e.Graphics.FillRectangle(Cor3, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox34_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox34.Click
        MãoDireita.Checked = Not MãoDireita.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub DinamicaME_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DinamicaME.CheckedChanged, Me.Load
        If DinamicaME.Checked Then
            PercentualDinamicaME.Enabled = True
        Else
            PercentualDinamicaME.Enabled = False
        End If
    End Sub

    Private Sub DinamicaMD_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DinamicaMD.CheckedChanged, Me.Load
        If DinamicaMD.Checked Then
            PercentualDinamicaMD.Enabled = True
        Else
            PercentualDinamicaMD.Enabled = False
        End If
    End Sub

    Private Sub PercentualDinMD()
        Label20.Text = CInt(PercentualDinamicaMD.Value / 60 * 100) & "%"
    End Sub

    Private Sub PercentualDinME()
        Label19.Text = CInt(PercentualDinamicaME.Value / 60 * 100) & "%"
    End Sub

    Private Sub PercentualLigMD()
        Label21.Text = CInt(PercentualLigadurasMD.Value / 60 * 100) & "%"
    End Sub

    Private Sub PercentualLigME()
        Label22.Text = CInt(PercentualLigadurasME.Value / 60 * 100) & "%"
    End Sub

    Private Sub PercentualDinamicaMD_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles PercentualDinamicaMD.Scroll
        PercentualDinMD()
    End Sub

    Private Sub PercentualDinamicaME_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles PercentualDinamicaME.Scroll
        PercentualDinME()
    End Sub

    Private Sub PercentualLigadurasMD_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles PercentualLigadurasMD.Scroll
        PercentualLigMD()
    End Sub

    Private Sub PercentualLigadurasME_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles PercentualLigadurasME.Scroll
        PercentualLigME()
    End Sub

    Private Sub PictureBox54_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox54.Paint
        If LigadurasMD.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox54_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox54.Click
        LigadurasMD.Checked = Not LigadurasMD.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox53_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox53.Paint
        If LigadurasME.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox53_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox53.Click
        LigadurasME.Checked = Not LigadurasME.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub LigadurasMD_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LigadurasMD.CheckedChanged, Me.Load
        If LigadurasMD.Checked Then
            PercentualLigadurasMD.Enabled = True
        Else
            PercentualLigadurasMD.Enabled = False
        End If
    End Sub

    Private Sub LigadurasME_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LigadurasME.CheckedChanged, Me.Load
        If LigadurasME.Checked Then
            PercentualLigadurasME.Enabled = True
        Else
            PercentualLigadurasME.Enabled = False
        End If
    End Sub

    Private Sub PictureBox55_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox55.Paint
        If PontoAumentoME.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox55_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox55.Click
        PontoAumentoME.Checked = Not PontoAumentoME.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox56_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox56.Paint
        If DuploPontoAumentoME.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox56_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox56.Click
        DuploPontoAumentoME.Checked = Not DuploPontoAumentoME.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox57_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox57.Paint
        If TriploPontoAumentoME.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox57_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox57.Click
        TriploPontoAumentoME.Checked = Not TriploPontoAumentoME.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox60_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox60.Paint
        If PontoAumentoMD.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox60_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox60.Click
        PontoAumentoMD.Checked = Not PontoAumentoMD.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox59_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox59.Paint
        If DuploPontoAumentoMD.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox59_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox59.Click
        DuploPontoAumentoMD.Checked = Not DuploPontoAumentoMD.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox58_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox58.Paint
        If TriploPontoAumentoMD.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox58_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox58.Click
        TriploPontoAumentoMD.Checked = Not TriploPontoAumentoMD.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox61_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox61.Paint
        If CheckBox42.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox61_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox61.Click
        CheckBox42.Checked = Not CheckBox42.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox62_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox62.Paint
        If CheckBox41.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox62_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox62.Click
        CheckBox41.Checked = Not CheckBox41.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox63_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox63.Paint
        If CheckBox36.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox63_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox63.Click
        CheckBox36.Checked = Not CheckBox36.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox64_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox64.Paint
        If CheckBox37.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox64_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox64.Click
        CheckBox37.Checked = Not CheckBox37.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox65_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox65.Paint
        If CheckBox40.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox65_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox65.Click
        CheckBox40.Checked = Not CheckBox40.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox66_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox66.Paint
        If CheckBox39.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox66_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox66.Click
        CheckBox39.Checked = Not CheckBox39.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox67_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox67.Paint
        If CheckBox38.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox67_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox67.Click
        CheckBox38.Checked = Not CheckBox38.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox68_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox68.Paint
        If CheckBox35.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox68_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox68.Click
        CheckBox35.Checked = Not CheckBox35.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox76_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox76.Paint
        If CheckBox50.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox76_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox76.Click
        CheckBox50.Checked = Not CheckBox50.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox75_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox75.Paint
        If CheckBox49.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox75_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox75.Click
        CheckBox49.Checked = Not CheckBox49.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox74_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox74.Paint
        If CheckBox44.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox74_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox74.Click
        CheckBox44.Checked = Not CheckBox44.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox73_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox73.Paint
        If CheckBox45.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox73_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox73.Click
        CheckBox45.Checked = Not CheckBox45.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox72_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox72.Paint
        If CheckBox48.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox72_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox72.Click
        CheckBox48.Checked = Not CheckBox48.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox71_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox71.Paint
        If CheckBox47.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox71_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox71.Click
        CheckBox47.Checked = Not CheckBox47.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox70_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox70.Paint
        If CheckBox46.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox70_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox70.Click
        CheckBox46.Checked = Not CheckBox46.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox69_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox69.Paint
        If CheckBox43.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox69_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox69.Click
        CheckBox43.Checked = Not CheckBox43.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub Condescendente_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Condescendente.CheckedChanged
        If Condescendente.Checked Then
            Normal.Checked = False
            Severa.Checked = False
        End If
    End Sub

    Private Sub Normal_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Normal.CheckedChanged
        If Normal.Checked Then
            Condescendente.Checked = False
            Severa.Checked = False
        End If
    End Sub

    Private Sub Severa_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Severa.CheckedChanged
        If Severa.Checked Then
            Normal.Checked = False
            Condescendente.Checked = False
        End If
    End Sub

    Private Sub CorME_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CorME.Click
        ColorDialog1.Color = My.Settings.NovaCorME
        ColorDialog1.ShowDialog()
        My.Settings.NovaCorME = ColorDialog1.Color
        CorME.BackColor = ColorDialog1.Color
    End Sub

    Private Sub CorMD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CorMD.Click
        ColorDialog1.Color = My.Settings.NovaCorMD
        ColorDialog1.ShowDialog()
        My.Settings.NovaCorMD = ColorDialog1.Color
        CorMD.BackColor = ColorDialog1.Color
    End Sub

    Private Sub PictureBox77_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox77.Paint
        If SwingME.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox77_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox77.Click
        SwingME.Checked = Not SwingME.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox78_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox78.Paint
        If SwingMD.Checked Then e.Graphics.FillRectangle(Cor2, 0, 0, sender.Width, sender.Height)
    End Sub

    Private Sub PictureBox78_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox78.Click
        SwingMD.Checked = Not SwingMD.Checked
        sender.Invalidate()
        ChecarFigurasNecessárias()
    End Sub

    Private Sub PictureBox79_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox79.Paint
        If LocalizaçãoCompassoAtual.Checked Then e.Graphics.FillRectangle(Cor4, 0, 0, sender.Width, 10)
        e.Graphics.FillRectangle(New SolidBrush(Color.FromArgb(PercentualTransparenciaCor.Value, PictureBox82.BackColor)), 0, 30, sender.Width, 4)
    End Sub

    Private Sub PictureBox79_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox79.Click
        LocalizaçãoCompassoAtual.Checked = Not LocalizaçãoCompassoAtual.Checked
        sender.Invalidate()
    End Sub

    Private Sub PictureBox80_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox80.Paint
        If LocalizaçãoSubdivisaoCompassoAtual.Checked Then e.Graphics.FillRectangle(Cor4, 0, 0, sender.Width, 10)
        e.Graphics.FillRectangle(New SolidBrush(Color.FromArgb(PercentualTransparenciaCor.Value, PictureBox82.BackColor)), 0, 30, sender.Width, 4)
        e.Graphics.FillRectangle(New SolidBrush(Color.FromArgb(PercentualTransparenciaCor.Value, PictureBox82.BackColor)), 10, 30, 36, 4)
    End Sub

    Private Sub PictureBox80_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox80.Click
        LocalizaçãoSubdivisaoCompassoAtual.Checked = Not LocalizaçãoSubdivisaoCompassoAtual.Checked
        sender.Invalidate()
    End Sub

    Private Sub PictureBox81_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox81.Paint
        If LocalizaçãoMicroSubdivisão.Checked Then e.Graphics.FillRectangle(Cor4, 0, 0, sender.Width, 10)
        e.Graphics.FillRectangle(New SolidBrush(Color.FromArgb(PercentualTransparenciaCor.Value, PictureBox82.BackColor)), 0, 30, sender.Width, 4)
        e.Graphics.FillRectangle(New SolidBrush(Color.FromArgb(PercentualTransparenciaCor.Value, PictureBox82.BackColor)), 10, 30, 36, 4)
        e.Graphics.DrawLine(Pens.Black, 15, 30, 15, 33)
    End Sub

    Private Sub PictureBox81_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox81.Click
        LocalizaçãoMicroSubdivisão.Checked = Not LocalizaçãoMicroSubdivisão.Checked
        sender.Invalidate()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            If ComboBox1.Focus = True Then MidiPlayer.Play(New NoteOn(0, CType(ComboBox1.Text.Substring(0, 2), GeneralMidiPercussion), NumericUpDown2.Value))
        Catch ex As Exception
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não
        End Try
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        Try
            If ComboBox2.Focus = True Then MidiPlayer.Play(New NoteOn(0, CType(ComboBox2.Text.Substring(0, 2), GeneralMidiPercussion), NumericUpDown3.Value))
        Catch ex As Exception
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não
        End Try
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        Try
            If ComboBox3.Focus = True Then MidiPlayer.Play(New NoteOn(0, CType(ComboBox3.Text.Substring(0, 2), GeneralMidiPercussion), NumericUpDown4.Value))
        Catch ex As Exception
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não
        End Try
    End Sub

    Private Sub NumericUpDown2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown2.ValueChanged
        Try
            If NumericUpDown2.Focus = True Then MidiPlayer.Play(New NoteOn(0, CType(ComboBox1.Text.Substring(0, 2), GeneralMidiPercussion), NumericUpDown2.Value))
        Catch ex As Exception
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não
        End Try
    End Sub

    Private Sub NumericUpDown3_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown3.ValueChanged
        Try
            If NumericUpDown3.Focus = True Then MidiPlayer.Play(New NoteOn(0, CType(ComboBox2.Text.Substring(0, 2), GeneralMidiPercussion), NumericUpDown3.Value))
        Catch ex As Exception
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não
        End Try
    End Sub

    Private Sub NumericUpDown4_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown4.ValueChanged
        Try
            If NumericUpDown4.Focus = True Then MidiPlayer.Play(New NoteOn(0, CType(ComboBox3.Text.Substring(0, 2), GeneralMidiPercussion), NumericUpDown4.Value))
        Catch ex As Exception
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não
        End Try
    End Sub

    Private Sub PictureBox82_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox82.Click
        ColorDialog1.Color = My.Settings.NovaCorLocalizaçãoDosCompassos
        ColorDialog1.ShowDialog()
        My.Settings.NovaCorLocalizaçãoDosCompassos = ColorDialog1.Color
        PictureBox82.BackColor = ColorDialog1.Color
    End Sub

    Private Sub PictureBox82_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox82.MouseUp
        PictureBox79.Refresh()
        PictureBox80.Refresh()
        PictureBox81.Refresh()
    End Sub

    Private Sub PercentualTransparenciaCor_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles PercentualTransparenciaCor.Scroll
        PictureBox79.Refresh()
        PictureBox80.Refresh()
        PictureBox81.Refresh()
    End Sub

    Private Sub ComboBox4_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedValueChanged, ComboBox5.SelectedValueChanged
        Try
            If ComboBox4.Text = ComboBox5.Text Then
                MsgBox("Você não poderá selecionar a mesma tecla para a mão esquerda e direita", MsgBoxStyle.Information)
                ComboBox4.Text = My.Settings.NovoValorTeclaMãoEsquerda
                ComboBox5.Text = My.Settings.NovoValorTeclaMãoDireita
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não
        End Try
    End Sub
End Class