Option Strict On
Option Explicit On
Imports System.Drawing.Bitmap
Imports System.Xml
Imports System.Windows.Forms.DataVisualization.Charting

Public Class MemoNotes
    Dim a, dd, b, d, f, h, k, kk, kkFinal(3), ii, n, w, ww, bb, ee, yyy, vv, zz, GeraNotaTransposta, NumeroDeNotas(1), AjusteNotas(2), ValorTotalDeNotas, Nota(3), _
        SetaPautaTop, SetaPautaLeft, SetaTecladoLeft, ValorLeft(7), ajusteDivisaoOitavas, AjusteArmaduraClaves, AjusteNomeNotas, AjusteDasNotasNaClave, AjusteDasNotasNaClaveFinal(3), _
        AjusteTrackBar(1), spin, valorOitava(1), FundamentalDoAcorde, _
        AcordeSustenidoOuBemol, FamiliaDoAcorde, AjusteLinhaSuplementar(1), AjusteLeftFigura, QtdeLoopsNotas, TeclaControlPressionada, NotaMidi, QtdeLoopingsPermitidos, ToneIndex2, NotaCantadaAjustadaParaPauta As Integer
    Private lista As New List(Of Keys)
    Dim CoresArmadura, CaptaçãoSom, SinalDetectado, EventoPaintGeradoPelo_ExecutaLoopNotasAleatórias As Boolean
    Dim e, m, mFinal(3), aa, NomeClave, TempoMedio, HoraInicio, HoraFim, ValorPercentual, LetraCifra, LetraCifraFinal(3), RetanguloBranco, NomeDoMenu, NomeCifra(1), _
        NomeTecla(3), NomeTeclaFinal, Ajuste_NomeNotas(6), CifraFundamentalDoAcorde, NomeAcorde, NotaCantadaArmazenada, NotaCantadaArmazenada2 As String
    Dim newPoint As New Point()
    'Dim Img As Bitmap = My.Resources.Pauta
    Dim DivisaoOitavas, CorretoIncorreto, NomeDaNota(3), NomedaNota2(3), ToolTipNota(7), ImagemFundo(1) As Bitmap
    Private Thread(1) As Thread
    Dim PosiçãoTecla(87, 0) As Point
    Dim Cronometro As New Stopwatch
    Dim CurrentFrequency, closestFrequency, MargemErroInferior, MargemErroSuperior As Double
    Dim sliderBrush1 As Brush, sliderBrush2 As Brush

    Dim KeyCol As New System.Collections.Generic.List(Of UserControl)


    ' Coordonnées de départ pour déplacer la Form
    Private X_Piano As Short
    Private Y_Piano As Short
    Private Oct As Short
    Private Xpose As Short = -12

    Dim canal As Byte ' canal midi
    Dim ccolor(16) As Color ' couleur selon canal
    Dim hMidiIn As Integer
    ' Délégué pour le callback Midi In
    Private DelgMidiIn As New MidiDelegate(AddressOf MidiInProc)
    ' Permet de transmettre les paramètre à la Win Form
    Delegate Sub SetParamCallback(ByVal [Param] As Byte, ByVal [KeyVelocity] As Byte)
    ' vpx
    'Delegate Sub SetParamCallback(ByVal [Param] As Byte, ByVal [canal] As Byte)
    Dim DelgParamON As New SetParamCallback(AddressOf TouchOn)
    Dim DelgParamOff As New SetParamCallback(AddressOf TouchOff)

    Private Sub MemoNotes_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        Try
            WinMM.midiInReset(hMidiIn) 'se não resetar o midiinclose não funcionará
            WinMM.midiInClose(hMidiIn)

            MidiPlayer.CloseMidi()

            If IsListenning Then
                StopListenning()
                UpdateListenStopButtons()
            End If

            IndicadorDeFrequencia.Close()

        Catch ex As Exception
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

            If Not RadioButton1.Checked Then RadioButton1.Checked = True

            SalvaSettings()
            'CloseClips()
        End Try

    End Sub



    Public Sub SalvaSettings()

        Try

            My.Settings.NovoValorTimers(0) = Minutos.Text
            My.Settings.NovoValorTimers(1) = Segundos.Text
            My.Settings.NovoValorTimers(2) = Decimos.Text

            My.Settings.NovoValorCheckBox3 = CheckBox3.Checked
            My.Settings.NovoValorCheckBox4 = CheckBox4.Checked
            My.Settings.NovoValorCheckBox11 = CheckBox11.Checked
            My.Settings.NovoValorCheckBox13 = DóPrimeiraLinhaToolStripMenuItem1.Checked
            My.Settings.NovoValorCheckBox14 = DóSegundaLinhaToolStripMenuItem1.Checked
            My.Settings.NovoValorCheckBox15 = DóTerceiraLinhaToolStripMenuItem1.Checked
            My.Settings.NovoValorCheckBox16 = DóQuartaLinhaToolStripMenuItem1.Checked
            My.Settings.NovoValorCheckBox18 = DóQuintaLinhaToolStripMenuItem1.Checked
            My.Settings.NovoValorCheckBox17 = FáTerceiraLinhaToolStripMenuItem1.Checked
            My.Settings.NovoValorCheckBox19 = FáQuintaLinhaToolStripMenuItem1.Checked
            My.Settings.NovoValorCheckBox24 = SolPrimeiraLinhaToolStripMenuItem1.Checked
            My.Settings.NovoValorCheckBox26 = CheckBox26.Checked

            My.Settings.NovoValorVisual = Visual.Checked
            My.Settings.NovoValorSonoro = Sonoro.Checked

            My.Settings.NovoValorTextBox1 = TextBox1.Text
            My.Settings.NovoValorQtdeNotas = QtdeNotas.Value

            My.Settings.NovoValorNotaExata = NotaExata.Checked
            My.Settings.NovoValorCifra = NotaGenérica.Checked

            My.Settings.ValorNovoTrackBar1 = TrackBar1.Value
            My.Settings.ValorNovoTrackBar2 = TrackBar2.Value
            My.Settings.ValorNovoTrackBar4 = TrackBar4.Value
            My.Settings.ValorNovoTrackBar5 = TrackBar5.Value
            My.Settings.ValorNovoAcidentes = Acidentes.Value

            My.Settings.NovoValorInstrumentoMusical(0) = ComboBox1.Text

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub SegundaThreadCode()
        Try

            Fechar.Image = My.Resources.Fechar_Inativo
            Minimizar.Image = My.Resources.Minimizar_Inativo
            DefineTecladoInicial()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub DefineTecladoInicial()

        Try

            For xTecla As Integer = 1 To 88
                Dim cControl() As Control = Me.Controls.Find("P" & xTecla.ToString, True)
                If cControl.Length > 0 Then
                    Dim pb As PictureBox = DirectCast(cControl(0), PictureBox)
                    If pb.Width = 13 Then
                        pb.Image = My.Resources.Tecla_Preta
                    Else
                        pb.Image = My.Resources.Tecla_Branca
                    End If
                    PosiçãoTecla(xTecla - 1, 0) = pb.Location
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Public Sub GerarImagemDeFundo()

        Try

            Dim FaceBit As New Bitmap(Me.Width, Me.Height)
            Dim gr As Graphics = Graphics.FromImage(FaceBit)
            gr.DrawImage(My.Resources.Pauta, 0, 0)



            AjusteArmaduraClaves = 0

            NomeClave = ""
            If CheckBox3.Checked OrElse CheckBox4.Checked OrElse ToolStripMenuItem14.Checked OrElse ToolStripMenuItem13.Checked Then
                gr.DrawImage(My.Resources.AbrangenciaNotas1, 248, 138, 74, 434)


                If My.Settings.NovoValorLinhaSuplementar = False Then
                    gr.DrawImage(My.Resources.ClaveSolFá, 18, 169, 221, 364)
                Else
                    gr.DrawImage(My.Resources.ClaveSolFá___2, 18, 169, 221, 364)
                End If

                If CheckBox3.Checked AndAlso CheckBox4.Checked Then
                    NomeClave = "Sol e Fá (Piano)"
                ElseIf CheckBox3.Checked AndAlso Not CheckBox4.Checked Then
                    NomeClave = "Sol (Piano)"
                ElseIf Not CheckBox3.Checked AndAlso CheckBox4.Checked Then
                    NomeClave = "Fá (Piano)"
                End If

            ElseIf Not CheckBox3.Checked AndAlso Not CheckBox4.Checked AndAlso Not ToolStripMenuItem14.Checked AndAlso Not ToolStripMenuItem13.Checked Then
                gr.DrawImage(My.Resources.AbrangenciaNotas2, 248, 138, 74, 248)

                If My.Settings.NovoValorLinhaSuplementar = False Then
                    gr.DrawImage(My.Resources.PautaClavesDeDó, 36, 171, 202, 171)
                Else
                    gr.DrawImage(My.Resources.PautaClavesDeDó___2, 36, 171, 202, 171)
                End If


                If DóPrimeiraLinhaToolStripMenuItem1.Checked OrElse DóPrimeiraLinhaToolStripMenuItem3.Checked Then
                    gr.DrawImage(My.Resources.ClaveDó, 39, 272, 51, 70)
                    AjusteArmaduraClaves = 25
                    NomeClave = "Dó na Primeira Linha"
                ElseIf DóSegundaLinhaToolStripMenuItem1.Checked OrElse ToolStripMenuItem16.Checked Then
                    gr.DrawImage(My.Resources.ClaveDó, 39, 262, 51, 70)
                    AjusteArmaduraClaves = 15
                    NomeClave = "Dó na Segunda Linha"
                ElseIf DóTerceiraLinhaToolStripMenuItem1.Checked OrElse ToolStripMenuItem17.Checked Then
                    gr.DrawImage(My.Resources.ClaveDó, 39, 252, 51, 70)
                    AjusteArmaduraClaves = 5
                    NomeClave = "Dó na Terceira Linha"
                ElseIf DóQuartaLinhaToolStripMenuItem1.Checked OrElse ToolStripMenuItem18.Checked Then
                    gr.DrawImage(My.Resources.ClaveDó, 39, 242, 51, 70)
                    AjusteArmaduraClaves = -5
                    NomeClave = "Dó na Quarta Linha"
                ElseIf FáTerceiraLinhaToolStripMenuItem1.Checked OrElse ToolStripMenuItem20.Checked Then
                    gr.DrawImage(My.Resources.ClaveFa2, 39, 265, 52, 60)
                    AjusteArmaduraClaves = 20
                    NomeClave = "Fá na Terceira Linha"
                ElseIf DóQuintaLinhaToolStripMenuItem1.Checked OrElse ToolStripMenuItem19.Checked Then
                    gr.DrawImage(My.Resources.ClaveDó, 39, 232, 51, 70)
                    AjusteArmaduraClaves = 20
                    NomeClave = "Dó na Quinta Linha"
                ElseIf FáQuintaLinhaToolStripMenuItem1.Checked OrElse ToolStripMenuItem21.Checked Then
                    gr.DrawImage(My.Resources.ClaveFa2, 39, 245, 52, 60)
                    'não precisa AjusteArmaduraClaves = 0
                    NomeClave = "Fá na Quinta Linha"
                ElseIf SolPrimeiraLinhaToolStripMenuItem1.Checked OrElse ToolStripMenuItem22.Checked Then
                    gr.DrawImage(My.Resources.ClaveSol2, 39, 235, 52, 113)
                    AjusteArmaduraClaves = 10
                    NomeClave = "Sol na Primeira Linha"
                End If
            End If


            ExibirTrackBars()



            If RadioButton1.Checked Then

                CoresArmadura = My.Settings.NovoValorCoresArmadura
                ArmadurasDeClave(gr)

                If My.Settings.NovoValorExibirNomeDaNota = True Then

                    If My.Settings.NovoValorExibirNumeraçãoDasNotas = True Then gr.DrawImage(My.Resources.NumeraçãoNotas, 89, 714, 884, 7)

                Else
                    SemNota(gr)
                End If

                If My.Settings.NovoValorIdentificarPosiçãoDóCentral = True Then
                    If My.Settings.NovoValorExibirNomeDaNota = True AndAlso My.Settings.NovoValorExibirNumeraçãoDasNotas = True Then
                        gr.DrawImage(My.Resources.Dó_Central, 458, 721, 62, 38)
                    Else
                        gr.DrawImage(My.Resources.Dó_Central, 458, 710, 62, 38)
                    End If
                End If

                If My.Settings.NovoValorIdentificarPosiçãoNotasSOLeFÁ = True Then
                    If My.Settings.NovoValorExibirNomeDaNota = True AndAlso My.Settings.NovoValorExibirNumeraçãoDasNotas = True Then
                        gr.DrawImage(My.Resources.Fá_no_Teclado, 390, 721, 62, 38)
                        gr.DrawImage(My.Resources.Sol_no_Teclado, 526, 721, 62, 38)
                    Else
                        gr.DrawImage(My.Resources.Fá_no_Teclado, 390, 710, 62, 38)
                        gr.DrawImage(My.Resources.Sol_no_Teclado, 526, 710, 62, 38)
                    End If
                End If


                IdentificarPosiçãoSetasDó(gr)

                IdentificarLinhasGuias(gr)

            Else 'radiobutton2 is checked – está em modo jogo

                CoresArmadura = False
                ArmadurasDeClave(gr)

                SemNota(gr)

            End If


            If My.Settings.NovoValorExibirNumeraçãoEspaçosLinhasSuplementares = True Then

                RotinaAjusteDivisaoOitavas()

                gr.DrawImage(My.Resources.Numeração_linhas_suplementares, 78, 164, 22, 185)

                If CheckBox4.Checked OrElse CheckBox3.Checked Then gr.DrawImage(My.Resources.Numeração_linhas_suplementares2, 78, 384, 22, 154)

                If My.Settings.NovoValorExibirNumeraçãoDasNotas = True Then gr.DrawImage(DivisaoOitavas, 101, 170 + ajusteDivisaoOitavas, 6, 365)

            Else

                RotinaAjusteDivisaoOitavas()

                If My.Settings.NovoValorExibirNumeraçãoDasNotas = True Then gr.DrawImage(DivisaoOitavas, 92, 170 + ajusteDivisaoOitavas, 6, 365)

            End If


            ImagemFundo(0) = FaceBit

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Public Sub GerarImagemDeFundo2()

        Try

            If ImagemFundo(0) IsNot Nothing Then
                ImagemFundo(1) = New Bitmap(ImagemFundo(0))
                Dim gr As Graphics = Graphics.FromImage(ImagemFundo(1))

                If RadioButton1.Checked Then

                    QtdeLoopsNotas = 1
                    LoopNotas(gr)

                    If My.Settings.NovoValorExibirNomeDaNota = True Then
                        gr.SmoothingMode = SmoothingMode.AntiAlias
                        Dim Fonte As New Font("Arial", 12, FontStyle.Regular, GraphicsUnit.Pixel)
                        Dim MedidasTexto As SizeF = gr.MeasureString(LetraCifra, Fonte)
                        Dim MedidasTexto2 As SizeF = gr.MeasureString("Posição no teclado: " & o, Fonte)
                        Dim PosicaoX As Integer = CInt((1032 / 2) - (MedidasTexto.Width / 2))
                        Dim PosicaoY As Integer = 402
                        Dim PosicaoX2 As Integer = CInt((1032 / 2) - (MedidasTexto2.Width / 2))
                        Dim PosicaoY2 As Integer = 414
                        Dim Branco As SolidBrush = New SolidBrush(Color.FromArgb(255, 255, 250))

                        gr.DrawString(LetraCifra, Fonte, Branco, PosicaoX, PosicaoY)
                        gr.DrawString("Posição no teclado: " & o, Fonte, Branco, PosicaoX2, PosicaoY2)
                        gr.SmoothingMode = SmoothingMode.None

                        gr.DrawImage(NomedaNota2(0), 413, 270, 200, 150)

                    End If


                    DesenharNomeNota0(gr)

                    If Not CifrasDeAcordes.Checked Then ExibiçãoSustenidosBemóis1(gr)

                    ExibiçãoDosIntervalos(gr)


                Else 'radiobutton2 is checked – está em modo jogo


                    QtdeLoopsNotas = CInt(TextBox1.Text)
                    LoopNotas(gr)

                    If CifrasDeAcordes.Checked Then
                        Dim Branco As SolidBrush = New SolidBrush(Color.FromArgb(255, 255, 250))
                        Dim Fonte4 As New Font("Arial", 40, FontStyle.Bold, GraphicsUnit.Pixel)
                        Dim MedidasTexto As SizeF = gr.MeasureString(NomeAcorde, Fonte4)
                        Dim PosicaoX As Integer = CInt((Me.Width / 2) - (MedidasTexto.Width / 2))
                        Dim PosicaoY As Integer = 330
                        gr.DrawString(NomeAcorde, Fonte4, Branco, PosicaoX - 15, PosicaoY)

                    Else

                        ExibiçãoSustenidosBemóis1(gr)
                        ExibiçãoSustenidosBemóis2(gr)

                    End If

                    If Visual.Checked Then DesenharNomeNota0(gr)

                    'Desenha as demais notas, se necessário
                    If CInt(TextBox1.Text) > 1 Then gr.DrawImage(NomeDaNota(1), 141 - AjusteLeftFigura, (AjusteNotas(0) - (Nota(1) * 5)) - AjusteDasNotasNaClaveFinal(1), AjusteNotas(1), AjusteNotas(2))
                    If CInt(TextBox1.Text) > 2 Then gr.DrawImage(NomeDaNota(2), 141 - AjusteLeftFigura, (AjusteNotas(0) - (Nota(2) * 5)) - AjusteDasNotasNaClaveFinal(2), AjusteNotas(1), AjusteNotas(2))
                    If CInt(TextBox1.Text) > 3 Then gr.DrawImage(NomeDaNota(3), 141 - AjusteLeftFigura, (AjusteNotas(0) - (Nota(3) * 5)) - AjusteDasNotasNaClaveFinal(3), AjusteNotas(1), AjusteNotas(2))


                End If




            End If


            AtualizaRegiões()



        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub Claves(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint

        Try
            e.Graphics.DrawImage(ImagemFundo(1), 0, 0)
            Dim Branco2 As SolidBrush = New SolidBrush(Color.FromArgb(150, 255, 255, 250))

            If CorretoIncorreto IsNot Nothing Then e.Graphics.DrawImage(CorretoIncorreto, 342, 184, 349, 61)

            If TrackBar1.Value >= 51 AndAlso TrackBar5.Value <= 51 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar, 231, 251, 7, 1)
            If TrackBar1.Value >= 53 AndAlso TrackBar5.Value <= 53 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar, 231, 241, 7, 1)
            If TrackBar1.Value >= 55 AndAlso TrackBar5.Value <= 55 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar, 231, 231, 7, 1)
            If TrackBar1.Value >= 57 AndAlso TrackBar5.Value <= 57 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar, 231, 221, 7, 1)
            If TrackBar1.Value >= 59 AndAlso TrackBar5.Value <= 59 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar, 231, 211, 7, 1)
            If TrackBar1.Value >= 61 AndAlso TrackBar5.Value <= 61 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar, 231, 201, 7, 1)
            If TrackBar1.Value >= 63 AndAlso TrackBar5.Value <= 63 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar, 231, 191, 7, 1)
            If TrackBar1.Value >= 65 AndAlso TrackBar5.Value <= 65 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar, 231, 181, 7, 1)
            If TrackBar1.Value >= 67 AndAlso TrackBar5.Value <= 67 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar, 231, 171, 7, 1)

            If TrackBar5.Value <= 39 AndAlso TrackBar1.Value >= 39 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar, 231, 311, 7, 1)
            If TrackBar5.Value <= 37 AndAlso TrackBar1.Value >= 37 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar, 231, 321, 7, 1)
            If TrackBar5.Value <= 35 AndAlso TrackBar1.Value >= 35 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar, 231, 331, 7, 1)
            If TrackBar5.Value <= 33 AndAlso TrackBar1.Value >= 33 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar, 231, 341, 7, 1)

            If CheckBox4.Checked Then
                If TrackBar4.Value >= 24 AndAlso TrackBar2.Value <= 24 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar, 231, 421, 7, 1)
                If TrackBar4.Value >= 26 AndAlso TrackBar2.Value <= 26 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar, 231, 411, 7, 1)
                If TrackBar4.Value >= 28 AndAlso TrackBar2.Value <= 28 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar, 231, 401, 7, 1)
                If TrackBar4.Value >= 30 AndAlso TrackBar2.Value <= 30 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar, 231, 391, 7, 1)

                If TrackBar2.Value <= 12 AndAlso TrackBar4.Value >= 12 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar, 231, 481, 7, 1)
                If TrackBar2.Value <= 10 AndAlso TrackBar4.Value >= 10 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar, 231, 491, 7, 1)
                If TrackBar2.Value <= 8 AndAlso TrackBar4.Value >= 8 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar, 231, 501, 7, 1)
                If TrackBar2.Value <= 6 AndAlso TrackBar4.Value >= 6 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar, 231, 511, 7, 1)
                If TrackBar2.Value <= 4 AndAlso TrackBar4.Value >= 4 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar, 231, 521, 7, 1)
                If TrackBar2.Value <= 2 AndAlso TrackBar4.Value >= 2 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar, 231, 531, 7, 1)
            End If


            If My.Settings.NovoValorTransposição(0) = "True" Then
                e.Graphics.DrawImage(My.Resources._8va, 61, 150, 178, 31)
            ElseIf My.Settings.NovoValorTransposição(1) = "True" AndAlso My.Settings.NovoValorFonteCaptaçãoÁudio(1) = "False" Then
                'o símbolo da 8vb não será desenhado se a opção captação de áudio for "Violão"
                'o sistema interpretará as notas como uma oitava abaixo do que está escrito na pauta, mas sem mostrar o símbolo 8vb
                'pois já está convencionando que a pauta de violão  exibe as notas uma oitava acima do que realmente soam
                e.Graphics.DrawImage(My.Resources._8vb, 61, 150, 178, 31)
            ElseIf My.Settings.NovoValorTransposição(2) = "True" Then
                e.Graphics.DrawImage(My.Resources._15ma, 61, 150, 178, 31)
            ElseIf My.Settings.NovoValorTransposição(3) = "True" Then
                e.Graphics.DrawImage(My.Resources._15mb, 61, 150, 178, 31)
            End If


            If SetaPautaTop <> 0 Then e.Graphics.DrawImage(My.Resources.Seta3, SetaPautaLeft - 14, SetaPautaTop, 9, 9)
            If SetaTecladoLeft <> 0 Then e.Graphics.DrawImage(My.Resources.Seta4, SetaTecladoLeft, 707, 9, 7)

            If RadioButton1.Checked Then
                If My.Settings.NovoValorExibirNomeDaNota = True Then

                    e.Graphics.SmoothingMode = SmoothingMode.None
                    If RetanguloBranco <> "" Then e.Graphics.FillRectangle(Branco2, SetaPautaLeft - 30, SetaPautaTop, 15, 9)
                    e.Graphics.SmoothingMode = SmoothingMode.AntiAlias

                    Dim Fonte3 As New Font("Arial", 10, FontStyle.Regular, GraphicsUnit.Pixel)
                    Dim CorDó As SolidBrush = New SolidBrush(Color.FromArgb(2, 0, 0))
                    e.Graphics.DrawString(RetanguloBranco, Fonte3, CorDó, SetaPautaLeft - 28, SetaPautaTop - 2)


                    If ToolTipNota(0) IsNot Nothing Then e.Graphics.DrawImage(ToolTipNota(0), 479, 558, 104, 58)
                End If
            End If





            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias
            'Const MinFrequency As Double = 70
            'Const MaxFrequency As Double = 1200
            'Const DisplayPadding As Integer = 175
            'Dim minStep As Integer = CInt(Math.Truncate(Math.Floor(GetToneStep(MinFrequency))))
            'Dim maxStep As Integer = CInt(Math.Truncate(Math.Ceiling(GetToneStep(MaxFrequency))))
            'Dim totalSteps As Integer = maxStep - minStep
            'Dim stepSize As Single = 5 'CSng(Me.Height - 2 * DisplayPadding) / totalSteps
            Dim fonteNotaCantada As Font = New Font("Times New Roman", 8, FontStyle.Regular, GraphicsUnit.Pixel)

            If CurrentFrequency > 0 Then

                MargemErroInferior = closestFrequency / (2 ^ (1 / (((TrackBar3.Maximum * 2) / (TrackBar3.Maximum - TrackBar3.Value)) * 12)))
                MargemErroSuperior = closestFrequency * (2 ^ (1 / (((TrackBar3.Maximum * 2) / (TrackBar3.Maximum - TrackBar3.Value)) * 12)))

                If Not SinalDetectado Then
                    sliderBrush1 = InactiveSliderBrush1
                    sliderBrush2 = InactiveSliderBrush2
                Else
                    If CurrentFrequency >= MargemErroInferior AndAlso CurrentFrequency <= MargemErroSuperior Then
                        sliderBrush1 = ActiveSliderBrush1
                        sliderBrush2 = ActiveSliderBrush2
                        CorTecla3.BackColor = Color.GreenYellow
                        CorTecla4.BackColor = Color.GreenYellow
                    Else
                        sliderBrush1 = ActiveSliderBrush1B
                        sliderBrush2 = ActiveSliderBrush2B
                        CorTecla3.BackColor = Color.Coral
                        CorTecla4.BackColor = Color.Coral
                    End If
                End If

                'Dim sliderStep As Double = GetToneStep(CurrentFrequency)

                Dim ajusteNotaCantada As Integer = 0
                If DóPrimeiraLinhaToolStripMenuItem1.Checked OrElse DóPrimeiraLinhaToolStripMenuItem3.Checked Then
                    ajusteNotaCantada += -10
                ElseIf DóSegundaLinhaToolStripMenuItem1.Checked OrElse ToolStripMenuItem16.Checked Then
                    ajusteNotaCantada += -20
                ElseIf DóTerceiraLinhaToolStripMenuItem1.Checked OrElse ToolStripMenuItem17.Checked Then
                    ajusteNotaCantada += -30
                ElseIf DóQuartaLinhaToolStripMenuItem1.Checked OrElse ToolStripMenuItem18.Checked Then
                    ajusteNotaCantada += -40
                ElseIf FáTerceiraLinhaToolStripMenuItem1.Checked OrElse ToolStripMenuItem20.Checked Then
                    ajusteNotaCantada += -50
                ElseIf DóQuintaLinhaToolStripMenuItem1.Checked OrElse ToolStripMenuItem19.Checked Then
                    ajusteNotaCantada += -50
                ElseIf FáQuintaLinhaToolStripMenuItem1.Checked OrElse ToolStripMenuItem21.Checked Then
                    ajusteNotaCantada += -70
                ElseIf SolPrimeiraLinhaToolStripMenuItem1.Checked OrElse ToolStripMenuItem22.Checked Then
                    ajusteNotaCantada += 10
                End If

                If My.Settings.NovoValorTransposição(0) = "True" Then
                    ajusteNotaCantada += 35
                ElseIf My.Settings.NovoValorTransposição(1) = "True" Then
                    ajusteNotaCantada += -35
                ElseIf My.Settings.NovoValorTransposição(2) = "True" Then
                    ajusteNotaCantada += 70
                ElseIf My.Settings.NovoValorTransposição(3) = "True" Then
                    ajusteNotaCantada += -70
                End If



                Dim sliderPosition As Single = (AjusteNotas(0) - (NotaCantadaAjustadaParaPauta * 5)) + 4 + ajusteNotaCantada   'CSng(stepSize * (maxStep - sliderStep) + DisplayPadding)

                If stopButton.Enabled Then
                    If My.Settings.NovoValorLinhaSuplementar = True Then

                        If CheckBox3.Checked OrElse (Not CheckBox3.Checked AndAlso Not CheckBox4.Checked) Then
                            If CInt(sliderPosition) <= 251 AndAlso CInt(sliderPosition) > 150 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar2, 195, 251, 30, 1)
                            If CInt(sliderPosition) <= 241 AndAlso CInt(sliderPosition) > 150 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar2, 195, 241, 30, 1)
                            If CInt(sliderPosition) <= 231 AndAlso CInt(sliderPosition) > 150 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar2, 195, 231, 30, 1)
                            If CInt(sliderPosition) <= 221 AndAlso CInt(sliderPosition) > 150 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar2, 195, 221, 30, 1)
                            If CInt(sliderPosition) <= 211 AndAlso CInt(sliderPosition) > 150 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar2, 195, 211, 30, 1)
                            If CInt(sliderPosition) <= 201 AndAlso CInt(sliderPosition) > 150 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar2, 195, 201, 30, 1)
                            If CInt(sliderPosition) <= 191 AndAlso CInt(sliderPosition) > 150 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar2, 195, 191, 30, 1)
                            If CInt(sliderPosition) <= 181 AndAlso CInt(sliderPosition) > 150 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar2, 195, 181, 30, 1)
                            If CInt(sliderPosition) <= 171 AndAlso CInt(sliderPosition) > 150 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar2, 195, 171, 30, 1)

                            If CInt(sliderPosition) >= 311 AndAlso CInt(sliderPosition) < 431 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar2, 195, 311, 30, 1)
                            If CInt(sliderPosition) >= 321 AndAlso CInt(sliderPosition) < 431 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar2, 195, 321, 30, 1)
                            If CInt(sliderPosition) >= 331 AndAlso CInt(sliderPosition) < 431 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar2, 195, 331, 30, 1)
                            If CInt(sliderPosition) >= 341 AndAlso CInt(sliderPosition) < 431 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar2, 195, 341, 30, 1)
                            If CInt(sliderPosition) >= 351 AndAlso CInt(sliderPosition) < 431 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar2, 195, 351, 30, 1)
                            If CInt(sliderPosition) >= 361 AndAlso CInt(sliderPosition) < 431 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar2, 195, 361, 30, 1)
                            If CInt(sliderPosition) >= 371 AndAlso CInt(sliderPosition) < 431 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar2, 195, 371, 30, 1)
                            If CInt(sliderPosition) >= 381 AndAlso CInt(sliderPosition) < 431 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar2, 195, 381, 30, 1)
                            If CInt(sliderPosition) >= 391 AndAlso CInt(sliderPosition) < 431 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar2, 195, 391, 30, 1)
                            If CInt(sliderPosition) >= 401 AndAlso CInt(sliderPosition) < 431 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar2, 195, 401, 30, 1)
                            If CInt(sliderPosition) >= 411 AndAlso CInt(sliderPosition) < 431 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar2, 195, 411, 30, 1)
                            If CInt(sliderPosition) >= 421 AndAlso CInt(sliderPosition) < 431 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar2, 195, 421, 30, 1)

                        End If

                        If CheckBox4.Checked Then
                            If NotaCantadaAjustadaParaPauta >= 39 AndAlso NotaCantadaAjustadaParaPauta < 60 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar2, 195, 421, 30, 1)
                            If NotaCantadaAjustadaParaPauta >= 41 AndAlso NotaCantadaAjustadaParaPauta < 60 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar2, 195, 411, 30, 1)
                            If NotaCantadaAjustadaParaPauta >= 43 AndAlso NotaCantadaAjustadaParaPauta < 60 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar2, 195, 401, 30, 1)
                            If NotaCantadaAjustadaParaPauta >= 45 AndAlso NotaCantadaAjustadaParaPauta < 60 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar2, 195, 391, 30, 1)
                            If NotaCantadaAjustadaParaPauta >= 47 AndAlso NotaCantadaAjustadaParaPauta < 60 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar2, 195, 381, 30, 1)
                            If NotaCantadaAjustadaParaPauta >= 49 AndAlso NotaCantadaAjustadaParaPauta < 60 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar2, 195, 371, 30, 1)
                            If NotaCantadaAjustadaParaPauta >= 51 AndAlso NotaCantadaAjustadaParaPauta < 60 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar2, 195, 361, 30, 1)
                            If NotaCantadaAjustadaParaPauta >= 53 AndAlso NotaCantadaAjustadaParaPauta < 60 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar2, 195, 351, 30, 1)
                            If NotaCantadaAjustadaParaPauta >= 55 AndAlso NotaCantadaAjustadaParaPauta < 60 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar2, 195, 341, 30, 1)
                            If NotaCantadaAjustadaParaPauta >= 57 AndAlso NotaCantadaAjustadaParaPauta < 60 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar2, 195, 331, 30, 1)
                            If NotaCantadaAjustadaParaPauta >= 59 AndAlso NotaCantadaAjustadaParaPauta < 60 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar2, 195, 321, 30, 1)

                            If NotaCantadaAjustadaParaPauta <= 27 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar2, 195, 481, 30, 1)
                            If NotaCantadaAjustadaParaPauta <= 25 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar2, 195, 491, 30, 1)
                            If NotaCantadaAjustadaParaPauta <= 23 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar2, 195, 501, 30, 1)
                            If NotaCantadaAjustadaParaPauta <= 21 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar2, 195, 511, 30, 1)
                            If NotaCantadaAjustadaParaPauta <= 19 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar2, 195, 521, 30, 1)
                            If NotaCantadaAjustadaParaPauta <= 17 Then e.Graphics.DrawImage(My.Resources.LinhaSuplementar2, 195, 531, 30, 1)
                        End If

                    End If






                    If CheckBox3.Checked OrElse (Not CheckBox3.Checked AndAlso Not CheckBox4.Checked) Then
                        e.Graphics.FillPolygon(sliderBrush1, New PointF() {New PointF(210 - 10, sliderPosition), New PointF(210, sliderPosition - 5), New PointF(210, sliderPosition + 5), New PointF(210 + 10, sliderPosition)})
                        e.Graphics.FillPolygon(sliderBrush2, New PointF() {New PointF(210 - 10, sliderPosition), New PointF(210, sliderPosition + 5), New PointF(210, sliderPosition - 5), New PointF(210 + 10, sliderPosition)})
                        e.Graphics.DrawString(CurrentFrequency.ToString("f2"), fonteNotaCantada, Brushes.DarkGreen, 220, sliderPosition - 11)
                        e.Graphics.DrawString(noteNameTextBox.Text, fonteNotaCantada, Brushes.DarkGreen, 220, sliderPosition - 4)
                        e.Graphics.DrawString(closestFrequency.ToString("f2"), fonteNotaCantada, Brushes.DarkGreen, 220, sliderPosition + 3)
                    End If

                    If NotaCantadaAjustadaParaPauta < 60 AndAlso CheckBox4.Checked Then
                        sliderPosition += 110
                        e.Graphics.FillPolygon(sliderBrush1, New PointF() {New PointF(210 - 10, sliderPosition), New PointF(210, sliderPosition - 5), New PointF(210, sliderPosition + 5), New PointF(210 + 10, sliderPosition)})
                        e.Graphics.FillPolygon(sliderBrush2, New PointF() {New PointF(210 - 10, sliderPosition), New PointF(210, sliderPosition + 5), New PointF(210, sliderPosition - 5), New PointF(210 + 10, sliderPosition)})
                        e.Graphics.DrawString(CurrentFrequency.ToString("f2"), fonteNotaCantada, Brushes.DarkGreen, 220, sliderPosition - 11)
                        e.Graphics.DrawString(noteNameTextBox.Text, fonteNotaCantada, Brushes.DarkGreen, 220, sliderPosition - 4)
                        e.Graphics.DrawString(closestFrequency.ToString("f2"), fonteNotaCantada, Brushes.DarkGreen, 220, sliderPosition + 3)
                    End If
                End If
            End If




        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Shared ActiveSliderBrush1 As Brush = New SolidBrush(Color.GreenYellow)
    Shared ActiveSliderBrush2 As Brush = New SolidBrush(Color.Green)
    Shared ActiveSliderBrush1B As Brush = New SolidBrush(Color.FromArgb(255, 150, 0))
    Shared ActiveSliderBrush2B As Brush = New SolidBrush(Color.Red)
    Shared InactiveSliderBrush1 As Brush = New SolidBrush(Color.FromArgb(70, Color.Gray))
    Shared InactiveSliderBrush2 As Brush = New SolidBrush(Color.FromArgb(50, Color.Black))

    Private Function GetToneStep(ByVal frequency As Double) As Double
        Const AFrequency As Double = 440
        Return Math.Log(frequency / AFrequency, ToneStep)
    End Function

    Private Sub DesenharNomeNota0(ByVal gr As Graphics)

        Try

            gr.DrawImage(NomeDaNota(0), 141 - AjusteLeftFigura, (AjusteNotas(0) - (Nota(0) * 5)) - AjusteDasNotasNaClaveFinal(0), AjusteNotas(1), AjusteNotas(2))
            ee = (AjusteNotas(0) - (Nota(0) * 5)) - AjusteDasNotasNaClaveFinal(0)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub SemNota(ByVal gr As Graphics)

        Try

            If Pauta.Checked Then 'só irá desenhar o "?" se opção pauta estiver selecionada. Se opção CifrasdeAcordes estiver selecionada, então "?" não deverá aparecer
                NomedaNota2(0) = My.Resources.SemNota
                gr.DrawImage(NomedaNota2(0), 413, 270, 200, 150)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub LoopNotas(ByVal gr As Graphics)

        Try

            NumeroDeNotas(0) = 0

            Do While NumeroDeNotas(0) < QtdeLoopsNotas

                If RadioButton2.Checked AndAlso CInt(TextBox1.Text) > 1 Then
                    NomeDaNota(NumeroDeNotas(0)) = My.Resources.Notas
                    AjusteNotas(0) = 529 : AjusteNotas(1) = 28 : AjusteNotas(2) = 24 : AjusteLeftFigura = 3
                Else
                    NomeDaNota(NumeroDeNotas(0)) = My.Resources.Dó
                    AjusteNotas(0) = 502 : AjusteNotas(1) = 30 : AjusteNotas(2) = 54 : AjusteLeftFigura = 0
                End If

                NomedaNota2(NumeroDeNotas(0)) = My.Resources.C

                If My.Settings.NovoValorCoresNotas = True Then
                    If LetraCifraFinal(NumeroDeNotas(0)) = "Lá" OrElse LetraCifraFinal(NumeroDeNotas(0)) = "Lá#" OrElse LetraCifraFinal(NumeroDeNotas(0)) = "Láb" Then
                        If Not RadioButton2.Checked Then NomeDaNota(NumeroDeNotas(0)) = My.Resources.Lá
                        NomedaNota2(NumeroDeNotas(0)) = My.Resources.A
                    ElseIf LetraCifraFinal(NumeroDeNotas(0)) = "Si" OrElse LetraCifraFinal(NumeroDeNotas(0)) = "Si#" OrElse LetraCifraFinal(NumeroDeNotas(0)) = "Sib" Then
                        If Not RadioButton2.Checked Then NomeDaNota(NumeroDeNotas(0)) = My.Resources.Si
                        NomedaNota2(NumeroDeNotas(0)) = My.Resources.B
                    ElseIf LetraCifraFinal(NumeroDeNotas(0)) = "Dó" OrElse LetraCifraFinal(NumeroDeNotas(0)) = "Dó#" OrElse LetraCifraFinal(NumeroDeNotas(0)) = "Dób" Then
                        If Not RadioButton2.Checked Then NomeDaNota(NumeroDeNotas(0)) = My.Resources.Dó
                        NomedaNota2(NumeroDeNotas(0)) = My.Resources.C
                    ElseIf LetraCifraFinal(NumeroDeNotas(0)) = "Ré" OrElse LetraCifraFinal(NumeroDeNotas(0)) = "Ré#" OrElse LetraCifraFinal(NumeroDeNotas(0)) = "Réb" Then
                        If Not RadioButton2.Checked Then NomeDaNota(NumeroDeNotas(0)) = My.Resources.Ré
                        NomedaNota2(NumeroDeNotas(0)) = My.Resources.D
                    ElseIf LetraCifraFinal(NumeroDeNotas(0)) = "Mi" OrElse LetraCifraFinal(NumeroDeNotas(0)) = "Mi#" OrElse LetraCifraFinal(NumeroDeNotas(0)) = "Mib" Then
                        If Not RadioButton2.Checked Then NomeDaNota(NumeroDeNotas(0)) = My.Resources.Mi
                        NomedaNota2(NumeroDeNotas(0)) = My.Resources.E
                    ElseIf LetraCifraFinal(NumeroDeNotas(0)) = "Fá" OrElse LetraCifraFinal(NumeroDeNotas(0)) = "Fá#" OrElse LetraCifraFinal(NumeroDeNotas(0)) = "Fáb" Then
                        If Not RadioButton2.Checked Then NomeDaNota(NumeroDeNotas(0)) = My.Resources.Fá
                        NomedaNota2(NumeroDeNotas(0)) = My.Resources.F
                    ElseIf LetraCifraFinal(NumeroDeNotas(0)) = "Sol" OrElse LetraCifraFinal(NumeroDeNotas(0)) = "Sol#" OrElse LetraCifraFinal(NumeroDeNotas(0)) = "Solb" Then
                        If Not RadioButton2.Checked Then NomeDaNota(NumeroDeNotas(0)) = My.Resources.Sol
                        NomedaNota2(NumeroDeNotas(0)) = My.Resources.G
                    End If
                Else
                    If LetraCifraFinal(NumeroDeNotas(0)) = "Lá" OrElse LetraCifraFinal(NumeroDeNotas(0)) = "Lá#" OrElse LetraCifraFinal(NumeroDeNotas(0)) = "Láb" Then
                        NomedaNota2(NumeroDeNotas(0)) = My.Resources.A_preto
                    ElseIf LetraCifraFinal(NumeroDeNotas(0)) = "Si" OrElse LetraCifraFinal(NumeroDeNotas(0)) = "Si#" OrElse LetraCifraFinal(NumeroDeNotas(0)) = "Sib" Then
                        NomedaNota2(NumeroDeNotas(0)) = My.Resources.B_preto
                    ElseIf LetraCifraFinal(NumeroDeNotas(0)) = "Dó" OrElse LetraCifraFinal(NumeroDeNotas(0)) = "Dó#" OrElse LetraCifraFinal(NumeroDeNotas(0)) = "Dób" Then
                        NomedaNota2(NumeroDeNotas(0)) = My.Resources.C
                    ElseIf LetraCifraFinal(NumeroDeNotas(0)) = "Ré" OrElse LetraCifraFinal(NumeroDeNotas(0)) = "Ré#" OrElse LetraCifraFinal(NumeroDeNotas(0)) = "Réb" Then
                        NomedaNota2(NumeroDeNotas(0)) = My.Resources.D_preto
                    ElseIf LetraCifraFinal(NumeroDeNotas(0)) = "Mi" OrElse LetraCifraFinal(NumeroDeNotas(0)) = "Mi#" OrElse LetraCifraFinal(NumeroDeNotas(0)) = "Mib" Then
                        NomedaNota2(NumeroDeNotas(0)) = My.Resources.E_preto
                    ElseIf LetraCifraFinal(NumeroDeNotas(0)) = "Fá" OrElse LetraCifraFinal(NumeroDeNotas(0)) = "Fá#" OrElse LetraCifraFinal(NumeroDeNotas(0)) = "Fáb" Then
                        NomedaNota2(NumeroDeNotas(0)) = My.Resources.F_preto
                    ElseIf LetraCifraFinal(NumeroDeNotas(0)) = "Sol" OrElse LetraCifraFinal(NumeroDeNotas(0)) = "Sol#" OrElse LetraCifraFinal(NumeroDeNotas(0)) = "Solb" Then
                        NomedaNota2(NumeroDeNotas(0)) = My.Resources.G_preto
                    End If
                End If

                If Nota(NumeroDeNotas(0)) >= 32 AndAlso EventoPaintGeradoPelo_ExecutaLoopNotasAleatórias = True Then AjusteDasNotasNaClaveFinal(NumeroDeNotas(0)) = AjusteDasNotasNaClaveFinal(NumeroDeNotas(0)) + 35

                If My.Settings.NovoValorLinhaSuplementar = True Then
                    If Not RadioButton2.Checked OrElse (RadioButton2.Checked AndAlso Visual.Checked) Then

                        'os valores 36, 40, 10 e 14, que são somados abaixo, foram colocados para ajustar a exibição das linhas suplementares.
                        AjusteLinhaSuplementar(0) = 36 : AjusteLinhaSuplementar(1) = 40 'figura ritmica com altura maior
                        If RadioButton2.Checked AndAlso CInt(TextBox1.Text) > 1 Then AjusteLinhaSuplementar(0) = 10 : AjusteLinhaSuplementar(1) = 14 'figura ritmica possui altura menor

                        Dim PosiçãoAlturaNotas1 As Integer = (AjusteNotas(0) - (Nota(NumeroDeNotas(0)) * 5)) - AjusteDasNotasNaClaveFinal(NumeroDeNotas(0)) + AjusteLinhaSuplementar(0)
                        Dim PosiçãoAlturaNotas2 As Integer = (AjusteNotas(0) - (Nota(NumeroDeNotas(0)) * 5)) - AjusteDasNotasNaClaveFinal(NumeroDeNotas(0)) + AjusteLinhaSuplementar(1)

                        'as linhas suplementares são exibidas conforme a posição da altura da nota.. exceto a clave de Fá (piano), a qual pode ser usada a variável k, que foram armazenados em Nota(0), Nota(1), Nota(2) e Nota(3)
                        If PosiçãoAlturaNotas1 <= 251 AndAlso PosiçãoAlturaNotas1 > 150 Then gr.DrawImage(My.Resources.LinhaSuplementar2, 136, 251, 30, 1)
                        If PosiçãoAlturaNotas1 <= 241 AndAlso PosiçãoAlturaNotas1 > 150 Then gr.DrawImage(My.Resources.LinhaSuplementar2, 136, 241, 30, 1)
                        If PosiçãoAlturaNotas1 <= 231 AndAlso PosiçãoAlturaNotas1 > 150 Then gr.DrawImage(My.Resources.LinhaSuplementar2, 136, 231, 30, 1)
                        If PosiçãoAlturaNotas1 <= 221 AndAlso PosiçãoAlturaNotas1 > 150 Then gr.DrawImage(My.Resources.LinhaSuplementar2, 136, 221, 30, 1)
                        If PosiçãoAlturaNotas1 <= 211 AndAlso PosiçãoAlturaNotas1 > 150 Then gr.DrawImage(My.Resources.LinhaSuplementar2, 136, 211, 30, 1)
                        If PosiçãoAlturaNotas1 <= 201 AndAlso PosiçãoAlturaNotas1 > 150 Then gr.DrawImage(My.Resources.LinhaSuplementar2, 136, 201, 30, 1)
                        If PosiçãoAlturaNotas1 <= 191 AndAlso PosiçãoAlturaNotas1 > 150 Then gr.DrawImage(My.Resources.LinhaSuplementar2, 136, 191, 30, 1)
                        If PosiçãoAlturaNotas1 <= 181 AndAlso PosiçãoAlturaNotas1 > 150 Then gr.DrawImage(My.Resources.LinhaSuplementar2, 136, 181, 30, 1)
                        If PosiçãoAlturaNotas1 <= 171 AndAlso PosiçãoAlturaNotas1 > 150 Then gr.DrawImage(My.Resources.LinhaSuplementar2, 136, 171, 30, 1)

                        If PosiçãoAlturaNotas2 >= 311 AndAlso PosiçãoAlturaNotas2 < 351 Then gr.DrawImage(My.Resources.LinhaSuplementar2, 136, 311, 30, 1)
                        If PosiçãoAlturaNotas2 >= 321 AndAlso PosiçãoAlturaNotas2 < 351 Then gr.DrawImage(My.Resources.LinhaSuplementar2, 136, 321, 30, 1)
                        If PosiçãoAlturaNotas2 >= 331 AndAlso PosiçãoAlturaNotas2 < 351 Then gr.DrawImage(My.Resources.LinhaSuplementar2, 136, 331, 30, 1)
                        If PosiçãoAlturaNotas2 >= 341 AndAlso PosiçãoAlturaNotas2 < 351 Then gr.DrawImage(My.Resources.LinhaSuplementar2, 136, 341, 30, 1)

                        If CheckBox4.Checked Then
                            'para a clave de Fá (piano) dá para usar os valores da variável k, que foram armazenados em Nota(0), Nota(1), Nota(2) e Nota(3)
                            If Nota(NumeroDeNotas(0)) >= 24 AndAlso Nota(NumeroDeNotas(0)) < 32 Then gr.DrawImage(My.Resources.LinhaSuplementar2, 136, 421, 30, 1)
                            If Nota(NumeroDeNotas(0)) >= 26 AndAlso Nota(NumeroDeNotas(0)) < 32 Then gr.DrawImage(My.Resources.LinhaSuplementar2, 136, 411, 30, 1)
                            If Nota(NumeroDeNotas(0)) >= 28 AndAlso Nota(NumeroDeNotas(0)) < 32 Then gr.DrawImage(My.Resources.LinhaSuplementar2, 136, 401, 30, 1)
                            If Nota(NumeroDeNotas(0)) >= 30 AndAlso Nota(NumeroDeNotas(0)) < 32 Then gr.DrawImage(My.Resources.LinhaSuplementar2, 136, 391, 30, 1)

                            If Nota(NumeroDeNotas(0)) <= 12 Then gr.DrawImage(My.Resources.LinhaSuplementar2, 136, 481, 30, 1)
                            If Nota(NumeroDeNotas(0)) <= 10 Then gr.DrawImage(My.Resources.LinhaSuplementar2, 136, 491, 30, 1)
                            If Nota(NumeroDeNotas(0)) <= 8 Then gr.DrawImage(My.Resources.LinhaSuplementar2, 136, 501, 30, 1)
                            If Nota(NumeroDeNotas(0)) <= 6 Then gr.DrawImage(My.Resources.LinhaSuplementar2, 136, 511, 30, 1)
                            If Nota(NumeroDeNotas(0)) <= 4 Then gr.DrawImage(My.Resources.LinhaSuplementar2, 136, 521, 30, 1)
                            If Nota(NumeroDeNotas(0)) <= 2 Then gr.DrawImage(My.Resources.LinhaSuplementar2, 136, 531, 30, 1)
                        End If
                    End If
                End If

                NumeroDeNotas(0) += 1
            Loop


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub ExibiçãoSustenidosBemóis1(ByVal gr As Graphics)

        'exibição dos sustenidos ou bemóis, só gera sustenidos e bemois se pontuação por "Cifras" não estiver marcado

        Try

            If mFinal(0) = "#" Then
                If NomeDoMenu = "Nenhum" Then gr.DrawImage(My.Resources.Sustenido_pequeno, 127, (515 - (Nota(0) * 5)) - AjusteDasNotasNaClaveFinal(0), 22, 36)
                If My.Settings.NovoValorExibirNomeDaNota = True AndAlso Not RadioButton2.Checked Then gr.DrawImage(My.Resources.Sustenido_grande, 560, 290, 48, 109)
            ElseIf mFinal(0) = "b" Then
                If NomeDoMenu = "Nenhum" Then gr.DrawImage(My.Resources.Bemol_pequeno, 127, (515 - (Nota(0) * 5)) - AjusteDasNotasNaClaveFinal(0), 21, 33)
                If My.Settings.NovoValorExibirNomeDaNota = True AndAlso Not RadioButton2.Checked Then gr.DrawImage(My.Resources.Bemol_grande, 560, 295, 43, 96)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub ExibiçãoSustenidosBemóis2(ByVal gr As Graphics)

        'exibição dos sustenidos ou bemóis, só gera sustenidos e bemois se pontuação por "Cifras" não estiver marcado

        Try

            If mFinal(1) = "#" Then
                If NomeDoMenu = "Nenhum" Then gr.DrawImage(My.Resources.Sustenido_pequeno, 127, (515 - (Nota(1) * 5)) - AjusteDasNotasNaClaveFinal(1), 22, 36)
            ElseIf mFinal(1) = "b" Then
                If NomeDoMenu = "Nenhum" Then gr.DrawImage(My.Resources.Bemol_pequeno, 127, (515 - (Nota(1) * 5)) - AjusteDasNotasNaClaveFinal(1), 21, 33)
            End If

            If mFinal(2) = "#" Then
                If NomeDoMenu = "Nenhum" Then gr.DrawImage(My.Resources.Sustenido_pequeno, 127, (515 - (Nota(2) * 5)) - AjusteDasNotasNaClaveFinal(2), 22, 36)
            ElseIf mFinal(2) = "b" Then
                If NomeDoMenu = "Nenhum" Then gr.DrawImage(My.Resources.Bemol_pequeno, 127, (515 - (Nota(2) * 5)) - AjusteDasNotasNaClaveFinal(2), 21, 33)
            End If

            If mFinal(3) = "#" Then
                If NomeDoMenu = "Nenhum" Then gr.DrawImage(My.Resources.Sustenido_pequeno, 127, (515 - (Nota(3) * 5)) - AjusteDasNotasNaClaveFinal(3), 22, 36)
            ElseIf mFinal(3) = "b" Then
                If NomeDoMenu = "Nenhum" Then gr.DrawImage(My.Resources.Bemol_pequeno, 127, (515 - (Nota(3) * 5)) - AjusteDasNotasNaClaveFinal(3), 21, 33)
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub ExibiçãoDosIntervalos(ByVal gr As Graphics)

        Try

            If (My.Settings.NovoValorIntervalosPauta(2) = "True" OrElse My.Settings.NovoValorIntervalosPauta(3) = "True" OrElse My.Settings.NovoValorIntervalosPauta(4) = "True" OrElse _
                My.Settings.NovoValorIntervalosPauta(5) = "True" OrElse My.Settings.NovoValorIntervalosPauta(6) = "True" OrElse My.Settings.NovoValorIntervalosPauta(7) = "True" OrElse My.Settings.NovoValorIntervalosPauta(8) = "True") Then

                ValorLeft(0) = 173
                ValorLeft(1) = 0
                ValorLeft(2) = 0
                ValorLeft(3) = 0
                ValorLeft(4) = 0
                ValorLeft(5) = 0
                ValorLeft(6) = 0
                ValorLeft(7) = 0
                gr.DrawImage(My.Resources.Fundamental, ValorLeft(0), ee + 34, 16, 11)

                If My.Settings.NovoValorIntervalosPauta(2) = "True" AndAlso k <= 67 Then
                    ValorLeft(1) = 173
                    If ValorLeft(0) = 173 Then ValorLeft(1) = 188
                    gr.DrawImage(My.Resources.Segunda, ValorLeft(1), ee + 29, 16, 11)

                    If LetraCifra = "Mi" OrElse LetraCifra = "Si" OrElse LetraCifra = "Fá#" OrElse LetraCifra = "Dó#" OrElse LetraCifra = "Sol#" OrElse LetraCifra = "Ré#" OrElse LetraCifra = "Lá#" OrElse LetraCifra = "Mi#" OrElse LetraCifra = "Si#" Then
                        gr.DrawImage(My.Resources.Sustenido_Intervalos, ValorLeft(1) - 10, ee + 29, 16, 11)
                    ElseIf LetraCifra = "Láb" OrElse LetraCifra = "Réb" OrElse LetraCifra = "Solb" OrElse LetraCifra = "Dób" OrElse LetraCifra = "Fáb" Then
                        gr.DrawImage(My.Resources.Bemol_Intervalos, ValorLeft(1) - 10, ee + 29, 16, 11)
                    End If
                End If
                If My.Settings.NovoValorIntervalosPauta(3) = "True" AndAlso k <= 66 Then
                    ValorLeft(2) = 173
                    If ValorLeft(1) = 173 Then ValorLeft(2) = 188
                    gr.DrawImage(My.Resources.Terça, ValorLeft(2), ee + 24, 16, 11)
                    If LetraCifra = "Ré" OrElse LetraCifra = "Lá" OrElse LetraCifra = "Mi" OrElse LetraCifra = "Si" OrElse LetraCifra = "Fá#" OrElse LetraCifra = "Dó#" OrElse LetraCifra = "Sol#" Then
                        gr.DrawImage(My.Resources.Sustenido_Intervalos, ValorLeft(2) - 10, ee + 24, 16, 11)
                    ElseIf LetraCifra = "Ré#" OrElse LetraCifra = "Lá#" OrElse LetraCifra = "Mi#" OrElse LetraCifra = "Si#" Then
                        gr.DrawImage(My.Resources.DobradoSustenido_Intervalos, ValorLeft(2) - 10, ee + 24, 16, 11)
                    ElseIf LetraCifra = "Solb" OrElse LetraCifra = "Dób" OrElse LetraCifra = "Fáb" Then
                        gr.DrawImage(My.Resources.Bemol_Intervalos, ValorLeft(2) - 10, ee + 24, 16, 11)
                    End If
                End If
                If My.Settings.NovoValorIntervalosPauta(4) = "True" AndAlso k <= 65 Then
                    ValorLeft(3) = 173
                    If ValorLeft(2) = 173 Then ValorLeft(3) = 188
                    gr.DrawImage(My.Resources.Quarta, ValorLeft(3), ee + 19, 16, 11)
                    If LetraCifra = "Dó#" OrElse LetraCifra = "Sol#" OrElse LetraCifra = "Ré#" OrElse LetraCifra = "Lá#" OrElse LetraCifra = "Mi#" OrElse LetraCifra = "Si#" Then
                        gr.DrawImage(My.Resources.Sustenido_Intervalos, ValorLeft(3) - 10, ee + 19, 16, 11)
                    ElseIf LetraCifra = "Fáb" Then
                        gr.DrawImage(My.Resources.DobradoBemol_Intervalos, ValorLeft(3) - 10, ee + 19, 16, 11)
                    ElseIf LetraCifra = "Fá" OrElse LetraCifra = "Sib" OrElse LetraCifra = "Mib" OrElse LetraCifra = "Láb" OrElse LetraCifra = "Réb" OrElse LetraCifra = "Solb" OrElse LetraCifra = "Dób" Then
                        gr.DrawImage(My.Resources.Bemol_Intervalos, ValorLeft(3) - 10, ee + 19, 16, 11)
                    End If
                End If
                If My.Settings.NovoValorIntervalosPauta(5) = "True" AndAlso k <= 64 Then
                    ValorLeft(4) = 173
                    If ValorLeft(3) = 173 Then ValorLeft(4) = 188
                    gr.DrawImage(My.Resources.Quinta, ValorLeft(4), ee + 14, 16, 11)
                    If LetraCifra = "Si" OrElse LetraCifra = "Fá#" OrElse LetraCifra = "Dó#" OrElse LetraCifra = "Sol#" OrElse LetraCifra = "Ré#" OrElse LetraCifra = "Lá#" OrElse LetraCifra = "Mi#" Then
                        gr.DrawImage(My.Resources.Sustenido_Intervalos, ValorLeft(4) - 10, ee + 14, 16, 11)
                    ElseIf LetraCifra = "Si#" Then
                        gr.DrawImage(My.Resources.DobradoSustenido_Intervalos, ValorLeft(4) - 10, ee + 14, 16, 11)
                    ElseIf LetraCifra = "Mib" OrElse LetraCifra = "Láb" OrElse LetraCifra = "Réb" OrElse LetraCifra = "Solb" OrElse LetraCifra = "Dób" OrElse LetraCifra = "Fáb" Then
                        gr.DrawImage(My.Resources.Bemol_Intervalos, ValorLeft(4) - 10, ee + 14, 16, 11)
                    End If
                End If
                If My.Settings.NovoValorIntervalosPauta(6) = "True" AndAlso k <= 63 Then
                    ValorLeft(5) = 173
                    If ValorLeft(4) = 173 Then ValorLeft(5) = 188
                    gr.DrawImage(My.Resources.Sexta, ValorLeft(5), ee + 9, 16, 11)
                    If LetraCifra = "Lá" OrElse LetraCifra = "Mi" OrElse LetraCifra = "Si" OrElse LetraCifra = "Fá#" OrElse LetraCifra = "Dó#" OrElse LetraCifra = "Sol#" OrElse LetraCifra = "Ré#" Then
                        gr.DrawImage(My.Resources.Sustenido_Intervalos, ValorLeft(5) - 10, ee + 9, 16, 11)
                    ElseIf LetraCifra = "Lá#" OrElse LetraCifra = "Mi#" OrElse LetraCifra = "Si#" Then
                        gr.DrawImage(My.Resources.DobradoSustenido_Intervalos, ValorLeft(5) - 10, ee + 9, 16, 11)
                    ElseIf LetraCifra = "Réb" OrElse LetraCifra = "Solb" OrElse LetraCifra = "Dób" OrElse LetraCifra = "Fáb" Then
                        gr.DrawImage(My.Resources.Bemol_Intervalos, ValorLeft(5) - 10, ee + 9, 16, 11)
                    End If
                End If
                If My.Settings.NovoValorIntervalosPauta(7) = "True" AndAlso k <= 62 Then
                    ValorLeft(6) = 173
                    If ValorLeft(5) = 173 Then ValorLeft(6) = 188
                    gr.DrawImage(My.Resources.Sétima, ValorLeft(6), ee + 4, 16, 11)
                    If LetraCifra = "Sol" OrElse LetraCifra = "Ré" OrElse LetraCifra = "Lá" OrElse LetraCifra = "Mi" OrElse LetraCifra = "Si" OrElse LetraCifra = "Fá#" OrElse LetraCifra = "Dó#" Then
                        gr.DrawImage(My.Resources.Sustenido_Intervalos, ValorLeft(6) - 10, ee + 4, 16, 11)
                    ElseIf LetraCifra = "Sol#" OrElse LetraCifra = "Ré#" OrElse LetraCifra = "Lá#" OrElse LetraCifra = "Mi#" OrElse LetraCifra = "Si#" Then
                        gr.DrawImage(My.Resources.DobradoSustenido_Intervalos, ValorLeft(6) - 10, ee + 4, 16, 11)
                    ElseIf LetraCifra = "Dób" OrElse LetraCifra = "Mib" Then
                        gr.DrawImage(My.Resources.Bemol_Intervalos, ValorLeft(6) - 10, ee + 4, 16, 11)
                    End If
                End If
                If My.Settings.NovoValorIntervalosPauta(8) = "True" AndAlso k <= 61 Then
                    ValorLeft(7) = 173
                    If ValorLeft(6) = 173 Then ValorLeft(7) = 188
                    gr.DrawImage(My.Resources.Oitava, ValorLeft(7), ee - 1, 16, 11)
                    If LetraCifra = "Fá#" OrElse LetraCifra = "Dó#" OrElse LetraCifra = "Sol#" OrElse LetraCifra = "Ré#" OrElse LetraCifra = "Lá#" OrElse LetraCifra = "Mi#" OrElse LetraCifra = "Si#" Then
                        gr.DrawImage(My.Resources.Sustenido_Intervalos, ValorLeft(7) - 10, ee - 1, 16, 11)
                    ElseIf LetraCifra = "Sib" OrElse LetraCifra = "Mib" OrElse LetraCifra = "Láb" OrElse LetraCifra = "Réb" OrElse LetraCifra = "Solb" OrElse LetraCifra = "Dób" OrElse LetraCifra = "Fáb" Then
                        gr.DrawImage(My.Resources.Bemol_Intervalos, ValorLeft(7) - 10, ee - 1, 16, 11)
                    End If
                End If
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub IdentificarPosiçãoSetasDó(ByVal gr As Graphics)

        Try

            If My.Settings.NovoValorIdentificarPosiçãoDóCentral = True AndAlso My.Settings.NovoValorIdentificarPosiçãoTodasNotasDó = False Then
                If CheckBox3.Checked Then
                    gr.DrawImage(My.Resources.Seta, 215, 307, 9, 9)
                End If
                If CheckBox4.Checked Then
                    gr.DrawImage(My.Resources.Seta, 215, 417, 9, 9)
                End If
                If DóPrimeiraLinhaToolStripMenuItem1.Checked Then
                    gr.DrawImage(My.Resources.Seta, 215, 297, 9, 9)
                ElseIf DóSegundaLinhaToolStripMenuItem1.Checked Then
                    gr.DrawImage(My.Resources.Seta, 215, 287, 9, 9)
                ElseIf DóTerceiraLinhaToolStripMenuItem1.Checked Then
                    gr.DrawImage(My.Resources.Seta, 215, 277, 9, 9)
                ElseIf DóQuartaLinhaToolStripMenuItem1.Checked Then
                    gr.DrawImage(My.Resources.Seta, 215, 267, 9, 9)
                ElseIf DóQuintaLinhaToolStripMenuItem1.Checked OrElse FáTerceiraLinhaToolStripMenuItem1.Checked Then
                    gr.DrawImage(My.Resources.Seta, 215, 257, 9, 9)
                ElseIf FáQuintaLinhaToolStripMenuItem1.Checked Then
                    gr.DrawImage(My.Resources.Seta, 215, 237, 9, 9)
                ElseIf SolPrimeiraLinhaToolStripMenuItem1.Checked Then
                    gr.DrawImage(My.Resources.Seta, 215, 317, 9, 9)
                End If
            ElseIf (My.Settings.NovoValorIdentificarPosiçãoDóCentral = False AndAlso My.Settings.NovoValorIdentificarPosiçãoTodasNotasDó = True) OrElse (My.Settings.NovoValorIdentificarPosiçãoDóCentral = True AndAlso My.Settings.NovoValorIdentificarPosiçãoTodasNotasDó = True) Then
                If CheckBox3.Checked OrElse FáQuintaLinhaToolStripMenuItem1.Checked Then
                    gr.DrawImage(My.Resources.Seta, 215, 167, 9, 9)
                    gr.DrawImage(My.Resources.Seta, 215, 202, 9, 9)
                    gr.DrawImage(My.Resources.Seta, 215, 237, 9, 9)
                    gr.DrawImage(My.Resources.Seta, 215, 272, 9, 9)
                    gr.DrawImage(My.Resources.Seta, 215, 307, 9, 9)
                    gr.DrawImage(My.Resources.Seta, 215, 342, 9, 9)
                End If
                If CheckBox4.Checked Then
                    gr.DrawImage(My.Resources.Seta, 215, 383, 9, 9)
                    gr.DrawImage(My.Resources.Seta, 215, 417, 9, 9)
                    gr.DrawImage(My.Resources.Seta, 215, 452, 9, 9)
                    gr.DrawImage(My.Resources.Seta, 215, 487, 9, 9)
                    gr.DrawImage(My.Resources.Seta, 215, 523, 9, 9)
                End If
                If DóPrimeiraLinhaToolStripMenuItem1.Checked Then
                    gr.DrawImage(My.Resources.Seta, 215, 192, 9, 9)
                    gr.DrawImage(My.Resources.Seta, 215, 227, 9, 9)
                    gr.DrawImage(My.Resources.Seta, 215, 262, 9, 9)
                    gr.DrawImage(My.Resources.Seta, 215, 297, 9, 9)
                    gr.DrawImage(My.Resources.Seta, 215, 332, 9, 9)
                ElseIf DóSegundaLinhaToolStripMenuItem1.Checked Then
                    gr.DrawImage(My.Resources.Seta, 215, 182, 9, 9)
                    gr.DrawImage(My.Resources.Seta, 215, 217, 9, 9)
                    gr.DrawImage(My.Resources.Seta, 215, 252, 9, 9)
                    gr.DrawImage(My.Resources.Seta, 215, 287, 9, 9)
                    gr.DrawImage(My.Resources.Seta, 215, 322, 9, 9)
                ElseIf DóTerceiraLinhaToolStripMenuItem1.Checked Then
                    gr.DrawImage(My.Resources.Seta, 215, 172, 9, 9)
                    gr.DrawImage(My.Resources.Seta, 215, 207, 9, 9)
                    gr.DrawImage(My.Resources.Seta, 215, 242, 9, 9)
                    gr.DrawImage(My.Resources.Seta, 215, 277, 9, 9)
                    gr.DrawImage(My.Resources.Seta, 215, 312, 9, 9)
                ElseIf DóQuartaLinhaToolStripMenuItem1.Checked Then
                    gr.DrawImage(My.Resources.Seta, 215, 197, 9, 9)
                    gr.DrawImage(My.Resources.Seta, 215, 232, 9, 9)
                    gr.DrawImage(My.Resources.Seta, 215, 267, 9, 9)
                    gr.DrawImage(My.Resources.Seta, 215, 302, 9, 9)
                    gr.DrawImage(My.Resources.Seta, 215, 337, 9, 9)
                ElseIf DóQuintaLinhaToolStripMenuItem1.Checked OrElse FáTerceiraLinhaToolStripMenuItem1.Checked Then
                    gr.DrawImage(My.Resources.Seta, 215, 187, 9, 9)
                    gr.DrawImage(My.Resources.Seta, 215, 222, 9, 9)
                    gr.DrawImage(My.Resources.Seta, 215, 257, 9, 9)
                    gr.DrawImage(My.Resources.Seta, 215, 292, 9, 9)
                    gr.DrawImage(My.Resources.Seta, 215, 327, 9, 9)
                ElseIf SolPrimeiraLinhaToolStripMenuItem1.Checked Then
                    gr.DrawImage(My.Resources.Seta, 215, 177, 9, 9)
                    gr.DrawImage(My.Resources.Seta, 215, 212, 9, 9)
                    gr.DrawImage(My.Resources.Seta, 215, 247, 9, 9)
                    gr.DrawImage(My.Resources.Seta, 215, 282, 9, 9)
                    gr.DrawImage(My.Resources.Seta, 215, 317, 9, 9)
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub IdentificarLinhasGuias(ByVal gr As Graphics)

        Try

            If My.Settings.NovoValorIdentificarLinhas = True Then
                If DóPrimeiraLinhaToolStripMenuItem1.Checked Then
                    gr.DrawImage(My.Resources.Linhas_guiasC1, 227, 256, 23, 51)
                ElseIf DóSegundaLinhaToolStripMenuItem1.Checked Then
                    gr.DrawImage(My.Resources.Linhas_guiasC2, 227, 256, 23, 51)
                ElseIf DóTerceiraLinhaToolStripMenuItem1.Checked Then
                    gr.DrawImage(My.Resources.Linhas_guiasC3, 227, 256, 23, 51)
                ElseIf DóQuartaLinhaToolStripMenuItem1.Checked Then
                    gr.DrawImage(My.Resources.Linhas_guiasC4, 227, 256, 23, 51)
                ElseIf DóQuintaLinhaToolStripMenuItem1.Checked OrElse FáTerceiraLinhaToolStripMenuItem1.Checked Then
                    gr.DrawImage(My.Resources.Linhas_guiasC5F3, 227, 256, 23, 51)
                ElseIf FáQuintaLinhaToolStripMenuItem1.Checked Then
                    gr.DrawImage(My.Resources.Linhas_guiasF5, 227, 256, 23, 51)
                ElseIf SolPrimeiraLinhaToolStripMenuItem1.Checked Then
                    gr.DrawImage(My.Resources.Linhas_guiasG1, 227, 256, 23, 51)
                Else
                    gr.DrawImage(My.Resources.Linhas_guias, 227, 235, 23, 265)
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub ArmadurasDeClave(ByVal gr As Graphics)

        Try

            'Tom escolhido no menu
            If NomeDoMenu = "G - Em (F#)" Then
                If CoresArmadura = False Then
                    gr.DrawImage(My.Resources.ClaveG2, 84, 246 + AjusteArmaduraClaves, 57, 52)
                    If CheckBox3.Checked OrElse CheckBox4.Checked Then gr.DrawImage(My.Resources.ClaveG2, 84, 426, 57, 52)
                Else
                    gr.DrawImage(My.Resources.ClaveG2Col, 84, 246 + AjusteArmaduraClaves, 57, 52)
                    If CheckBox3.Checked OrElse CheckBox4.Checked Then gr.DrawImage(My.Resources.ClaveG2Col, 84, 426, 57, 52)
                End If
            ElseIf NomeDoMenu = "D - Bm (F# - C#)" Then
                If CoresArmadura = False Then
                    gr.DrawImage(My.Resources.ClaveD2, 84, 246 + AjusteArmaduraClaves, 57, 52)
                    If CheckBox3.Checked OrElse CheckBox4.Checked Then gr.DrawImage(My.Resources.ClaveD2, 84, 426, 57, 52)
                Else
                    gr.DrawImage(My.Resources.ClaveD2Col, 84, 246 + AjusteArmaduraClaves, 57, 52)
                    If CheckBox3.Checked OrElse CheckBox4.Checked Then gr.DrawImage(My.Resources.ClaveD2Col, 84, 426, 57, 52)
                End If
            ElseIf NomeDoMenu = "A - F#m (F# - C# - G#)" Then
                If CoresArmadura = False Then
                    gr.DrawImage(My.Resources.ClaveA2, 84, 246 + AjusteArmaduraClaves, 57, 52)
                    If CheckBox3.Checked OrElse CheckBox4.Checked Then gr.DrawImage(My.Resources.ClaveA2, 84, 426, 57, 52)
                Else
                    gr.DrawImage(My.Resources.ClaveA2Col, 84, 246 + AjusteArmaduraClaves, 57, 52)
                    If CheckBox3.Checked OrElse CheckBox4.Checked Then gr.DrawImage(My.Resources.ClaveA2Col, 84, 426, 57, 52)
                End If
            ElseIf NomeDoMenu = "E - C#m (F# - C# - G# - D#)" Then
                If CoresArmadura = False Then
                    gr.DrawImage(My.Resources.ClaveE2, 84, 246 + AjusteArmaduraClaves, 57, 52)
                    If CheckBox3.Checked OrElse CheckBox4.Checked Then gr.DrawImage(My.Resources.ClaveE2, 84, 426, 57, 52)
                Else
                    gr.DrawImage(My.Resources.ClaveE2Col, 84, 246 + AjusteArmaduraClaves, 57, 52)
                    If CheckBox3.Checked OrElse CheckBox4.Checked Then gr.DrawImage(My.Resources.ClaveE2Col, 84, 426, 57, 52)
                End If
            ElseIf NomeDoMenu = "B - G#m (F# - C# - G# - D# - A#)" Then
                If CoresArmadura = False Then
                    gr.DrawImage(My.Resources.ClaveB2, 84, 246 + AjusteArmaduraClaves, 57, 52)
                    If CheckBox3.Checked OrElse CheckBox4.Checked Then gr.DrawImage(My.Resources.ClaveB2, 84, 426, 57, 52)
                Else
                    gr.DrawImage(My.Resources.ClaveB2Col, 84, 246 + AjusteArmaduraClaves, 57, 52)
                    If CheckBox3.Checked OrElse CheckBox4.Checked Then gr.DrawImage(My.Resources.ClaveB2Col, 84, 426, 57, 52)
                End If
            ElseIf NomeDoMenu = "F# - D#m (F# - C# - G# - D# - A# - E#)" Then
                If CoresArmadura = False Then
                    gr.DrawImage(My.Resources.ClaveFsus2, 84, 246 + AjusteArmaduraClaves, 57, 52)
                    If CheckBox3.Checked OrElse CheckBox4.Checked Then gr.DrawImage(My.Resources.ClaveFsus2, 84, 426, 57, 52)
                Else
                    gr.DrawImage(My.Resources.ClaveFsus2Col, 84, 246 + AjusteArmaduraClaves, 57, 52)
                    If CheckBox3.Checked OrElse CheckBox4.Checked Then gr.DrawImage(My.Resources.ClaveFsus2Col, 84, 426, 57, 52)
                End If
            ElseIf NomeDoMenu = "C# - A#m (F# - C# - G# - D# - A# - E# - B#)" Then
                If CoresArmadura = False Then
                    gr.DrawImage(My.Resources.ClaveCsus2, 84, 246 + AjusteArmaduraClaves, 57, 52)
                    If CheckBox3.Checked OrElse CheckBox4.Checked Then gr.DrawImage(My.Resources.ClaveCsus2, 84, 426, 57, 52)
                Else
                    gr.DrawImage(My.Resources.ClaveCsus2Col, 84, 246 + AjusteArmaduraClaves, 57, 52)
                    If CheckBox3.Checked OrElse CheckBox4.Checked Then gr.DrawImage(My.Resources.ClaveCsus2Col, 84, 426, 57, 52)
                End If
            ElseIf NomeDoMenu = "Cb - Abm (Bb - Eb - Ab - Db - Gb - Cb - Fb)" Then
                If CoresArmadura = False Then
                    gr.DrawImage(My.Resources.ClaveCbemol2, 84, 251 + AjusteArmaduraClaves, 56, 51)
                    If CheckBox3.Checked OrElse CheckBox4.Checked Then gr.DrawImage(My.Resources.ClaveCbemol2, 84, 431, 56, 51)
                Else
                    gr.DrawImage(My.Resources.ClaveCbemol2Col, 84, 251 + AjusteArmaduraClaves, 56, 51)
                    If CheckBox3.Checked OrElse CheckBox4.Checked Then gr.DrawImage(My.Resources.ClaveCbemol2Col, 84, 431, 56, 51)
                End If
            ElseIf NomeDoMenu = "Gb - Ebm (Bb - Eb - Ab - Db - Gb - Cb)" Then
                If CoresArmadura = False Then
                    gr.DrawImage(My.Resources.ClaveGbemol2, 84, 251 + AjusteArmaduraClaves, 56, 51)
                    If CheckBox3.Checked OrElse CheckBox4.Checked Then gr.DrawImage(My.Resources.ClaveGbemol2, 84, 431, 56, 51)
                Else
                    gr.DrawImage(My.Resources.ClaveGbemol2Col, 84, 251 + AjusteArmaduraClaves, 56, 51)
                    If CheckBox3.Checked OrElse CheckBox4.Checked Then gr.DrawImage(My.Resources.ClaveGbemol2Col, 84, 431, 56, 51)
                End If
            ElseIf NomeDoMenu = "Db - Bbm (Bb - Eb - Ab - Db - Gb)" Then
                If CoresArmadura = False Then
                    gr.DrawImage(My.Resources.ClaveDbemol2, 84, 251 + AjusteArmaduraClaves, 56, 51)
                    If CheckBox3.Checked OrElse CheckBox4.Checked Then gr.DrawImage(My.Resources.ClaveDbemol2, 84, 431, 56, 51)
                Else
                    gr.DrawImage(My.Resources.ClaveDbemol2Col, 84, 251 + AjusteArmaduraClaves, 56, 51)
                    If CheckBox3.Checked OrElse CheckBox4.Checked Then gr.DrawImage(My.Resources.ClaveDbemol2Col, 84, 431, 56, 51)
                End If
            ElseIf NomeDoMenu = "Ab - Fm (Bb - Eb - Ab - Db)" Then
                If CoresArmadura = False Then
                    gr.DrawImage(My.Resources.ClaveAbemol2, 84, 251 + AjusteArmaduraClaves, 56, 51)
                    If CheckBox3.Checked OrElse CheckBox4.Checked Then gr.DrawImage(My.Resources.ClaveAbemol2, 84, 431, 56, 51)
                Else
                    gr.DrawImage(My.Resources.ClaveAbemol2Col, 84, 251 + AjusteArmaduraClaves, 56, 51)
                    If CheckBox3.Checked OrElse CheckBox4.Checked Then gr.DrawImage(My.Resources.ClaveAbemol2Col, 84, 431, 56, 51)
                End If
            ElseIf NomeDoMenu = "Eb - Cm (Bb - Eb - Ab)" Then
                If CoresArmadura = False Then
                    gr.DrawImage(My.Resources.ClaveEbemol2, 84, 251 + AjusteArmaduraClaves, 56, 51)
                    If CheckBox3.Checked OrElse CheckBox4.Checked Then gr.DrawImage(My.Resources.ClaveEbemol2, 84, 431, 56, 51)
                Else
                    gr.DrawImage(My.Resources.ClaveEbemol2Col, 84, 251 + AjusteArmaduraClaves, 56, 51)
                    If CheckBox3.Checked OrElse CheckBox4.Checked Then gr.DrawImage(My.Resources.ClaveEbemol2Col, 84, 431, 56, 51)
                End If
            ElseIf NomeDoMenu = "Bb - Gm (Bb - Eb)" Then
                If CoresArmadura = False Then
                    gr.DrawImage(My.Resources.ClaveBbemol2, 84, 251 + AjusteArmaduraClaves, 56, 51)
                    If CheckBox3.Checked OrElse CheckBox4.Checked Then gr.DrawImage(My.Resources.ClaveBbemol2, 84, 431, 56, 51)
                Else
                    gr.DrawImage(My.Resources.ClaveBbemol2Col, 84, 251 + AjusteArmaduraClaves, 56, 51)
                    If CheckBox3.Checked OrElse CheckBox4.Checked Then gr.DrawImage(My.Resources.ClaveBbemol2Col, 84, 431, 56, 51)
                End If
            ElseIf NomeDoMenu = "F - Dm (Bb)" Then
                If CoresArmadura = False Then
                    gr.DrawImage(My.Resources.ClaveF2, 84, 251 + AjusteArmaduraClaves, 56, 51)
                    If CheckBox3.Checked OrElse CheckBox4.Checked Then gr.DrawImage(My.Resources.ClaveF2, 84, 431, 56, 51)
                Else
                    gr.DrawImage(My.Resources.ClaveF2Col, 84, 251 + AjusteArmaduraClaves, 56, 51)
                    If CheckBox3.Checked OrElse CheckBox4.Checked Then gr.DrawImage(My.Resources.ClaveF2Col, 84, 431, 56, 51)
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub RotinaAjusteDivisaoOitavas()

        Try

            ajusteDivisaoOitavas = 0
            DivisaoOitavas = Nothing
            If DóPrimeiraLinhaToolStripMenuItem1.Checked Then
                ajusteDivisaoOitavas = -10
            ElseIf DóSegundaLinhaToolStripMenuItem1.Checked Then
                ajusteDivisaoOitavas = -20
            ElseIf DóTerceiraLinhaToolStripMenuItem1.Checked Then
                ajusteDivisaoOitavas = -30
            ElseIf DóQuartaLinhaToolStripMenuItem1.Checked Then
                ajusteDivisaoOitavas = -40
            ElseIf DóQuintaLinhaToolStripMenuItem1.Checked OrElse FáTerceiraLinhaToolStripMenuItem1.Checked Then
                ajusteDivisaoOitavas = -50
            ElseIf FáQuintaLinhaToolStripMenuItem1.Checked Then
                ajusteDivisaoOitavas = -70
            ElseIf SolPrimeiraLinhaToolStripMenuItem1.Checked Then
                ajusteDivisaoOitavas = 10
            End If

            If My.Settings.NovoValorTransposição(0) = "False" AndAlso My.Settings.NovoValorTransposição(1) = "False" AndAlso My.Settings.NovoValorTransposição(2) = "False" AndAlso My.Settings.NovoValorTransposição(3) = "False" Then
                DivisaoOitavas = My.Resources.DivisaoDasOitavasNaPauta
            ElseIf My.Settings.NovoValorTransposição(0) = "True" Then
                DivisaoOitavas = My.Resources.DivisaoDasOitavasNaPauta8va
                If Not CheckBox4.Checked AndAlso Not CheckBox3.Checked Then ajusteDivisaoOitavas = ajusteDivisaoOitavas + 35
            ElseIf My.Settings.NovoValorTransposição(1) = "True" Then
                DivisaoOitavas = My.Resources.DivisaoDasOitavasNaPauta8vb
                If Not CheckBox4.Checked AndAlso Not CheckBox3.Checked Then ajusteDivisaoOitavas = ajusteDivisaoOitavas - 35
            ElseIf My.Settings.NovoValorTransposição(2) = "True" Then
                DivisaoOitavas = My.Resources.DivisaoDasOitavasNaPauta15va
                If Not CheckBox4.Checked AndAlso Not CheckBox3.Checked Then ajusteDivisaoOitavas = ajusteDivisaoOitavas + 70
            ElseIf My.Settings.NovoValorTransposição(3) = "True" Then
                DivisaoOitavas = My.Resources.DivisaoDasOitavasNaPauta15vb
                If Not CheckBox4.Checked AndAlso Not CheckBox3.Checked Then ajusteDivisaoOitavas = ajusteDivisaoOitavas - 70
            End If
            If Not CheckBox4.Checked AndAlso Not CheckBox3.Checked Then DivisaoOitavas = My.Resources.DivisaoDasOitavasNaPauta2

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub Notas_na_Pauta_Musical_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            NotaMidi = 1000
            MidiPlayer.OpenMidi()
        Catch ex As Exception
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não


            Thread(0) = New Thread(AddressOf SegundaThreadCode)
            Thread(0).Name = "Segunda Thread"
            Thread(0).Start()

            Minutos.Text = My.Settings.NovoValorTimers(0)
            Segundos.Text = My.Settings.NovoValorTimers(1)
            Decimos.Text = My.Settings.NovoValorTimers(2)

            Voz.Checked = CBool(My.Settings.NovoValorFonteCaptaçãoÁudio(0))
            Violão.Checked = CBool(My.Settings.NovoValorFonteCaptaçãoÁudio(1))



            If My.Settings.NovoValorLineInput(0) = "1" Then
                LineInputVoz1.Checked = True
            Else
                LineInputVoz2.Checked = True
            End If

            If My.Settings.NovoValorLineInput(1) = "1" Then
                LineInputViolão1.Checked = True
            Else
                LineInputViolão2.Checked = True
            End If



            Tocanota = 1
            NomeTecla(0) = ""

            NomeDoMenu = My.Settings.NovoValorNomeDoMenu
            If NomeDoMenu = "Nenhum" Then
                My.Settings.NovoValorCoresArmadura = False
                Acidentes.Enabled = True
                TomToolStripMenuItem.Image = My.Resources.ClaveNenhum
            Else
                Acidentes.Enabled = False
                If NomeDoMenu = "C - Am" Then
                    TomToolStripMenuItem.Image = My.Resources.ClaveC1
                ElseIf NomeDoMenu = "G - Em (F#)" Then
                    TomToolStripMenuItem.Image = My.Resources.ClaveG1
                ElseIf NomeDoMenu = "D - Bm (F# - C#)" Then
                    TomToolStripMenuItem.Image = My.Resources.ClaveD1
                ElseIf NomeDoMenu = "A - F#m (F# - C# - G#)" Then
                    TomToolStripMenuItem.Image = My.Resources.ClaveA1
                ElseIf NomeDoMenu = "E - C#m (F# - C# - G# - D#)" Then
                    TomToolStripMenuItem.Image = My.Resources.ClaveE1
                ElseIf NomeDoMenu = "B - G#m (F# - C# - G# - D# - A#)" Then
                    TomToolStripMenuItem.Image = My.Resources.ClaveB1
                ElseIf NomeDoMenu = "F# - D#m (F# - C# - G# - D# - A# - E#)" Then
                    TomToolStripMenuItem.Image = My.Resources.ClaveFsus1
                ElseIf NomeDoMenu = "C# - A#m (F# - C# - G# - D# - A# - E# - B#)" Then
                    TomToolStripMenuItem.Image = My.Resources.ClaveCsus1
                ElseIf NomeDoMenu = "Cb - Abm (Bb - Eb - Ab - Db - Gb - Cb - Fb)" Then
                    TomToolStripMenuItem.Image = My.Resources.ClaveCbemol1
                ElseIf NomeDoMenu = "Gb - Ebm (Bb - Eb - Ab - Db - Gb - Cb)" Then
                    TomToolStripMenuItem.Image = My.Resources.ClaveGbemol1
                ElseIf NomeDoMenu = "Db - Bbm (Bb - Eb - Ab - Db - Gb)" Then
                    TomToolStripMenuItem.Image = My.Resources.ClaveDbemol1
                ElseIf NomeDoMenu = "Ab - Fm (Bb - Eb - Ab - Db)" Then
                    TomToolStripMenuItem.Image = My.Resources.ClaveAbemol1
                ElseIf NomeDoMenu = "Eb - Cm (Bb - Eb - Ab)" Then
                    TomToolStripMenuItem.Image = My.Resources.ClaveEbemol1
                ElseIf NomeDoMenu = "Bb - Gm (Bb - Eb)" Then
                    TomToolStripMenuItem.Image = My.Resources.ClaveBbemol1
                ElseIf NomeDoMenu = "F - Dm (Bb)" Then
                    TomToolStripMenuItem.Image = My.Resources.ClaveF1
                End If
            End If

            n = 1
            vv = 0
            zz = 0
            'The color at Pixel(10,10) is rendered as transparent for the complete background.      
            'Img.MakeTransparent(Img.GetPixel(10, 10))
            'Me.BackgroundImage = Img
            'Me.TransparencyKey = Img.GetPixel(10, 10)
            'Me.BringToFront()

            ValorTotalDeNotas = 1
            DefineTecladoInicial()
            ExecutaLoopNotasAleatórias()

            Timer1.Interval = CInt(((CDbl(Minutos.Text) * 60) + CDbl(Segundos.Text) + (CDbl(Decimos.Text) / 100)) * 1000) 'Velocidade.Value * 1000


            GerarImagemDeFundo()
            GerarImagemDeFundo2()




            Dim iii As Short
            Dim Swk As UserControl
            For iii = 0 To 127
                Swk = New UserControl           ' Créé L'objet !
                With Swk
                    .Visible = False
                End With
                KeyCol.Add(Swk)             ' Rajoute un pointeur sur l'objet dans la collection
                Me.Controls.Add(Swk)        ' Rajoute un pointeur sur ME.controls
            Next iii

            '*** On unitialise le Midi In ***
            Dim MidiInCaps As New MIDIINCAPS
            Dim DrvNumber As Long

            For DrvNumber = 0 To (WinMM.midiInGetNumDevs - 1)            'on parcours tous les drivers
                WinMM.midiInGetDevCaps(CInt(DrvNumber), _
                                       MidiInCaps, _
                                       Marshal.SizeOf(MidiInCaps))
                Dim MenuItem As New ToolStripMenuItem
                MenuItem.Checked = False
                MenuItem.Tag = DrvNumber
                MenuItem.Text = Encoding.Unicode.GetString(MidiInCaps.ProductName)
                Me.Cms1.Items.Add(MenuItem)


                Dim midiError As Integer
                MenuItem.Checked = True 'testar para ver se já inicia com MIDI conectado
                ' On scanne le port Midi In
                midiError = WinMM.midiInOpen(hMidiIn, CInt(MenuItem.Tag), DelgMidiIn, 0, &H30000)
                midiError = WinMM.midiInStart(hMidiIn)

            Next
            VGhMidiIn = hMidiIn


            ComboBox1.Items.Add("Piano Bösendorfer")
            ComboBox1.Items.Add("Strings (Trio)")
            ComboBox1.Items.Add("000 - Acoustic Grand Piano")
            ComboBox1.Items.Add("001 - Bright Acoustic Piano")
            ComboBox1.Items.Add("002 - Electric Grand Piano")
            ComboBox1.Items.Add("003 - Honky-tonk Piano")
            ComboBox1.Items.Add("004 - Electric Piano 1")
            ComboBox1.Items.Add("005 - Electric Piano 2")
            ComboBox1.Items.Add("006 - Harpsichord")
            ComboBox1.Items.Add("007 - Clavi")
            ComboBox1.Items.Add("008 - Celesta")
            ComboBox1.Items.Add("009 - Glockenspiel")
            ComboBox1.Items.Add("010 - Music Box")
            ComboBox1.Items.Add("011 - Vibraphone")
            ComboBox1.Items.Add("012 - Marimba")
            ComboBox1.Items.Add("013 - Xylophone")
            ComboBox1.Items.Add("014 - Tubular Bells")
            ComboBox1.Items.Add("015 - Dulcimer")
            ComboBox1.Items.Add("016 - Drawbar Organ")
            ComboBox1.Items.Add("017 - Percussive Organ")
            ComboBox1.Items.Add("018 - Rock Organ")
            ComboBox1.Items.Add("019 - Church Organ")
            ComboBox1.Items.Add("020 - Reed Organ")
            ComboBox1.Items.Add("021 - Accordion")
            ComboBox1.Items.Add("022 - Harmonica")
            ComboBox1.Items.Add("023 - Tango Accordion")
            ComboBox1.Items.Add("024 - Acoustic Guitar (nylon)")
            ComboBox1.Items.Add("025 - Acoustic Guitar (steel)")
            ComboBox1.Items.Add("026 - Electric Guitar (jazz)")
            ComboBox1.Items.Add("027 - Electric Guitar (clean)")
            ComboBox1.Items.Add("028 - Electric Guitar (muted)")
            ComboBox1.Items.Add("029 - Overdriven Guitar")
            ComboBox1.Items.Add("030 - Distortion Guitar")
            ComboBox1.Items.Add("031 - Guitar harmonics")
            ComboBox1.Items.Add("032 - Acoustic Bass")
            ComboBox1.Items.Add("033 - Electric Bass (finger)")
            ComboBox1.Items.Add("034 - Electric Bass (pick)")
            ComboBox1.Items.Add("035 - Fretless Bass")
            ComboBox1.Items.Add("036 - Slap Bass 1")
            ComboBox1.Items.Add("037 - Slap Bass 2")
            ComboBox1.Items.Add("038 - Synth Bass 1")
            ComboBox1.Items.Add("039 - Synth Bass 2")
            ComboBox1.Items.Add("040 - Violin")
            ComboBox1.Items.Add("041 - Viola")
            ComboBox1.Items.Add("042 - Cello")
            ComboBox1.Items.Add("043 - Contrabass")
            ComboBox1.Items.Add("044 - Tremolo Strings")
            ComboBox1.Items.Add("045 - Pizzicato Strings")
            ComboBox1.Items.Add("046 - Orchestral Harp")
            ComboBox1.Items.Add("047 - Timpani")
            ComboBox1.Items.Add("048 - String Ensemble 1")
            ComboBox1.Items.Add("049 - String Ensemble 2")
            ComboBox1.Items.Add("050 - SynthStrings 1")
            ComboBox1.Items.Add("051 - SynthStrings 2")
            ComboBox1.Items.Add("052 - Choir Aahs")
            ComboBox1.Items.Add("053 - Voice Oohs")
            ComboBox1.Items.Add("054 - Synth Voice")
            ComboBox1.Items.Add("055 - Orchestra Hit")
            ComboBox1.Items.Add("056 - Trumpet")
            ComboBox1.Items.Add("057 - Trombone")
            ComboBox1.Items.Add("058 - Tuba")
            ComboBox1.Items.Add("059 - Muted Trumpet")
            ComboBox1.Items.Add("060 - French Horn")
            ComboBox1.Items.Add("061 - Brass Section")
            ComboBox1.Items.Add("062 - SynthBrass 1")
            ComboBox1.Items.Add("063 - SynthBrass 2")
            ComboBox1.Items.Add("064 - Soprano Sax")
            ComboBox1.Items.Add("065 - Alto Sax")
            ComboBox1.Items.Add("066 - Tenor Sax")
            ComboBox1.Items.Add("067 - Baritone Sax")
            ComboBox1.Items.Add("068 - Oboe")
            ComboBox1.Items.Add("069 - English Horn")
            ComboBox1.Items.Add("070 - Bassoon")
            ComboBox1.Items.Add("071 - Clarinet")
            ComboBox1.Items.Add("072 - Piccolo")
            ComboBox1.Items.Add("073 - Flute")
            ComboBox1.Items.Add("074 - Recorder")
            ComboBox1.Items.Add("075 - Pan Flute")
            ComboBox1.Items.Add("076 - Blown Bottle")
            ComboBox1.Items.Add("077 - Shakuhachi")
            ComboBox1.Items.Add("078 - Whistle")
            ComboBox1.Items.Add("079 - Ocarina")
            ComboBox1.Items.Add("080 - Lead 1 (square)")
            ComboBox1.Items.Add("081 - Lead 2 (sawtooth)")
            ComboBox1.Items.Add("082 - Lead 3 (calliope)")
            ComboBox1.Items.Add("083 - Lead 4 (chiff)")
            ComboBox1.Items.Add("084 - Lead 5 (charang)")
            ComboBox1.Items.Add("085 - Lead 6 (voice)")
            ComboBox1.Items.Add("086 - Lead 7 (fifths)")
            ComboBox1.Items.Add("087 - Lead 8 (bass + lead)")
            ComboBox1.Items.Add("088 - Pad 1 (new age)")
            ComboBox1.Items.Add("089 - Pad 2 (warm)")
            ComboBox1.Items.Add("090 - Pad 3 (polysynth)")
            ComboBox1.Items.Add("091 - Pad 4 (choir)")
            ComboBox1.Items.Add("092 - Pad 5 (bowed)")
            ComboBox1.Items.Add("093 - Pad 6 (metallic)")
            ComboBox1.Items.Add("094 - Pad 7 (halo)")
            ComboBox1.Items.Add("095 - Pad 8 (sweep)")
            ComboBox1.Items.Add("096 - FX 1 (rain)")
            ComboBox1.Items.Add("097 - FX 2 (soundtrack)")
            ComboBox1.Items.Add("098 - FX 3 (crystal)")
            ComboBox1.Items.Add("099 - FX 4 (atmosphere)")
            ComboBox1.Items.Add("100 - FX 5 (brightness)")
            ComboBox1.Items.Add("101 - FX 6 (goblins)")
            ComboBox1.Items.Add("102 - FX 7 (echoes)")
            ComboBox1.Items.Add("103 - FX 8 (sci-fi)")
            ComboBox1.Items.Add("104 - Sitar")
            ComboBox1.Items.Add("105 - Banjo")
            ComboBox1.Items.Add("106 - Shamisen")
            ComboBox1.Items.Add("107 - Koto")
            ComboBox1.Items.Add("108 - Kalimba")
            ComboBox1.Items.Add("109 - Bag pipe")
            ComboBox1.Items.Add("110 - Fiddle")
            ComboBox1.Items.Add("111 - Shanai")
            ComboBox1.Items.Add("112 - Tinkle Bell")
            ComboBox1.Items.Add("113 - Agogo")
            ComboBox1.Items.Add("114 - Steel Drums")
            ComboBox1.Items.Add("115 - Woodblock")
            ComboBox1.Items.Add("116 - Taiko Drum")
            ComboBox1.Items.Add("117 - Melodic Tom")
            ComboBox1.Items.Add("118 - Synth Drum")
            ComboBox1.Items.Add("119 - Reverse Cymbal")
            ComboBox1.Items.Add("120 - Guitar Fret Noise")
            ComboBox1.Items.Add("121 - Breath Noise")
            ComboBox1.Items.Add("122 - Seashore")
            ComboBox1.Items.Add("123 - Bird Tweet")
            ComboBox1.Items.Add("124 - Telephone Ring")
            ComboBox1.Items.Add("125 - Helicopter")
            ComboBox1.Items.Add("126 - Applause")
            ComboBox1.Items.Add("127 - Gunshot")

            ComboBox1.Text = My.Settings.NovoValorInstrumentoMusical(0)

            If ComboBox1.Text = "" Then ComboBox1.Text = "000 - Acoustic Grand Piano"


            ExibirLigadoDesligadoNomeDaNota()
        End Try

    End Sub

    Private Sub ab_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown, Pontuação.MouseDown, PictureBox14.MouseDown, Acertos.MouseDown, Erros.MouseDown, CheckBox3.MouseDown, CheckBox4.MouseDown, RadioButton1.MouseDown, RadioButton2.MouseDown, ProgressBar1.MouseDown
        a = Me.MousePosition.X - Me.Location.X
        b = Me.MousePosition.Y - Me.Location.Y
    End Sub

    Private Sub AtualizaSetasNaPautaETeclado()
        Try

            'atualiza ToolTipNota(0)
            Dim Rect8 As New Rectangle(479, 558, 104, 58)
            Me.Invalidate(Rect8)
            'atualiza Setas na Pauta e no Teclado
            Dim Rect9 As New Rectangle(50, 162, 180, 382)
            Me.Invalidate(Rect9)
            Dim Rect10 As New Rectangle(89, 707, 884, 7)
            Me.Invalidate(Rect10)
            dd = 0

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub ab_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove, Pontuação.MouseMove, PictureBox14.MouseMove, Acertos.MouseMove, Erros.MouseMove, CheckBox3.MouseMove, CheckBox4.MouseMove, RadioButton1.MouseMove, RadioButton2.MouseMove, ProgressBar1.MouseMove

        Try

            If TeclaControlPressionada = 1 Then
                Dim posição As Point = Me.PointToClient(Me.MousePosition)
                SetaPautaLeft = posição.X
                If dd = 1 Then
                    AtualizaSetasNaPautaETeclado()
                    dd = 0
                End If
                If RadioButton1.Checked Then
                    If (posição.X >= 96 AndAlso posição.X <= 216 AndAlso posição.Y >= 170 AndAlso posição.Y <= 349) OrElse (posição.X >= 96 AndAlso posição.X <= 216 AndAlso posição.Y >= 383 AndAlso posição.Y <= 539) Then

                        If My.Settings.NovoValorTransposição(0) = "True" Then
                            valorOitava(0) = 119 : valorOitava(1) = 1
                        ElseIf My.Settings.NovoValorTransposição(1) = "True" Then
                            valorOitava(0) = -119 : valorOitava(1) = -1
                        ElseIf My.Settings.NovoValorTransposição(2) = "True" Then
                            valorOitava(0) = 238 : valorOitava(1) = 2
                        ElseIf My.Settings.NovoValorTransposição(3) = "True" Then
                            valorOitava(0) = -238 : valorOitava(1) = -2
                        Else
                            valorOitava(0) = 0 : valorOitava(1) = 0
                        End If

                        If DóPrimeiraLinhaToolStripMenuItem1.Checked Then
                            Ajuste_NomeNotas(0) = "A"
                            Ajuste_NomeNotas(1) = "G"
                            Ajuste_NomeNotas(2) = "F"
                            Ajuste_NomeNotas(3) = "E"
                            Ajuste_NomeNotas(4) = "D"
                            Ajuste_NomeNotas(5) = "C"
                            Ajuste_NomeNotas(6) = "B"
                            AjusteNomeNotas = 34
                            ToolTipNota(6) = My.Resources.ToolTip_C_Bsus
                            ToolTipNota(7) = My.Resources.ToolTip_B_Cb
                            ToolTipNota(1) = My.Resources.ToolTip_A
                            ToolTipNota(2) = My.Resources.ToolTip_G
                            ToolTipNota(3) = My.Resources.ToolTip_F_Esus
                            ToolTipNota(4) = My.Resources.ToolTip_E_Fb
                            ToolTipNota(5) = My.Resources.ToolTip_D
                        ElseIf DóSegundaLinhaToolStripMenuItem1.Checked Then
                            Ajuste_NomeNotas(0) = "F"
                            Ajuste_NomeNotas(1) = "E"
                            Ajuste_NomeNotas(2) = "D"
                            Ajuste_NomeNotas(3) = "C"
                            Ajuste_NomeNotas(4) = "B"
                            Ajuste_NomeNotas(5) = "A"
                            Ajuste_NomeNotas(6) = "G"
                            AjusteNomeNotas = 68
                            ToolTipNota(4) = My.Resources.ToolTip_C_Bsus
                            ToolTipNota(5) = My.Resources.ToolTip_B_Cb
                            ToolTipNota(6) = My.Resources.ToolTip_A
                            ToolTipNota(7) = My.Resources.ToolTip_G
                            ToolTipNota(1) = My.Resources.ToolTip_F_Esus
                            ToolTipNota(2) = My.Resources.ToolTip_E_Fb
                            ToolTipNota(3) = My.Resources.ToolTip_D
                        ElseIf DóTerceiraLinhaToolStripMenuItem1.Checked Then
                            Ajuste_NomeNotas(0) = "D"
                            Ajuste_NomeNotas(1) = "C"
                            Ajuste_NomeNotas(2) = "B"
                            Ajuste_NomeNotas(3) = "A"
                            Ajuste_NomeNotas(4) = "G"
                            Ajuste_NomeNotas(5) = "F"
                            Ajuste_NomeNotas(6) = "E"
                            AjusteNomeNotas = 102
                            ToolTipNota(2) = My.Resources.ToolTip_C_Bsus
                            ToolTipNota(3) = My.Resources.ToolTip_B_Cb
                            ToolTipNota(4) = My.Resources.ToolTip_A
                            ToolTipNota(5) = My.Resources.ToolTip_G
                            ToolTipNota(6) = My.Resources.ToolTip_F_Esus
                            ToolTipNota(7) = My.Resources.ToolTip_E_Fb
                            ToolTipNota(1) = My.Resources.ToolTip_D
                        ElseIf DóQuartaLinhaToolStripMenuItem1.Checked Then
                            Ajuste_NomeNotas(0) = "B"
                            Ajuste_NomeNotas(1) = "A"
                            Ajuste_NomeNotas(2) = "G"
                            Ajuste_NomeNotas(3) = "F"
                            Ajuste_NomeNotas(4) = "E"
                            Ajuste_NomeNotas(5) = "D"
                            Ajuste_NomeNotas(6) = "C"
                            AjusteNomeNotas = 136
                            ToolTipNota(7) = My.Resources.ToolTip_C_Bsus
                            ToolTipNota(1) = My.Resources.ToolTip_B_Cb
                            ToolTipNota(2) = My.Resources.ToolTip_A
                            ToolTipNota(3) = My.Resources.ToolTip_G
                            ToolTipNota(4) = My.Resources.ToolTip_F_Esus
                            ToolTipNota(5) = My.Resources.ToolTip_E_Fb
                            ToolTipNota(6) = My.Resources.ToolTip_D
                        ElseIf FáTerceiraLinhaToolStripMenuItem1.Checked OrElse DóQuintaLinhaToolStripMenuItem1.Checked Then
                            Ajuste_NomeNotas(0) = "G"
                            Ajuste_NomeNotas(1) = "F"
                            Ajuste_NomeNotas(2) = "E"
                            Ajuste_NomeNotas(3) = "D"
                            Ajuste_NomeNotas(4) = "C"
                            Ajuste_NomeNotas(5) = "B"
                            Ajuste_NomeNotas(6) = "A"
                            AjusteNomeNotas = 170
                            ToolTipNota(5) = My.Resources.ToolTip_C_Bsus
                            ToolTipNota(6) = My.Resources.ToolTip_B_Cb
                            ToolTipNota(7) = My.Resources.ToolTip_A
                            ToolTipNota(1) = My.Resources.ToolTip_G
                            ToolTipNota(2) = My.Resources.ToolTip_F_Esus
                            ToolTipNota(3) = My.Resources.ToolTip_E_Fb
                            ToolTipNota(4) = My.Resources.ToolTip_D
                        ElseIf FáQuintaLinhaToolStripMenuItem1.Checked Then
                            Ajuste_NomeNotas(0) = "C"
                            Ajuste_NomeNotas(1) = "B"
                            Ajuste_NomeNotas(2) = "A"
                            Ajuste_NomeNotas(3) = "G"
                            Ajuste_NomeNotas(4) = "F"
                            Ajuste_NomeNotas(5) = "E"
                            Ajuste_NomeNotas(6) = "D"
                            AjusteNomeNotas = 238
                            ToolTipNota(1) = My.Resources.ToolTip_C_Bsus
                            ToolTipNota(2) = My.Resources.ToolTip_B_Cb
                            ToolTipNota(3) = My.Resources.ToolTip_A
                            ToolTipNota(4) = My.Resources.ToolTip_G
                            ToolTipNota(5) = My.Resources.ToolTip_F_Esus
                            ToolTipNota(6) = My.Resources.ToolTip_E_Fb
                            ToolTipNota(7) = My.Resources.ToolTip_D
                        ElseIf SolPrimeiraLinhaToolStripMenuItem1.Checked Then
                            Ajuste_NomeNotas(0) = "E"
                            Ajuste_NomeNotas(1) = "D"
                            Ajuste_NomeNotas(2) = "C"
                            Ajuste_NomeNotas(3) = "B"
                            Ajuste_NomeNotas(4) = "A"
                            Ajuste_NomeNotas(5) = "G"
                            Ajuste_NomeNotas(6) = "F"
                            AjusteNomeNotas = -34
                            ToolTipNota(3) = My.Resources.ToolTip_C_Bsus
                            ToolTipNota(4) = My.Resources.ToolTip_B_Cb
                            ToolTipNota(5) = My.Resources.ToolTip_A
                            ToolTipNota(6) = My.Resources.ToolTip_G
                            ToolTipNota(7) = My.Resources.ToolTip_F_Esus
                            ToolTipNota(1) = My.Resources.ToolTip_E_Fb
                            ToolTipNota(2) = My.Resources.ToolTip_D
                        Else
                            Ajuste_NomeNotas(0) = "C"
                            Ajuste_NomeNotas(1) = "B"
                            Ajuste_NomeNotas(2) = "A"
                            Ajuste_NomeNotas(3) = "G"
                            Ajuste_NomeNotas(4) = "F"
                            Ajuste_NomeNotas(5) = "E"
                            Ajuste_NomeNotas(6) = "D"
                            AjusteNomeNotas = 0
                            ToolTipNota(1) = My.Resources.ToolTip_C_Bsus
                            ToolTipNota(2) = My.Resources.ToolTip_B_Cb
                            ToolTipNota(3) = My.Resources.ToolTip_A
                            ToolTipNota(4) = My.Resources.ToolTip_G
                            ToolTipNota(5) = My.Resources.ToolTip_F_Esus
                            ToolTipNota(6) = My.Resources.ToolTip_E_Fb
                            ToolTipNota(7) = My.Resources.ToolTip_D
                        End If
                    End If

                    If posição.X >= 96 AndAlso posição.X <= 216 AndAlso posição.Y >= 170 AndAlso posição.Y <= 349 Then
                        If My.Settings.NovoValorTransposição(2) = "False" Then
                            If My.Settings.NovoValorTransposição(0) = "False" Then
                                If posição.Y >= 170 AndAlso posição.Y <= 172 Then
                                    ToolTipNota(0) = ToolTipNota(1) : SetaPautaTop = 167 : SetaTecladoLeft = 960 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(0)
                                ElseIf posição.Y >= 173 AndAlso posição.Y <= 179 Then
                                    ToolTipNota(0) = ToolTipNota(2) : SetaPautaTop = 172 : SetaTecladoLeft = 943 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(1)
                                ElseIf posição.Y >= 180 AndAlso posição.Y <= 182 Then
                                    ToolTipNota(0) = ToolTipNota(3) : SetaPautaTop = 177 : SetaTecladoLeft = 926 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(2)
                                ElseIf posição.Y >= 183 AndAlso posição.Y <= 189 Then
                                    ToolTipNota(0) = ToolTipNota(4) : SetaPautaTop = 182 : SetaTecladoLeft = 909 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(3)
                                ElseIf posição.Y >= 190 AndAlso posição.Y <= 192 Then
                                    ToolTipNota(0) = ToolTipNota(5) : SetaPautaTop = 187 : SetaTecladoLeft = 892 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(4)
                                ElseIf posição.Y >= 193 AndAlso posição.Y <= 199 Then
                                    ToolTipNota(0) = ToolTipNota(6) : SetaPautaTop = 192 : SetaTecladoLeft = 875 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(5)
                                ElseIf posição.Y >= 200 AndAlso posição.Y <= 202 Then
                                    ToolTipNota(0) = ToolTipNota(7) : SetaPautaTop = 197 : SetaTecladoLeft = 858 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(6)
                                End If
                            End If
                            If posição.Y >= 203 AndAlso posição.Y <= 209 Then
                                ToolTipNota(0) = ToolTipNota(1) : SetaPautaTop = 202 : SetaTecladoLeft = 841 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(0)
                            ElseIf posição.Y >= 210 AndAlso posição.Y <= 212 Then
                                ToolTipNota(0) = ToolTipNota(2) : SetaPautaTop = 207 : SetaTecladoLeft = 824 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(1)
                            ElseIf posição.Y >= 213 AndAlso posição.Y <= 219 Then
                                ToolTipNota(0) = ToolTipNota(3) : SetaPautaTop = 212 : SetaTecladoLeft = 807 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(2)
                            ElseIf posição.Y >= 220 AndAlso posição.Y <= 222 Then
                                ToolTipNota(0) = ToolTipNota(4) : SetaPautaTop = 217 : SetaTecladoLeft = 790 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(3)
                            ElseIf posição.Y >= 223 AndAlso posição.Y <= 229 Then
                                ToolTipNota(0) = ToolTipNota(5) : SetaPautaTop = 222 : SetaTecladoLeft = 773 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(4)
                            ElseIf posição.Y >= 230 AndAlso posição.Y <= 232 Then
                                ToolTipNota(0) = ToolTipNota(6) : SetaPautaTop = 227 : SetaTecladoLeft = 756 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(5)
                            ElseIf posição.Y >= 233 AndAlso posição.Y <= 239 Then
                                ToolTipNota(0) = ToolTipNota(7) : SetaPautaTop = 232 : SetaTecladoLeft = 739 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(6)
                            End If
                        End If
                        If posição.Y >= 240 AndAlso posição.Y <= 242 Then
                            ToolTipNota(0) = ToolTipNota(1) : SetaPautaTop = 237 : SetaTecladoLeft = 722 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(0)
                        ElseIf posição.Y >= 243 AndAlso posição.Y <= 249 Then
                            ToolTipNota(0) = ToolTipNota(2) : SetaPautaTop = 242 : SetaTecladoLeft = 705 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(1)
                        ElseIf posição.Y >= 250 AndAlso posição.Y <= 252 Then
                            ToolTipNota(0) = ToolTipNota(3) : SetaPautaTop = 247 : SetaTecladoLeft = 688 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(2)
                        ElseIf posição.Y >= 253 AndAlso posição.Y <= 259 Then
                            ToolTipNota(0) = ToolTipNota(4) : SetaPautaTop = 252 : SetaTecladoLeft = 671 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(3)
                        ElseIf posição.Y >= 260 AndAlso posição.Y <= 262 Then
                            ToolTipNota(0) = ToolTipNota(5) : SetaPautaTop = 257 : SetaTecladoLeft = 654 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(4)
                        ElseIf posição.Y >= 263 AndAlso posição.Y <= 269 Then
                            ToolTipNota(0) = ToolTipNota(6) : SetaPautaTop = 262 : SetaTecladoLeft = 637 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(5)
                        ElseIf posição.Y >= 270 AndAlso posição.Y <= 272 Then
                            ToolTipNota(0) = ToolTipNota(7) : SetaPautaTop = 267 : SetaTecladoLeft = 620 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(6)
                        ElseIf posição.Y >= 273 AndAlso posição.Y <= 279 Then
                            ToolTipNota(0) = ToolTipNota(1) : SetaPautaTop = 272 : SetaTecladoLeft = 603 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(0)
                        ElseIf posição.Y >= 280 AndAlso posição.Y <= 282 Then
                            ToolTipNota(0) = ToolTipNota(2) : SetaPautaTop = 277 : SetaTecladoLeft = 586 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(1)
                        ElseIf posição.Y >= 283 AndAlso posição.Y <= 289 Then
                            ToolTipNota(0) = ToolTipNota(3) : SetaPautaTop = 282 : SetaTecladoLeft = 569 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(2)
                        ElseIf posição.Y >= 290 AndAlso posição.Y <= 292 Then
                            ToolTipNota(0) = ToolTipNota(4) : SetaPautaTop = 287 : SetaTecladoLeft = 552 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(3)
                        ElseIf posição.Y >= 293 AndAlso posição.Y <= 299 Then
                            ToolTipNota(0) = ToolTipNota(5) : SetaPautaTop = 292 : SetaTecladoLeft = 535 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(4)
                        ElseIf posição.Y >= 300 AndAlso posição.Y <= 302 Then
                            ToolTipNota(0) = ToolTipNota(6) : SetaPautaTop = 297 : SetaTecladoLeft = 518 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(5)
                        ElseIf posição.Y >= 303 AndAlso posição.Y <= 309 Then
                            ToolTipNota(0) = ToolTipNota(7) : SetaPautaTop = 302 : SetaTecladoLeft = 501 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(6)
                        ElseIf posição.Y >= 310 AndAlso posição.Y <= 312 Then
                            ToolTipNota(0) = ToolTipNota(1) : SetaPautaTop = 307 : SetaTecladoLeft = 484 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(0)
                        ElseIf posição.Y >= 313 AndAlso posição.Y <= 319 Then
                            ToolTipNota(0) = ToolTipNota(2) : SetaPautaTop = 312 : SetaTecladoLeft = 467 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(1)
                        ElseIf posição.Y >= 320 AndAlso posição.Y <= 322 Then
                            ToolTipNota(0) = ToolTipNota(3) : SetaPautaTop = 317 : SetaTecladoLeft = 450 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(2)
                        ElseIf posição.Y >= 323 AndAlso posição.Y <= 329 Then
                            ToolTipNota(0) = ToolTipNota(4) : SetaPautaTop = 322 : SetaTecladoLeft = 433 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(3)
                        ElseIf posição.Y >= 330 AndAlso posição.Y <= 332 Then
                            ToolTipNota(0) = ToolTipNota(5) : SetaPautaTop = 327 : SetaTecladoLeft = 416 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(4)
                        ElseIf posição.Y >= 333 AndAlso posição.Y <= 339 Then
                            ToolTipNota(0) = ToolTipNota(6) : SetaPautaTop = 332 : SetaTecladoLeft = 399 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(5)
                        ElseIf posição.Y >= 340 AndAlso posição.Y <= 342 Then
                            ToolTipNota(0) = ToolTipNota(7) : SetaPautaTop = 337 : SetaTecladoLeft = 382 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(6)
                        ElseIf posição.Y >= 343 AndAlso posição.Y <= 349 Then
                            ToolTipNota(0) = ToolTipNota(1) : SetaPautaTop = 342 : SetaTecladoLeft = 365 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(0)
                        End If

                    ElseIf posição.X >= 96 AndAlso posição.X <= 216 AndAlso posição.Y >= 383 AndAlso posição.Y <= 539 Then
                        If CheckBox3.Checked OrElse CheckBox4.Checked Then
                            If posição.Y >= 383 AndAlso posição.Y <= 389 Then
                                ToolTipNota(0) = ToolTipNota(1) : SetaPautaTop = 382 : SetaTecladoLeft = 603 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(0)
                            ElseIf posição.Y >= 390 AndAlso posição.Y <= 392 Then
                                ToolTipNota(0) = ToolTipNota(2) : SetaPautaTop = 387 : SetaTecladoLeft = 586 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(1)
                            ElseIf posição.Y >= 393 AndAlso posição.Y <= 399 Then
                                ToolTipNota(0) = ToolTipNota(3) : SetaPautaTop = 392 : SetaTecladoLeft = 569 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(2)
                            ElseIf posição.Y >= 400 AndAlso posição.Y <= 402 Then
                                ToolTipNota(0) = ToolTipNota(4) : SetaPautaTop = 397 : SetaTecladoLeft = 552 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(3)
                            ElseIf posição.Y >= 403 AndAlso posição.Y <= 409 Then
                                ToolTipNota(0) = ToolTipNota(5) : SetaPautaTop = 402 : SetaTecladoLeft = 535 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(4)
                            ElseIf posição.Y >= 410 AndAlso posição.Y <= 412 Then
                                ToolTipNota(0) = ToolTipNota(6) : SetaPautaTop = 407 : SetaTecladoLeft = 518 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(5)
                            ElseIf posição.Y >= 413 AndAlso posição.Y <= 419 Then
                                ToolTipNota(0) = ToolTipNota(7) : SetaPautaTop = 412 : SetaTecladoLeft = 501 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(6)
                            ElseIf posição.Y >= 420 AndAlso posição.Y <= 422 Then
                                ToolTipNota(0) = ToolTipNota(1) : SetaPautaTop = 417 : SetaTecladoLeft = 484 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(0)
                            ElseIf posição.Y >= 423 AndAlso posição.Y <= 429 Then
                                ToolTipNota(0) = ToolTipNota(2) : SetaPautaTop = 422 : SetaTecladoLeft = 467 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(1)
                            ElseIf posição.Y >= 430 AndAlso posição.Y <= 432 Then
                                ToolTipNota(0) = ToolTipNota(3) : SetaPautaTop = 427 : SetaTecladoLeft = 450 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(2)
                            ElseIf posição.Y >= 433 AndAlso posição.Y <= 439 Then
                                ToolTipNota(0) = ToolTipNota(4) : SetaPautaTop = 432 : SetaTecladoLeft = 433 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(3)
                            ElseIf posição.Y >= 440 AndAlso posição.Y <= 442 Then
                                ToolTipNota(0) = ToolTipNota(5) : SetaPautaTop = 437 : SetaTecladoLeft = 416 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(4)
                            ElseIf posição.Y >= 443 AndAlso posição.Y <= 449 Then
                                ToolTipNota(0) = ToolTipNota(6) : SetaPautaTop = 442 : SetaTecladoLeft = 399 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(5)
                            ElseIf posição.Y >= 450 AndAlso posição.Y <= 452 Then
                                ToolTipNota(0) = ToolTipNota(7) : SetaPautaTop = 447 : SetaTecladoLeft = 382 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(6)
                            ElseIf posição.Y >= 453 AndAlso posição.Y <= 459 Then
                                ToolTipNota(0) = ToolTipNota(1) : SetaPautaTop = 452 : SetaTecladoLeft = 365 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(0)
                            ElseIf posição.Y >= 460 AndAlso posição.Y <= 462 Then
                                ToolTipNota(0) = ToolTipNota(2) : SetaPautaTop = 457 : SetaTecladoLeft = 348 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(1)
                            ElseIf posição.Y >= 463 AndAlso posição.Y <= 469 Then
                                ToolTipNota(0) = ToolTipNota(3) : SetaPautaTop = 462 : SetaTecladoLeft = 331 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(2)
                            End If
                            If My.Settings.NovoValorTransposição(3) = "False" Then
                                If posição.Y >= 470 AndAlso posição.Y <= 472 Then
                                    ToolTipNota(0) = ToolTipNota(4) : SetaPautaTop = 467 : SetaTecladoLeft = 314 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(3)
                                ElseIf posição.Y >= 473 AndAlso posição.Y <= 479 Then
                                    ToolTipNota(0) = ToolTipNota(5) : SetaPautaTop = 472 : SetaTecladoLeft = 297 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(4)
                                ElseIf posição.Y >= 480 AndAlso posição.Y <= 482 Then
                                    ToolTipNota(0) = ToolTipNota(6) : SetaPautaTop = 477 : SetaTecladoLeft = 280 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(5)
                                ElseIf posição.Y >= 483 AndAlso posição.Y <= 489 Then
                                    ToolTipNota(0) = ToolTipNota(7) : SetaPautaTop = 482 : SetaTecladoLeft = 263 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(6)
                                ElseIf posição.Y >= 490 AndAlso posição.Y <= 492 Then
                                    ToolTipNota(0) = ToolTipNota(1) : SetaPautaTop = 487 : SetaTecladoLeft = 246 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(0)
                                ElseIf posição.Y >= 493 AndAlso posição.Y <= 499 Then
                                    ToolTipNota(0) = ToolTipNota(2) : SetaPautaTop = 492 : SetaTecladoLeft = 229 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(1)
                                ElseIf posição.Y >= 500 AndAlso posição.Y <= 502 Then
                                    ToolTipNota(0) = ToolTipNota(3) : SetaPautaTop = 497 : SetaTecladoLeft = 212 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(2)
                                End If
                                If My.Settings.NovoValorTransposição(1) = "False" Then
                                    If posição.Y >= 503 AndAlso posição.Y <= 509 Then
                                        ToolTipNota(0) = ToolTipNota(4) : SetaPautaTop = 502 : SetaTecladoLeft = 195 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(3)
                                    ElseIf posição.Y >= 510 AndAlso posição.Y <= 512 Then
                                        ToolTipNota(0) = ToolTipNota(5) : SetaPautaTop = 507 : SetaTecladoLeft = 178 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(4)
                                    ElseIf posição.Y >= 513 AndAlso posição.Y <= 519 Then
                                        ToolTipNota(0) = ToolTipNota(6) : SetaPautaTop = 512 : SetaTecladoLeft = 161 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(5)
                                    ElseIf posição.Y >= 520 AndAlso posição.Y <= 522 Then
                                        ToolTipNota(0) = ToolTipNota(7) : SetaPautaTop = 517 : SetaTecladoLeft = 144 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(6)
                                    ElseIf posição.Y >= 523 AndAlso posição.Y <= 529 Then
                                        ToolTipNota(0) = ToolTipNota(1) : SetaPautaTop = 522 : SetaTecladoLeft = 127 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(0)
                                    ElseIf posição.Y >= 530 AndAlso posição.Y <= 532 Then
                                        ToolTipNota(0) = ToolTipNota(2) : SetaPautaTop = 527 : SetaTecladoLeft = 110 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(1)
                                    ElseIf posição.Y >= 533 AndAlso posição.Y <= 539 Then
                                        ToolTipNota(0) = ToolTipNota(3) : SetaPautaTop = 532 : SetaTecladoLeft = 93 + valorOitava(0) - AjusteNomeNotas : RetanguloBranco = Ajuste_NomeNotas(2)
                                    End If
                                End If
                            End If

                        End If

                    Else
                        ToolTipNota(0) = Nothing
                        SetaPautaTop = 0 : SetaTecladoLeft = 0 : RetanguloBranco = ""
                    End If
                End If

                If (posição.X >= 96 AndAlso posição.X <= 216 AndAlso posição.Y >= 170 AndAlso posição.Y <= 349) OrElse (posição.X >= 96 AndAlso posição.X <= 216 AndAlso posição.Y >= 383 AndAlso posição.Y <= 539) Then
                    AtualizaSetasNaPautaETeclado()
                    dd = 1
                End If
            End If


            If e.Button = MouseButtons.Left Then
                newPoint = Me.MousePosition
                newPoint.X = newPoint.X - a
                newPoint.Y = newPoint.Y - b
                Me.Location = newPoint
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Try

            If RadioButton2.Checked Then
                ValorTotalDeNotas = CInt(TextBox1.Text)
                DefineValorCorretoIncorreto()
                MostraNotaCorreta()
                ExibePontuação()
            Else
                ValorTotalDeNotas = 1
            End If

            ExecutaLoopNotasAleatórias()
            Timer1.Interval = CInt(((CDbl(Minutos.Text) * 60) + CDbl(Segundos.Text) + (CDbl(Decimos.Text) / 100)) * 1000)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub ExecutaLoopNotasAleatórias()

        Try

            NumeroDeNotas(1) = 0
            Nota(0) = 0 : Nota(1) = 0 : Nota(2) = 0 : Nota(3) = 0

            Do While NumeroDeNotas(1) < ValorTotalDeNotas
                Gera_Nota_Aleatoria()
                Nota(NumeroDeNotas(1)) = k
                LetraCifraFinal(NumeroDeNotas(1)) = LetraCifra
                AjusteDasNotasNaClaveFinal(NumeroDeNotas(1)) = AjusteDasNotasNaClave
                mFinal(NumeroDeNotas(1)) = m
                kkFinal(NumeroDeNotas(1)) = kk

                If NumeroDeNotas(1) = 1 AndAlso (Nota(1) = Nota(0) OrElse (Nota(1) > Nota(0) + QtdeNotas.Value OrElse Nota(1) < Nota(0) - QtdeNotas.Value)) Then
                    NumeroDeNotas(1) -= 1
                ElseIf NumeroDeNotas(1) = 2 AndAlso (Nota(2) = Nota(1) OrElse Nota(2) = Nota(0) OrElse (Nota(2) > Nota(0) + QtdeNotas.Value OrElse Nota(2) < Nota(0) - QtdeNotas.Value)) Then
                    NumeroDeNotas(1) -= 1
                ElseIf NumeroDeNotas(1) = 3 AndAlso (Nota(3) = Nota(2) OrElse Nota(3) = Nota(1) OrElse Nota(3) = Nota(0) OrElse (Nota(3) > Nota(0) + QtdeNotas.Value OrElse Nota(3) < Nota(0) - QtdeNotas.Value)) Then
                    NumeroDeNotas(1) -= 1
                End If

                NumeroDeNotas(1) += 1
            Loop

            If CifrasDeAcordes.Checked Then GerarAcordeAleatório()

            'atualiza notas na tela
            'Dim Rect2 As New Rectangle(121, 130, 84, 433)
            'Me.Invalidate(Rect2)
            'Dim Rect3 As New Rectangle(350, 260, 300, 170)
            'Me.Invalidate(Rect3)

            EventoPaintGeradoPelo_ExecutaLoopNotasAleatórias = True
            GerarImagemDeFundo2()
            EventoPaintGeradoPelo_ExecutaLoopNotasAleatórias = False

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub GerarAcordeAleatório()

        Try

            CifraFundamentalDoAcorde = ""


            'valor das notas deve ficar entre 1 e 12... 
            FundamentalDoAcorde = 100
            Do While FundamentalDoAcorde > 12 OrElse FundamentalDoAcorde < 1

                ' Create a byte array to hold the random value.
                Dim randomNumber(0) As Byte

                ' Create a new instance of the RNGCryptoServiceProvider.
                Dim Gen As New Security.Cryptography.RNGCryptoServiceProvider()

                ' Fill the array with a random value.
                Gen.GetBytes(randomNumber)

                ' Convert the byte to an integer value to make the modulus operation easier.
                FundamentalDoAcorde = Int(Convert.ToInt32(randomNumber(0)))

            Loop

            'sustenido ou bemol 
            AcordeSustenidoOuBemol = 100
            Do While AcordeSustenidoOuBemol > 2 OrElse AcordeSustenidoOuBemol < 1

                ' Create a byte array to hold the random value.
                Dim randomNumber(0) As Byte

                ' Create a new instance of the RNGCryptoServiceProvider.
                Dim Gen As New Security.Cryptography.RNGCryptoServiceProvider()

                ' Fill the array with a random value.
                Gen.GetBytes(randomNumber)

                ' Convert the byte to an integer value to make the modulus operation easier.
                AcordeSustenidoOuBemol = Int(Convert.ToInt32(randomNumber(0)))

            Loop

            If FundamentalDoAcorde = 1 Then
                CifraFundamentalDoAcorde = "A"
            ElseIf FundamentalDoAcorde = 2 Then
                If AcordeSustenidoOuBemol = 1 Then
                    CifraFundamentalDoAcorde = "A#"
                Else
                    CifraFundamentalDoAcorde = "Bb"
                End If
            ElseIf FundamentalDoAcorde = 3 Then
                CifraFundamentalDoAcorde = "B"
            ElseIf FundamentalDoAcorde = 4 Then
                CifraFundamentalDoAcorde = "C"
            ElseIf FundamentalDoAcorde = 5 Then
                If AcordeSustenidoOuBemol = 1 Then
                    CifraFundamentalDoAcorde = "C#"
                Else
                    CifraFundamentalDoAcorde = "Db"
                End If
            ElseIf FundamentalDoAcorde = 6 Then
                CifraFundamentalDoAcorde = "D"
            ElseIf FundamentalDoAcorde = 7 Then
                If AcordeSustenidoOuBemol = 1 Then
                    CifraFundamentalDoAcorde = "D#"
                Else
                    CifraFundamentalDoAcorde = "Eb"
                End If
            ElseIf FundamentalDoAcorde = 8 Then
                CifraFundamentalDoAcorde = "E"
            ElseIf FundamentalDoAcorde = 9 Then
                CifraFundamentalDoAcorde = "F"
            ElseIf FundamentalDoAcorde = 10 Then
                If AcordeSustenidoOuBemol = 1 Then
                    CifraFundamentalDoAcorde = "F#"
                Else
                    CifraFundamentalDoAcorde = "Gb"
                End If
            ElseIf FundamentalDoAcorde = 11 Then
                CifraFundamentalDoAcorde = "G"
            ElseIf FundamentalDoAcorde = 12 Then
                If AcordeSustenidoOuBemol = 1 Then
                    CifraFundamentalDoAcorde = "G#"
                Else
                    CifraFundamentalDoAcorde = "Ab"
                End If
            End If

            'Gera num aleatorio para definir a familia de acorde
            FamiliaDoAcorde = 100
            Do While FamiliaDoAcorde > 41 OrElse FamiliaDoAcorde < 1

                ' Create a byte array to hold the random value.
                Dim randomNumber(0) As Byte

                ' Create a new instance of the RNGCryptoServiceProvider.
                Dim Gen As New Security.Cryptography.RNGCryptoServiceProvider()

                ' Fill the array with a random value.
                Gen.GetBytes(randomNumber)

                ' Convert the byte to an integer value to make the modulus operation easier.
                FamiliaDoAcorde = Int(Convert.ToInt32(randomNumber(0)))

                If FamiliaDoAcorde > 0 AndAlso FamiliaDoAcorde < 42 AndAlso My.Settings.NovoValorCifras(FamiliaDoAcorde - 1) = "False" Then FamiliaDoAcorde = 100

            Loop

            TextBox1.Text = CStr(4)
            If FamiliaDoAcorde = 1 Then 'X
                kkFinal(0) = FundamentalDoAcorde : kkFinal(1) = FundamentalDoAcorde + 4 : kkFinal(2) = FundamentalDoAcorde + 7 : TextBox1.Text = CStr(3)
                NomeAcorde = CifraFundamentalDoAcorde
            ElseIf FamiliaDoAcorde = 2 Then 'Xm
                kkFinal(0) = FundamentalDoAcorde : kkFinal(1) = FundamentalDoAcorde + 3 : kkFinal(2) = FundamentalDoAcorde + 7 : TextBox1.Text = CStr(3)
                NomeAcorde = CifraFundamentalDoAcorde & "m"
            ElseIf FamiliaDoAcorde = 3 Then 'Xº
                kkFinal(0) = FundamentalDoAcorde : kkFinal(1) = FundamentalDoAcorde + 3 : kkFinal(2) = FundamentalDoAcorde + 6 : kkFinal(3) = FundamentalDoAcorde + 9
                NomeAcorde = CifraFundamentalDoAcorde & "º"
            ElseIf FamiliaDoAcorde = 4 Then 'X4
                kkFinal(0) = FundamentalDoAcorde + 10 : kkFinal(1) = FundamentalDoAcorde + 2 : kkFinal(2) = FundamentalDoAcorde + 5 : kkFinal(3) = FundamentalDoAcorde + 7
                NomeAcorde = CifraFundamentalDoAcorde & "4"
            ElseIf FamiliaDoAcorde = 5 Then 'X(#5)
                kkFinal(0) = FundamentalDoAcorde : kkFinal(1) = FundamentalDoAcorde + 4 : kkFinal(2) = FundamentalDoAcorde + 8 : TextBox1.Text = CStr(3)
                NomeAcorde = CifraFundamentalDoAcorde & "(#5)"
            ElseIf FamiliaDoAcorde = 6 Then 'X6
                kkFinal(0) = FundamentalDoAcorde : kkFinal(1) = FundamentalDoAcorde + 4 : kkFinal(2) = FundamentalDoAcorde + 7 : kkFinal(3) = FundamentalDoAcorde + 9
                NomeAcorde = CifraFundamentalDoAcorde & "6"
            ElseIf FamiliaDoAcorde = 7 Then 'X6/9
                kkFinal(0) = FundamentalDoAcorde + 2 : kkFinal(1) = FundamentalDoAcorde + 4 : kkFinal(2) = FundamentalDoAcorde + 7 : kkFinal(3) = FundamentalDoAcorde + 9
                NomeAcorde = CifraFundamentalDoAcorde & "6/9"
            ElseIf FamiliaDoAcorde = 8 Then 'Xm6
                kkFinal(0) = FundamentalDoAcorde : kkFinal(1) = FundamentalDoAcorde + 3 : kkFinal(2) = FundamentalDoAcorde + 7 : kkFinal(3) = FundamentalDoAcorde + 9
                NomeAcorde = CifraFundamentalDoAcorde & "m6"
            ElseIf FamiliaDoAcorde = 9 Then 'Xm6/9
                kkFinal(0) = FundamentalDoAcorde + 2 : kkFinal(1) = FundamentalDoAcorde + 3 : kkFinal(2) = FundamentalDoAcorde + 7 : kkFinal(3) = FundamentalDoAcorde + 9
                NomeAcorde = CifraFundamentalDoAcorde & "m6/9"
            ElseIf FamiliaDoAcorde = 10 Then 'X7
                kkFinal(0) = FundamentalDoAcorde : kkFinal(1) = FundamentalDoAcorde + 4 : kkFinal(2) = FundamentalDoAcorde + 7 : kkFinal(3) = FundamentalDoAcorde + 10
                NomeAcorde = CifraFundamentalDoAcorde & "7"
            ElseIf FamiliaDoAcorde = 11 Then 'X7(b5)
                kkFinal(0) = FundamentalDoAcorde : kkFinal(1) = FundamentalDoAcorde + 4 : kkFinal(2) = FundamentalDoAcorde + 6 : kkFinal(3) = FundamentalDoAcorde + 10
                NomeAcorde = CifraFundamentalDoAcorde & "7(b5)"
            ElseIf FamiliaDoAcorde = 12 Then 'X7(#5)
                kkFinal(0) = FundamentalDoAcorde : kkFinal(1) = FundamentalDoAcorde + 4 : kkFinal(2) = FundamentalDoAcorde + 8 : kkFinal(3) = FundamentalDoAcorde + 10
                NomeAcorde = CifraFundamentalDoAcorde & "7(#5)"
            ElseIf FamiliaDoAcorde = 13 Then 'X7(9)
                kkFinal(0) = FundamentalDoAcorde + 2 : kkFinal(1) = FundamentalDoAcorde + 4 : kkFinal(2) = FundamentalDoAcorde + 7 : kkFinal(3) = FundamentalDoAcorde + 10
                NomeAcorde = CifraFundamentalDoAcorde & "7(9)"
            ElseIf FamiliaDoAcorde = 14 Then 'X7(b9)
                kkFinal(0) = FundamentalDoAcorde + 1 : kkFinal(1) = FundamentalDoAcorde + 4 : kkFinal(2) = FundamentalDoAcorde + 7 : kkFinal(3) = FundamentalDoAcorde + 10
                NomeAcorde = CifraFundamentalDoAcorde & "7(b9)"
            ElseIf FamiliaDoAcorde = 15 Then 'X7(#9)
                kkFinal(0) = FundamentalDoAcorde + 3 : kkFinal(1) = FundamentalDoAcorde + 4 : kkFinal(2) = FundamentalDoAcorde + 7 : kkFinal(3) = FundamentalDoAcorde + 10
                NomeAcorde = CifraFundamentalDoAcorde & "7(#9)"
            ElseIf FamiliaDoAcorde = 16 Then 'X7(13)
                kkFinal(0) = FundamentalDoAcorde + 10 : kkFinal(1) = FundamentalDoAcorde + 2 : kkFinal(2) = FundamentalDoAcorde + 4 : kkFinal(3) = FundamentalDoAcorde + 9
                NomeAcorde = CifraFundamentalDoAcorde & "7(13)"
            ElseIf FamiliaDoAcorde = 17 Then 'X7(b13)
                kkFinal(0) = FundamentalDoAcorde + 10 : kkFinal(1) = FundamentalDoAcorde + 2 : kkFinal(2) = FundamentalDoAcorde + 4 : kkFinal(3) = FundamentalDoAcorde + 8
                NomeAcorde = CifraFundamentalDoAcorde & "7(b13)"
            ElseIf FamiliaDoAcorde = 18 Then 'X7(#11)
                kkFinal(0) = FundamentalDoAcorde : kkFinal(1) = FundamentalDoAcorde + 4 : kkFinal(2) = FundamentalDoAcorde + 6 : kkFinal(3) = FundamentalDoAcorde + 10
                NomeAcorde = CifraFundamentalDoAcorde & "7(#11)"
            ElseIf FamiliaDoAcorde = 19 Then 'X7(13/b9)
                kkFinal(0) = FundamentalDoAcorde + 10 : kkFinal(1) = FundamentalDoAcorde + 1 : kkFinal(2) = FundamentalDoAcorde + 4 : kkFinal(3) = FundamentalDoAcorde + 9
                NomeAcorde = CifraFundamentalDoAcorde & "7(13/b9)"
            ElseIf FamiliaDoAcorde = 20 Then 'X7(b13/b9)
                kkFinal(0) = FundamentalDoAcorde + 10 : kkFinal(1) = FundamentalDoAcorde + 1 : kkFinal(2) = FundamentalDoAcorde + 4 : kkFinal(3) = FundamentalDoAcorde + 8
                NomeAcorde = CifraFundamentalDoAcorde & "7(b13/b9)"
            ElseIf FamiliaDoAcorde = 21 Then 'X7/4
                kkFinal(0) = FundamentalDoAcorde : kkFinal(1) = FundamentalDoAcorde + 5 : kkFinal(2) = FundamentalDoAcorde + 7 : kkFinal(3) = FundamentalDoAcorde + 10
                NomeAcorde = CifraFundamentalDoAcorde & "7/4"
            ElseIf FamiliaDoAcorde = 22 Then 'X7/4(9)
                kkFinal(0) = FundamentalDoAcorde + 2 : kkFinal(1) = FundamentalDoAcorde + 5 : kkFinal(2) = FundamentalDoAcorde + 7 : kkFinal(3) = FundamentalDoAcorde + 10
                NomeAcorde = CifraFundamentalDoAcorde & "7/4(9)"
            ElseIf FamiliaDoAcorde = 23 Then 'X7/4(b9)
                kkFinal(0) = FundamentalDoAcorde + 1 : kkFinal(1) = FundamentalDoAcorde + 5 : kkFinal(2) = FundamentalDoAcorde + 7 : kkFinal(3) = FundamentalDoAcorde + 10
                NomeAcorde = CifraFundamentalDoAcorde & "7/4(b9)"
            ElseIf FamiliaDoAcorde = 24 Then 'Xm7
                kkFinal(0) = FundamentalDoAcorde : kkFinal(1) = FundamentalDoAcorde + 3 : kkFinal(2) = FundamentalDoAcorde + 7 : kkFinal(3) = FundamentalDoAcorde + 10
                NomeAcorde = CifraFundamentalDoAcorde & "m7"
            ElseIf FamiliaDoAcorde = 25 Then 'Xm7(11)
                kkFinal(0) = FundamentalDoAcorde : kkFinal(1) = FundamentalDoAcorde + 3 : kkFinal(2) = FundamentalDoAcorde + 5 : kkFinal(3) = FundamentalDoAcorde + 10
                NomeAcorde = CifraFundamentalDoAcorde & "m7(11)"
            ElseIf FamiliaDoAcorde = 26 Then 'Xm7(b5)
                kkFinal(0) = FundamentalDoAcorde : kkFinal(1) = FundamentalDoAcorde + 3 : kkFinal(2) = FundamentalDoAcorde + 6 : kkFinal(3) = FundamentalDoAcorde + 10
                NomeAcorde = CifraFundamentalDoAcorde & "m7(b5)"
            ElseIf FamiliaDoAcorde = 27 Then 'Xm7(9)
                kkFinal(0) = FundamentalDoAcorde + 2 : kkFinal(1) = FundamentalDoAcorde + 3 : kkFinal(2) = FundamentalDoAcorde + 7 : kkFinal(3) = FundamentalDoAcorde + 10
                NomeAcorde = CifraFundamentalDoAcorde & "m7(9)"
            ElseIf FamiliaDoAcorde = 28 Then 'Xm(7M)
                kkFinal(0) = FundamentalDoAcorde : kkFinal(1) = FundamentalDoAcorde + 3 : kkFinal(2) = FundamentalDoAcorde + 7 : kkFinal(3) = FundamentalDoAcorde + 11
                NomeAcorde = CifraFundamentalDoAcorde & "m(7M)"
            ElseIf FamiliaDoAcorde = 29 Then 'Xm(7M/6)
                kkFinal(0) = FundamentalDoAcorde : kkFinal(1) = FundamentalDoAcorde + 3 : kkFinal(2) = FundamentalDoAcorde + 9 : kkFinal(3) = FundamentalDoAcorde + 11
                NomeAcorde = CifraFundamentalDoAcorde & "m(7M/6)"
            ElseIf FamiliaDoAcorde = 30 Then 'Xm(7M/9)
                kkFinal(0) = FundamentalDoAcorde : kkFinal(1) = FundamentalDoAcorde + 2 : kkFinal(2) = FundamentalDoAcorde + 3 : kkFinal(3) = FundamentalDoAcorde + 11
                NomeAcorde = CifraFundamentalDoAcorde & "m(7M/9)"
            ElseIf FamiliaDoAcorde = 31 Then 'X7M
                kkFinal(0) = FundamentalDoAcorde : kkFinal(1) = FundamentalDoAcorde + 4 : kkFinal(2) = FundamentalDoAcorde + 7 : kkFinal(3) = FundamentalDoAcorde + 11
                NomeAcorde = CifraFundamentalDoAcorde & "7M"
            ElseIf FamiliaDoAcorde = 32 Then 'X7M(9)
                kkFinal(0) = FundamentalDoAcorde + 2 : kkFinal(1) = FundamentalDoAcorde + 4 : kkFinal(2) = FundamentalDoAcorde + 7 : kkFinal(3) = FundamentalDoAcorde + 11
                NomeAcorde = CifraFundamentalDoAcorde & "7M(9)"
            ElseIf FamiliaDoAcorde = 33 Then 'X7M(#5)
                kkFinal(0) = FundamentalDoAcorde : kkFinal(1) = FundamentalDoAcorde + 4 : kkFinal(2) = FundamentalDoAcorde + 8 : kkFinal(3) = FundamentalDoAcorde + 11
                NomeAcorde = CifraFundamentalDoAcorde & "7M(#5)"
            ElseIf FamiliaDoAcorde = 34 Then 'X7M(6)
                kkFinal(0) = FundamentalDoAcorde : kkFinal(1) = FundamentalDoAcorde + 4 : kkFinal(2) = FundamentalDoAcorde + 9 : kkFinal(3) = FundamentalDoAcorde + 11
                NomeAcorde = CifraFundamentalDoAcorde & "7M(6)"
            ElseIf FamiliaDoAcorde = 35 Then 'X7M(#11)
                kkFinal(0) = FundamentalDoAcorde : kkFinal(1) = FundamentalDoAcorde + 4 : kkFinal(2) = FundamentalDoAcorde + 6 : kkFinal(3) = FundamentalDoAcorde + 11
                NomeAcorde = CifraFundamentalDoAcorde & "7M(#11)"
            ElseIf FamiliaDoAcorde = 36 Then 'X(add9)
                kkFinal(0) = FundamentalDoAcorde : kkFinal(1) = FundamentalDoAcorde + 2 : kkFinal(2) = FundamentalDoAcorde + 4 : kkFinal(3) = FundamentalDoAcorde + 7
                NomeAcorde = CifraFundamentalDoAcorde & "(add9)"
            ElseIf FamiliaDoAcorde = 37 Then 'Xm(add9)
                kkFinal(0) = FundamentalDoAcorde : kkFinal(1) = FundamentalDoAcorde + 2 : kkFinal(2) = FundamentalDoAcorde + 3 : kkFinal(3) = FundamentalDoAcorde + 7
                NomeAcorde = CifraFundamentalDoAcorde & "m(add9)"
            ElseIf FamiliaDoAcorde = 38 Then 'X9(b5)
                kkFinal(0) = FundamentalDoAcorde + 2 : kkFinal(1) = FundamentalDoAcorde + 4 : kkFinal(2) = FundamentalDoAcorde + 6 : kkFinal(3) = FundamentalDoAcorde + 10
                NomeAcorde = CifraFundamentalDoAcorde & "9(b5)"
            ElseIf FamiliaDoAcorde = 39 Then 'X9(#5)
                kkFinal(0) = FundamentalDoAcorde + 2 : kkFinal(1) = FundamentalDoAcorde + 4 : kkFinal(2) = FundamentalDoAcorde + 8 : kkFinal(3) = FundamentalDoAcorde + 10
                NomeAcorde = CifraFundamentalDoAcorde & "9(#5)"
            ElseIf FamiliaDoAcorde = 40 Then 'X11
                kkFinal(0) = FundamentalDoAcorde + 2 : kkFinal(1) = FundamentalDoAcorde + 5 : kkFinal(2) = FundamentalDoAcorde + 7 : kkFinal(3) = FundamentalDoAcorde + 10
                NomeAcorde = CifraFundamentalDoAcorde & "11"
            ElseIf FamiliaDoAcorde = 41 Then 'X(#11)
                kkFinal(0) = FundamentalDoAcorde + 2 : kkFinal(1) = FundamentalDoAcorde + 6 : kkFinal(2) = FundamentalDoAcorde + 7 : kkFinal(3) = FundamentalDoAcorde + 10
                NomeAcorde = CifraFundamentalDoAcorde & "(#11)"
            End If


            'Assegurar que nenhuma nota será superior a 12
            If kkFinal(1) > 12 Then kkFinal(1) -= 12
            If kkFinal(2) > 12 Then kkFinal(2) -= 12
            If kkFinal(3) > 12 Then kkFinal(3) -= 12

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub Gera_Nota_Aleatoria()

        Try

            P2.Refresh() : P5.Refresh() : P7.Refresh() : P10.Refresh() : P12.Refresh() : P14.Refresh() : P17.Refresh() : P19.Refresh() : P22.Refresh() : P24.Refresh() : P26.Refresh() : P29.Refresh() : P31.Refresh() : P34.Refresh() : P36.Refresh() : P38.Refresh() : P41.Refresh() : P43.Refresh() : P46.Refresh() : P48.Refresh() : P50.Refresh() : P53.Refresh() : P55.Refresh() : P58.Refresh() : P60.Refresh() : P62.Refresh() : P65.Refresh() : P67.Refresh() : P70.Refresh() : P72.Refresh() : P74.Refresh() : P77.Refresh() : P79.Refresh() : P82.Refresh() : P84.Refresh() : P86.Refresh()
            P1.Refresh() : P3.Refresh() : P4.Refresh() : P6.Refresh() : P8.Refresh() : P9.Refresh() : P11.Refresh() : P13.Refresh() : P15.Refresh() : P16.Refresh() : P18.Refresh() : P20.Refresh() : P21.Refresh() : P23.Refresh() : P25.Refresh() : P27.Refresh() : P28.Refresh() : P30.Refresh() : P32.Refresh() : P33.Refresh() : P35.Refresh() : P37.Refresh() : P39.Refresh() : P40.Refresh() : P42.Refresh() : P44.Refresh() : P45.Refresh() : P47.Refresh() : P49.Refresh() : P51.Refresh() : P52.Refresh() : P54.Refresh() : P56.Refresh() : P57.Refresh() : P59.Refresh() : P61.Refresh() : P63.Refresh() : P64.Refresh() : P66.Refresh() : P68.Refresh() : P69.Refresh() : P71.Refresh() : P73.Refresh() : P75.Refresh() : P76.Refresh() : P78.Refresh() : P80.Refresh() : P81.Refresh() : P83.Refresh() : P85.Refresh() : P87.Refresh() : P88.Refresh()

            Dim Rect As New Rectangle(0, 0, 200, 200) 'limpa imagens das notas que as vezes ficavam nesta região da tela
            Me.Invalidate(Rect)

            CorTecla2.Left = -100 : CorTecla.Left = -100

            If CheckBox26.Checked Then CorretoIncorreto = Nothing
            'atualiza CorretoIncorreto na tela
            Dim Rect7 As New Rectangle(342, 184, 349, 61)
            Me.Invalidate(Rect7)

            ww = 1
            ee = 0
            NomeTeclaFinal = "" : NomeCifra(1) = ""

            m = ""

            'ProgressBar1.Value = FormatNumber(((zz / vv) * 100), 2)
            '((Minutos.Text * 60) + Segundos.Text + (Decimos.Text / 100)) * 1000

            GeraNotaTransposta = 100
            If My.Settings.NovoValorTransposição(4) = "True" Then
                My.Settings.NovoValorTransposição(0) = "False"
                My.Settings.NovoValorTransposição(1) = "False"
                My.Settings.NovoValorTransposição(2) = "False"
                My.Settings.NovoValorTransposição(3) = "False"


                Dim Rect1 As New Rectangle(0, 70, 325, 500)
                Me.Invalidate(Rect1)
                Do While GeraNotaTransposta > 5

                    ' Create a byte array to hold the random value.
                    Dim randomNumber(0) As Byte

                    ' Create a new instance of the RNGCryptoServiceProvider.
                    Dim Gen As New Security.Cryptography.RNGCryptoServiceProvider()

                    ' Fill the array with a random value.
                    Gen.GetBytes(randomNumber)

                    ' Convert the byte to an integer value to make the modulus operation easier.
                    GeraNotaTransposta = Int(Convert.ToInt32(randomNumber(0)))

                Loop

                If GeraNotaTransposta = 3 Then
                    GeraNotaTransposta = 100
                    Do While GeraNotaTransposta > 4

                        ' Create a byte array to hold the random value.
                        Dim randomNumber(0) As Byte

                        ' Create a new instance of the RNGCryptoServiceProvider.
                        Dim Gen As New Security.Cryptography.RNGCryptoServiceProvider()

                        ' Fill the array with a random value.
                        Gen.GetBytes(randomNumber)

                        ' Convert the byte to an integer value to make the modulus operation easier.
                        GeraNotaTransposta = Int(Convert.ToInt32(randomNumber(0)))

                    Loop

                    If GeraNotaTransposta = 1 Then
                        My.Settings.NovoValorTransposição(0) = "True"
                    ElseIf GeraNotaTransposta = 2 Then
                        My.Settings.NovoValorTransposição(1) = "True"
                    ElseIf GeraNotaTransposta = 3 Then
                        My.Settings.NovoValorTransposição(2) = "True"
                    ElseIf GeraNotaTransposta = 4 Then
                        My.Settings.NovoValorTransposição(3) = "True"
                    End If

                    Dim Rect2 As New Rectangle(0, 70, 325, 500)
                    Me.Invalidate(Rect2)

                End If
            End If


            QtdeLoopingsPermitidos = 0

            If (CheckBox3.Checked AndAlso Not CheckBox4.Checked) OrElse DóPrimeiraLinhaToolStripMenuItem1.Checked OrElse DóSegundaLinhaToolStripMenuItem1.Checked OrElse DóTerceiraLinhaToolStripMenuItem1.Checked OrElse DóQuartaLinhaToolStripMenuItem1.Checked OrElse DóQuintaLinhaToolStripMenuItem1.Checked OrElse FáQuintaLinhaToolStripMenuItem1.Checked OrElse FáTerceiraLinhaToolStripMenuItem1.Checked OrElse SolPrimeiraLinhaToolStripMenuItem1.Checked Then
                'gera números aleatórios
                k = TrackBar1.Value + 1
                AjusteTrackBar(0) = 0 : AjusteTrackBar(1) = 0
                If DóPrimeiraLinhaToolStripMenuItem1.Checked Then
                    AjusteTrackBar(0) = 17 : AjusteTrackBar(1) = 2
                ElseIf DóSegundaLinhaToolStripMenuItem1.Checked Then
                    AjusteTrackBar(0) = 19 : AjusteTrackBar(1) = 4
                ElseIf DóTerceiraLinhaToolStripMenuItem1.Checked Then
                    AjusteTrackBar(0) = 21 : AjusteTrackBar(1) = 6
                ElseIf DóQuartaLinhaToolStripMenuItem1.Checked Then
                    AjusteTrackBar(0) = 23 : AjusteTrackBar(1) = 8
                ElseIf DóQuintaLinhaToolStripMenuItem1.Checked Then
                    AjusteTrackBar(0) = 25 : AjusteTrackBar(1) = 10
                ElseIf FáTerceiraLinhaToolStripMenuItem1.Checked Then
                    AjusteTrackBar(0) = 25 : AjusteTrackBar(1) = 10
                ElseIf FáQuintaLinhaToolStripMenuItem1.Checked Then
                    AjusteTrackBar(0) = 29 : AjusteTrackBar(1) = 14
                ElseIf SolPrimeiraLinhaToolStripMenuItem1.Checked Then
                    AjusteTrackBar(0) = -2 : AjusteTrackBar(1) = 0
                End If


                Do While k > (TrackBar1.Value - AjusteTrackBar(1)) OrElse k < (TrackBar5.Value - AjusteTrackBar(0)) OrElse k = f

                    ' Create a byte array to hold the random value.
                    Dim randomNumber(0) As Byte

                    ' Create a new instance of the RNGCryptoServiceProvider.
                    Dim Gen As New Security.Cryptography.RNGCryptoServiceProvider()

                    ' Fill the array with a random value.
                    Gen.GetBytes(randomNumber)

                    ' Convert the byte to an integer value to make the modulus operation easier.
                    k = Int(Convert.ToInt32(randomNumber(0)))
                    If k < (32 - AjusteTrackBar(0)) Then k = 68

                    If My.Settings.NovoValorTransposição(0) = "True" Then
                        If k >= 60 Then k = 68
                    End If
                    If My.Settings.NovoValorTransposição(2) = "True" Then
                        If k >= 53 Then k = 68
                    End If

                    Select Case k
                        ' A
                        Case 1, 8, 15, 22, 29, 37, 44, 51, 58, 65
                            If My.Settings.NovoValorExibirNotas(5) = "False" Then k = 68
                            ' B
                        Case 2, 9, 16, 23, 30, 38, 45, 52, 59, 66
                            If My.Settings.NovoValorExibirNotas(6) = "False" Then k = 68
                            ' C
                        Case 3, 10, 17, 24, 32, 39, 46, 53, 60, 67
                            If My.Settings.NovoValorExibirNotas(0) = "False" Then k = 68
                            ' D
                        Case 4, 11, 18, 25, 33, 40, 47, 54, 61
                            If My.Settings.NovoValorExibirNotas(1) = "False" Then k = 68
                            ' E
                        Case 5, 12, 19, 26, 34, 41, 48, 55, 62
                            If My.Settings.NovoValorExibirNotas(2) = "False" Then k = 68
                            ' F
                        Case 6, 13, 20, 27, 35, 42, 49, 56, 63
                            If My.Settings.NovoValorExibirNotas(3) = "False" Then k = 68
                            ' G
                        Case 7, 14, 21, 28, 36, 43, 50, 57, 64
                            If My.Settings.NovoValorExibirNotas(4) = "False" Then k = 68
                    End Select

                    If k = f Then
                        If QtdeLoopingsPermitidos = 20 Then
                            QuantidadeDeLoopsQuePoderãoSerExecutados()
                            Exit Sub
                        Else
                            QtdeLoopingsPermitidos += 1
                        End If
                    End If

                Loop
            End If

            If Not CheckBox3.Checked AndAlso CheckBox4.Checked Then
                'gera números aleatórios
                k = TrackBar2.Value - 1
                Do While k < TrackBar2.Value OrElse k > TrackBar4.Value OrElse k = f

                    ' Create a byte array to hold the random value.
                    Dim randomNumber(0) As Byte

                    ' Create a new instance of the RNGCryptoServiceProvider.
                    Dim Gen As New Security.Cryptography.RNGCryptoServiceProvider()

                    ' Fill the array with a random value.
                    Gen.GetBytes(randomNumber)

                    ' Convert the byte to an integer value to make the modulus operation easier.
                    k = Int(Convert.ToInt32(randomNumber(0)))
                    If k > 31 Then k = 0

                    If My.Settings.NovoValorTransposição(1) = "True" Then
                        If k <= 8 Then k = 0
                    End If
                    If My.Settings.NovoValorTransposição(3) = "True" Then
                        If k <= 15 Then k = 0
                    End If

                    Select Case k
                        ' A
                        Case 1, 8, 15, 22, 29, 37, 44, 51, 58, 65
                            If My.Settings.NovoValorExibirNotas(5) = "False" Then k = 0
                            ' B
                        Case 2, 9, 16, 23, 30, 38, 45, 52, 59, 66
                            If My.Settings.NovoValorExibirNotas(6) = "False" Then k = 0
                            ' C
                        Case 3, 10, 17, 24, 32, 39, 46, 53, 60, 67
                            If My.Settings.NovoValorExibirNotas(0) = "False" Then k = 0
                            ' D
                        Case 4, 11, 18, 25, 33, 40, 47, 54, 61
                            If My.Settings.NovoValorExibirNotas(1) = "False" Then k = 0
                            ' E
                        Case 5, 12, 19, 26, 34, 41, 48, 55, 62
                            If My.Settings.NovoValorExibirNotas(2) = "False" Then k = 0
                            ' F
                        Case 6, 13, 20, 27, 35, 42, 49, 56, 63
                            If My.Settings.NovoValorExibirNotas(3) = "False" Then k = 0
                            ' G
                        Case 7, 14, 21, 28, 36, 43, 50, 57, 64
                            If My.Settings.NovoValorExibirNotas(4) = "False" Then k = 0
                    End Select

                    If k = f Then
                        If QtdeLoopingsPermitidos = 20 Then
                            QuantidadeDeLoopsQuePoderãoSerExecutados()
                            Exit Sub
                        Else
                            QtdeLoopingsPermitidos += 1
                        End If
                    End If

                Loop
            End If


            If CheckBox3.Checked AndAlso CheckBox4.Checked Then
                'gera números aleatórios
                k = TrackBar1.Value + 1
                Do While (k < TrackBar2.Value OrElse k > TrackBar4.Value) AndAlso (k > TrackBar1.Value OrElse k < TrackBar5.Value) OrElse k = f

                    ' Create a byte array to hold the random value.
                    Dim randomNumber(0) As Byte

                    ' Create a new instance of the RNGCryptoServiceProvider.
                    Dim Gen As New Security.Cryptography.RNGCryptoServiceProvider()

                    ' Fill the array with a random value.
                    Gen.GetBytes(randomNumber)

                    ' Convert the byte to an integer value to make the modulus operation easier.
                    k = Int(Convert.ToInt32(randomNumber(0)))


                    If My.Settings.NovoValorTransposição(0) = "True" Then
                        If k >= 60 Then k = 68
                    ElseIf My.Settings.NovoValorTransposição(1) = "True" Then
                        If k <= 8 Then k = 68
                    ElseIf My.Settings.NovoValorTransposição(2) = "True" Then
                        If k >= 53 Then k = 68
                    ElseIf My.Settings.NovoValorTransposição(3) = "True" Then
                        If k <= 15 Then k = 68
                    End If

                    Select Case k
                        ' A
                        Case 1, 8, 15, 22, 29, 37, 44, 51, 58, 65
                            If My.Settings.NovoValorExibirNotas(5) = "False" Then k = 68
                            ' B
                        Case 2, 9, 16, 23, 30, 38, 45, 52, 59, 66
                            If My.Settings.NovoValorExibirNotas(6) = "False" Then k = 68
                            ' C
                        Case 3, 10, 17, 24, 32, 39, 46, 53, 60, 67
                            If My.Settings.NovoValorExibirNotas(0) = "False" Then k = 68
                            ' D
                        Case 4, 11, 18, 25, 33, 40, 47, 54, 61
                            If My.Settings.NovoValorExibirNotas(1) = "False" Then k = 68
                            ' E
                        Case 5, 12, 19, 26, 34, 41, 48, 55, 62
                            If My.Settings.NovoValorExibirNotas(2) = "False" Then k = 68
                            ' F
                        Case 6, 13, 20, 27, 35, 42, 49, 56, 63
                            If My.Settings.NovoValorExibirNotas(3) = "False" Then k = 68
                            ' G
                        Case 7, 14, 21, 28, 36, 43, 50, 57, 64
                            If My.Settings.NovoValorExibirNotas(4) = "False" Then k = 68
                    End Select

                    If k = f Then
                        If QtdeLoopingsPermitidos = 20 Then
                            QuantidadeDeLoopsQuePoderãoSerExecutados()
                            Exit Sub
                        Else
                            QtdeLoopingsPermitidos += 1
                        End If
                    End If

                Loop
            End If

            f = k

            'sustenidos e bemois
            'gera números aleatórios
            ii = 1000
            m = ""
            If (NotaExata.Checked AndAlso RadioButton1.Checked) OrElse (NotaExata.Checked AndAlso RadioButton2.Checked) OrElse (NotaGenérica.Checked AndAlso RadioButton1.Checked) Then 'só gera sustenidos e bemois se pontuação por "Cifras" não estiver marcado
                If Acidentes.Enabled = True Then
                    Do While ii > 19 'array é zero based, então isso abrange 20 números, 0 à 19

                        ' Create a byte array to hold the random value.
                        Dim randomNumber(0) As Byte

                        ' Create a new instance of the RNGCryptoServiceProvider.
                        Dim Gen As New Security.Cryptography.RNGCryptoServiceProvider()


                        ' Fill the array with a random value.
                        Gen.GetBytes(randomNumber)

                        ' Convert the byte to an integer value to make the modulus operation easier.
                        ii = Int(Convert.ToInt32(randomNumber(0)))

                    Loop


                    If Percentual(ii, Acidentes.Value) = "x" Then 'ii dará o valor da linha, Acidentes.Value dará o valor da coluna do array
                        ii = 1000
                        Do While ii <> 1 AndAlso ii <> 2 'define se a nota alterada será # ou b

                            ' Create a byte array to hold the random value.
                            Dim randomNumber(0) As Byte

                            ' Create a new instance of the RNGCryptoServiceProvider.
                            Dim Gen As New Security.Cryptography.RNGCryptoServiceProvider()


                            ' Fill the array with a random value.
                            Gen.GetBytes(randomNumber)

                            ' Convert the byte to an integer value to make the modulus operation easier.
                            ii = Int(Convert.ToInt32(randomNumber(0)))

                        Loop


                        If (ii = 1 AndAlso k < 67) OrElse (ii = 2 AndAlso k = 1) Then
                            m = "#"
                        ElseIf (ii = 2 AndAlso k > 1) OrElse (ii = 1 AndAlso k = 67) Then
                            m = "b"
                        End If

                    End If
                End If
            End If



            AjusteDasNotasNaClave = 0
            If DóPrimeiraLinhaToolStripMenuItem1.Checked Then
                AjusteDasNotasNaClave = 10
            ElseIf DóSegundaLinhaToolStripMenuItem1.Checked Then
                AjusteDasNotasNaClave = 20
            ElseIf DóTerceiraLinhaToolStripMenuItem1.Checked Then
                AjusteDasNotasNaClave = 30
            ElseIf DóQuartaLinhaToolStripMenuItem1.Checked Then
                AjusteDasNotasNaClave = 40
            ElseIf DóQuintaLinhaToolStripMenuItem1.Checked Then
                AjusteDasNotasNaClave = 50
            ElseIf FáTerceiraLinhaToolStripMenuItem1.Checked Then
                AjusteDasNotasNaClave = 50
            ElseIf FáQuintaLinhaToolStripMenuItem1.Checked Then
                AjusteDasNotasNaClave = 70
            ElseIf SolPrimeiraLinhaToolStripMenuItem1.Checked Then
                AjusteDasNotasNaClave = -10
            End If


            'o código a seguir é para compensar as notas cujo k < 32 e que estejam nas demais claves que
            'não são as do piano. Os valores menores que 32
            'são exibidos na clave de Fá, e então para serem exibidos na clave de Sol
            'é necessário compensar a distância que separa estas duas claves
            If DóPrimeiraLinhaToolStripMenuItem1.Checked OrElse DóSegundaLinhaToolStripMenuItem1.Checked OrElse DóTerceiraLinhaToolStripMenuItem1.Checked OrElse DóQuartaLinhaToolStripMenuItem1.Checked OrElse DóQuintaLinhaToolStripMenuItem1.Checked OrElse FáQuintaLinhaToolStripMenuItem1.Checked OrElse FáTerceiraLinhaToolStripMenuItem1.Checked OrElse SolPrimeiraLinhaToolStripMenuItem1.Checked Then
                If k < 32 Then
                    AjusteDasNotasNaClave = AjusteDasNotasNaClave + 110
                End If
            End If


            If k = 1 OrElse k = 2 Then
                d = 0
            ElseIf k >= 3 AndAlso k <= 9 Then
                d = 1
            ElseIf k >= 10 AndAlso k <= 16 Then
                d = 2
            ElseIf (k >= 17 AndAlso k <= 23) OrElse (k >= 32 AndAlso k <= 38) Then
                d = 3
            ElseIf (k >= 24 AndAlso k <= 30) OrElse (k >= 39 AndAlso k <= 45) Then
                d = 4
            ElseIf k = 31 OrElse (k >= 46 AndAlso k <= 52) Then
                d = 5
            ElseIf k >= 53 AndAlso k <= 59 Then
                d = 6
            ElseIf k >= 60 AndAlso k <= 66 Then
                d = 7
            ElseIf k = 67 Then
                d = 8
            End If


            If My.Settings.NovoValorTransposição(0) = "True" Then
                d += 1
            ElseIf My.Settings.NovoValorTransposição(1) = "True" Then
                d -= 1
            ElseIf My.Settings.NovoValorTransposição(2) = "True" Then
                d += 2
            ElseIf My.Settings.NovoValorTransposição(3) = "True" Then
                d -= 2
            End If

            Select Case k
                ' A
                Case 1, 8, 15, 22, 29, 37, 44, 51, 58, 65

                    If NomeDoMenu = "B - G#m (F# - C# - G# - D# - A#)" OrElse _
                    NomeDoMenu = "F# - D#m (F# - C# - G# - D# - A# - E#)" OrElse _
                    NomeDoMenu = "C# - A#m (F# - C# - G# - D# - A# - E# - B#)" Then
                        m = "#"
                    ElseIf NomeDoMenu = "Cb - Abm (Bb - Eb - Ab - Db - Gb - Cb - Fb)" OrElse _
                    NomeDoMenu = "Gb - Ebm (Bb - Eb - Ab - Db - Gb - Cb)" OrElse _
                    NomeDoMenu = "Db - Bbm (Bb - Eb - Ab - Db - Gb)" OrElse _
                    NomeDoMenu = "Ab - Fm (Bb - Eb - Ab - Db)" OrElse _
                    NomeDoMenu = "Eb - Cm (Bb - Eb - Ab)" Then
                        m = "b"
                    End If

                    o = "A" & m & d
                    LetraCifra = "Lá" & m

                    If My.Settings.NovoValorCoresNotas = True Then
                        If k = 1 Then
                            CorTecla.Image = My.Resources.Tecla_BrancaPressionadaA
                        Else
                            CorTecla.Image = My.Resources.Tecla_BrancaPressionadaA2
                        End If
                        CorTecla2.Image = My.Resources.Tecla_PretaPressionadaA
                    Else
                        If k = 1 Then
                            CorTecla.Image = My.Resources.Tecla_BrancaPressionada
                        Else
                            CorTecla.Image = My.Resources.Tecla_BrancaPressionada2
                        End If
                        CorTecla2.Image = My.Resources.Tecla_PretaPressionada
                    End If

                    ' B
                Case 2, 9, 16, 23, 30, 38, 45, 52, 59, 66

                    If NomeDoMenu = "Cb - Abm (Bb - Eb - Ab - Db - Gb - Cb - Fb)" OrElse _
                    NomeDoMenu = "Gb - Ebm (Bb - Eb - Ab - Db - Gb - Cb)" OrElse _
                    NomeDoMenu = "Db - Bbm (Bb - Eb - Ab - Db - Gb)" OrElse _
                    NomeDoMenu = "Ab - Fm (Bb - Eb - Ab - Db)" OrElse _
                    NomeDoMenu = "Eb - Cm (Bb - Eb - Ab)" OrElse _
                    NomeDoMenu = "Bb - Gm (Bb - Eb)" OrElse _
                    NomeDoMenu = "F - Dm (Bb)" Then
                        m = "b"
                    ElseIf NomeDoMenu = "C# - A#m (F# - C# - G# - D# - A# - E# - B#)" Then
                        m = "#"
                    End If

                    o = "B" & m & d
                    LetraCifra = "Si" & m

                    If My.Settings.NovoValorCoresNotas = True Then
                        If m = "#" Then
                            If k = 66 Then
                                CorTecla.Image = My.Resources.Tecla_BrancaPressionadaBsustenido2
                            Else
                                CorTecla.Image = My.Resources.Tecla_BrancaPressionadaBsustenido
                            End If
                        Else
                            CorTecla.Image = My.Resources.Tecla_BrancaPressionadaB
                        End If
                        CorTecla2.Image = My.Resources.Tecla_PretaPressionadaB
                    Else
                        If m = "#" Then
                            CorTecla.Image = My.Resources.Tecla_BrancaPressionada
                        Else
                            CorTecla.Image = My.Resources.Tecla_BrancaPressionada3
                        End If
                        CorTecla2.Image = My.Resources.Tecla_PretaPressionada
                    End If

                    ' C
                Case 3, 10, 17, 24, 32, 39, 46, 53, 60, 67

                    If NomeDoMenu = "D - Bm (F# - C#)" OrElse _
                        NomeDoMenu = "A - F#m (F# - C# - G#)" OrElse _
                        NomeDoMenu = "E - C#m (F# - C# - G# - D#)" OrElse _
                        NomeDoMenu = "B - G#m (F# - C# - G# - D# - A#)" OrElse _
                        NomeDoMenu = "F# - D#m (F# - C# - G# - D# - A# - E#)" OrElse _
                        NomeDoMenu = "C# - A#m (F# - C# - G# - D# - A# - E# - B#)" Then
                        m = "#"
                    ElseIf NomeDoMenu = "Cb - Abm (Bb - Eb - Ab - Db - Gb - Cb - Fb)" OrElse _
                    NomeDoMenu = "Gb - Ebm (Bb - Eb - Ab - Db - Gb - Cb)" Then
                        m = "b"
                    End If

                    o = "C" & m & d
                    LetraCifra = "Dó" & m

                    If My.Settings.NovoValorCoresNotas = True Then
                        If k = 67 Then
                            CorTecla.Image = My.Resources.Tecla_BrancaPressionadaC2
                        Else
                            If m = "b" Then
                                CorTecla.Image = My.Resources.Tecla_BrancaPressionadaCbemol
                            Else
                                CorTecla.Image = My.Resources.Tecla_BrancaPressionadaC
                            End If
                        End If
                        CorTecla2.Image = My.Resources.Tecla_PretaPressionadaC
                    Else
                        If m = "b" Then
                            CorTecla.Image = My.Resources.Tecla_BrancaPressionada3
                        Else
                            CorTecla.Image = My.Resources.Tecla_BrancaPressionada
                        End If
                        CorTecla2.Image = My.Resources.Tecla_PretaPressionada
                    End If

                    ' D
                Case 4, 11, 18, 25, 33, 40, 47, 54, 61

                    If NomeDoMenu = "E - C#m (F# - C# - G# - D#)" OrElse _
                        NomeDoMenu = "B - G#m (F# - C# - G# - D# - A#)" OrElse _
                        NomeDoMenu = "F# - D#m (F# - C# - G# - D# - A# - E#)" OrElse _
                        NomeDoMenu = "C# - A#m (F# - C# - G# - D# - A# - E# - B#)" Then
                        m = "#"
                    ElseIf NomeDoMenu = "Cb - Abm (Bb - Eb - Ab - Db - Gb - Cb - Fb)" OrElse _
                        NomeDoMenu = "Gb - Ebm (Bb - Eb - Ab - Db - Gb - Cb)" OrElse _
                        NomeDoMenu = "Db - Bbm (Bb - Eb - Ab - Db - Gb)" OrElse _
                        NomeDoMenu = "Ab - Fm (Bb - Eb - Ab - Db)" Then
                        m = "b"
                    End If

                    o = "D" & m & d
                    LetraCifra = "Ré" & m

                    If My.Settings.NovoValorCoresNotas = True Then
                        CorTecla.Image = My.Resources.Tecla_BrancaPressionadaD
                        CorTecla2.Image = My.Resources.Tecla_PretaPressionadaD
                    Else
                        CorTecla.Image = My.Resources.Tecla_BrancaPressionada2
                        CorTecla2.Image = My.Resources.Tecla_PretaPressionada
                    End If

                    ' E
                Case 5, 12, 19, 26, 34, 41, 48, 55, 62

                    If NomeDoMenu = "F# - D#m (F# - C# - G# - D# - A# - E#)" OrElse _
                    NomeDoMenu = "C# - A#m (F# - C# - G# - D# - A# - E# - B#)" Then
                        m = "#"
                    ElseIf NomeDoMenu = "Cb - Abm (Bb - Eb - Ab - Db - Gb - Cb - Fb)" OrElse _
                        NomeDoMenu = "Gb - Ebm (Bb - Eb - Ab - Db - Gb - Cb)" OrElse _
                        NomeDoMenu = "Db - Bbm (Bb - Eb - Ab - Db - Gb)" OrElse _
                        NomeDoMenu = "Ab - Fm (Bb - Eb - Ab - Db)" OrElse _
                        NomeDoMenu = "Eb - Cm (Bb - Eb - Ab)" OrElse _
                        NomeDoMenu = "Bb - Gm (Bb - Eb)" Then
                        m = "b"
                    End If

                    o = "E" & m & d
                    LetraCifra = "Mi" & m

                    If My.Settings.NovoValorCoresNotas = True Then
                        If m = "#" Then
                            CorTecla.Image = My.Resources.Tecla_BrancaPressionadaEsustenido
                        Else
                            CorTecla.Image = My.Resources.Tecla_BrancaPressionadaE
                        End If
                        CorTecla2.Image = My.Resources.Tecla_PretaPressionadaE
                    Else
                        If m = "#" Then
                            CorTecla.Image = My.Resources.Tecla_BrancaPressionada
                        Else
                            CorTecla.Image = My.Resources.Tecla_BrancaPressionada3
                        End If
                        CorTecla2.Image = My.Resources.Tecla_PretaPressionada
                    End If

                    ' F
                Case 6, 13, 20, 27, 35, 42, 49, 56, 63

                    If NomeDoMenu = "G - Em (F#)" OrElse _
                        NomeDoMenu = "D - Bm (F# - C#)" OrElse _
                        NomeDoMenu = "A - F#m (F# - C# - G#)" OrElse _
                        NomeDoMenu = "E - C#m (F# - C# - G# - D#)" OrElse _
                        NomeDoMenu = "B - G#m (F# - C# - G# - D# - A#)" OrElse _
                        NomeDoMenu = "F# - D#m (F# - C# - G# - D# - A# - E#)" OrElse _
                        NomeDoMenu = "C# - A#m (F# - C# - G# - D# - A# - E# - B#)" Then
                        m = "#"
                    ElseIf NomeDoMenu = "Cb - Abm (Bb - Eb - Ab - Db - Gb - Cb - Fb)" Then
                        m = "b"
                    End If

                    o = "F" & m & d
                    LetraCifra = "Fá" & m

                    If My.Settings.NovoValorCoresNotas = True Then
                        If m = "b" Then
                            CorTecla.Image = My.Resources.Tecla_BrancaPressionadaFbemol
                        Else
                            CorTecla.Image = My.Resources.Tecla_BrancaPressionadaF
                        End If
                        CorTecla2.Image = My.Resources.Tecla_PretaPressionadaF
                    Else
                        If m = "b" Then
                            CorTecla.Image = My.Resources.Tecla_BrancaPressionada3
                        Else
                            CorTecla.Image = My.Resources.Tecla_BrancaPressionada
                        End If
                        CorTecla2.Image = My.Resources.Tecla_PretaPressionada
                    End If

                    ' G
                Case 7, 14, 21, 28, 36, 43, 50, 57, 64

                    If NomeDoMenu = "A - F#m (F# - C# - G#)" OrElse _
                        NomeDoMenu = "E - C#m (F# - C# - G# - D#)" OrElse _
                        NomeDoMenu = "B - G#m (F# - C# - G# - D# - A#)" OrElse _
                        NomeDoMenu = "F# - D#m (F# - C# - G# - D# - A# - E#)" OrElse _
                        NomeDoMenu = "C# - A#m (F# - C# - G# - D# - A# - E# - B#)" Then
                        m = "#"
                    ElseIf NomeDoMenu = "Cb - Abm (Bb - Eb - Ab - Db - Gb - Cb - Fb)" OrElse _
                    NomeDoMenu = "Gb - Ebm (Bb - Eb - Ab - Db - Gb - Cb)" OrElse _
                    NomeDoMenu = "Db - Bbm (Bb - Eb - Ab - Db - Gb)" Then
                        m = "b"
                    End If

                    o = "G" & m & d
                    LetraCifra = "Sol" & m

                    If My.Settings.NovoValorCoresNotas = True Then
                        CorTecla.Image = My.Resources.Tecla_BrancaPressionadaG
                        CorTecla2.Image = My.Resources.Tecla_PretaPressionadaG
                    Else
                        CorTecla.Image = My.Resources.Tecla_BrancaPressionada2
                        CorTecla2.Image = My.Resources.Tecla_PretaPressionada
                    End If
            End Select


            If o = "A0" Then
                kk = 1
            ElseIf o = "A#0" OrElse o = "Bb0" Then
                kk = 2
            ElseIf o = "B0" OrElse o = "Cb1" Then
                kk = 3
            ElseIf o = "C1" OrElse o = "B#0" Then
                kk = 4
            ElseIf o = "C#1" OrElse o = "Db1" Then
                kk = 5
            ElseIf o = "D1" Then
                kk = 6
            ElseIf o = "D#1" OrElse o = "Eb1" Then
                kk = 7
            ElseIf o = "E1" OrElse o = "Fb1" Then
                kk = 8
            ElseIf o = "F1" OrElse o = "E#1" Then
                kk = 9
            ElseIf o = "F#1" OrElse o = "Gb1" Then
                kk = 10
            ElseIf o = "G1" Then
                kk = 11
            ElseIf o = "G#1" OrElse o = "Ab1" Then
                kk = 12
            ElseIf o = "A1" Then
                kk = 13
            ElseIf o = "A#1" OrElse o = "Bb1" Then
                kk = 14
            ElseIf o = "B1" OrElse o = "Cb2" Then
                kk = 15
            ElseIf o = "C2" OrElse o = "B#1" Then
                kk = 16
            ElseIf o = "C#2" OrElse o = "Db2" Then
                kk = 17
            ElseIf o = "D2" Then
                kk = 18
            ElseIf o = "D#2" OrElse o = "Eb2" Then
                kk = 19
            ElseIf o = "E2" OrElse o = "Fb2" Then
                kk = 20
            ElseIf o = "F2" OrElse o = "E#2" Then
                kk = 21
            ElseIf o = "F#2" OrElse o = "Gb2" Then
                kk = 22
            ElseIf o = "G2" Then
                kk = 23
            ElseIf o = "G#2" OrElse o = "Ab2" Then
                kk = 24
            ElseIf o = "A2" Then
                kk = 25
            ElseIf o = "A#2" OrElse o = "Bb2" Then
                kk = 26
            ElseIf o = "B2" OrElse o = "Cb3" Then
                kk = 27
            ElseIf o = "C3" OrElse o = "B#2" Then
                kk = 28
            ElseIf o = "C#3" OrElse o = "Db3" Then
                kk = 29
            ElseIf o = "D3" Then
                kk = 30
            ElseIf o = "D#3" OrElse o = "Eb3" Then
                kk = 31
            ElseIf o = "E3" OrElse o = "Fb3" Then
                kk = 32
            ElseIf o = "F3" OrElse o = "E#3" Then
                kk = 33
            ElseIf o = "F#3" OrElse o = "Gb3" Then
                kk = 34
            ElseIf o = "G3" Then
                kk = 35
            ElseIf o = "G#3" OrElse o = "Ab3" Then
                kk = 36
            ElseIf o = "A3" Then
                kk = 37
            ElseIf o = "A#3" OrElse o = "Bb3" Then
                kk = 38
            ElseIf o = "B3" OrElse o = "Cb4" Then
                kk = 39
            ElseIf o = "C4" OrElse o = "B#3" Then
                kk = 40
            ElseIf o = "C#4" OrElse o = "Db4" Then
                kk = 41
            ElseIf o = "D4" Then
                kk = 42
            ElseIf o = "D#4" OrElse o = "Eb4" Then
                kk = 43
            ElseIf o = "E4" OrElse o = "Fb4" Then
                kk = 44
            ElseIf o = "F4" OrElse o = "E#4" Then
                kk = 45
            ElseIf o = "F#4" OrElse o = "Gb4" Then
                kk = 46
            ElseIf o = "G4" Then
                kk = 47
            ElseIf o = "G#4" OrElse o = "Ab4" Then
                kk = 48
            ElseIf o = "A4" Then
                kk = 49
            ElseIf o = "A#4" OrElse o = "Bb4" Then
                kk = 50
            ElseIf o = "B4" OrElse o = "Cb5" Then
                kk = 51
            ElseIf o = "C5" OrElse o = "B#4" Then
                kk = 52
            ElseIf o = "C#5" OrElse o = "Db5" Then
                kk = 53
            ElseIf o = "D5" Then
                kk = 54
            ElseIf o = "D#5" OrElse o = "Eb5" Then
                kk = 55
            ElseIf o = "E5" OrElse o = "Fb5" Then
                kk = 56
            ElseIf o = "F5" OrElse o = "E#5" Then
                kk = 57
            ElseIf o = "F#5" OrElse o = "Gb5" Then
                kk = 58
            ElseIf o = "G5" Then
                kk = 59
            ElseIf o = "G#5" OrElse o = "Ab5" Then
                kk = 60
            ElseIf o = "A5" Then
                kk = 61
            ElseIf o = "A#5" OrElse o = "Bb5" Then
                kk = 62
            ElseIf o = "B5" OrElse o = "Cb6" Then
                kk = 63
            ElseIf o = "C6" OrElse o = "B#5" Then
                kk = 64
            ElseIf o = "C#6" OrElse o = "Db6" Then
                kk = 65
            ElseIf o = "D6" Then
                kk = 66
            ElseIf o = "D#6" OrElse o = "Eb6" Then
                kk = 67
            ElseIf o = "E6" OrElse o = "Fb6" Then
                kk = 68
            ElseIf o = "F6" OrElse o = "E#6" Then
                kk = 69
            ElseIf o = "F#6" OrElse o = "Gb6" Then
                kk = 70
            ElseIf o = "G6" Then
                kk = 71
            ElseIf o = "G#6" OrElse o = "Ab6" Then
                kk = 72
            ElseIf o = "A6" Then
                kk = 73
            ElseIf o = "A#6" OrElse o = "Bb6" Then
                kk = 74
            ElseIf o = "B6" OrElse o = "Cb7" Then
                kk = 75
            ElseIf o = "C7" OrElse o = "B#6" Then
                kk = 76
            ElseIf o = "C#7" OrElse o = "Db7" Then
                kk = 77
            ElseIf o = "D7" Then
                kk = 78
            ElseIf o = "D#7" OrElse o = "Eb7" Then
                kk = 79
            ElseIf o = "E7" OrElse o = "Fb7" Then
                kk = 80
            ElseIf o = "F7" OrElse o = "E#7" Then
                kk = 81
            ElseIf o = "F#7" OrElse o = "Gb7" Then
                kk = 82
            ElseIf o = "G7" Then
                kk = 83
            ElseIf o = "G#7" OrElse o = "Ab7" Then
                kk = 84
            ElseIf o = "A7" Then
                kk = 85
            ElseIf o = "A#7" OrElse o = "Bb7" Then
                kk = 86
            ElseIf o = "B7" OrElse o = "Cb8" Then
                kk = 87
            ElseIf o = "C8" OrElse o = "B#7" Then
                kk = 88
            End If

            If RadioButton2.Checked Then
                CorTecla.Left = -100
                CorTecla2.Left = -100
                If Sonoro.Checked Then ColorirEDefinirQualSomSeraTocado()
            Else
                CorTecla.Left = -100
                CorTecla2.Left = -100
                CorTecla.Visible = True
                CorTecla2.Visible = True

                ColorirEDefinirQualSomSeraTocado()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Public Sub ExibirLigadoDesligadoNomeDaNota()
        Try

            If My.Settings.NovoValorExibirNomeDaNota = False Then
                Ligado1.Visible = False
                Desligado1.Visible = True
            Else
                Ligado1.Visible = True
                Desligado1.Visible = False
            End If

            AtualizaRegiões()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.Click, Me.Load, ToolStripMenuItem13.Click, Violão.Click
        Try

            If CheckBox3.Checked OrElse ToolStripMenuItem13.Checked Then
                DóPrimeiraLinhaToolStripMenuItem1.Checked = False
                DóSegundaLinhaToolStripMenuItem1.Checked = False
                DóTerceiraLinhaToolStripMenuItem1.Checked = False
                DóQuartaLinhaToolStripMenuItem1.Checked = False
                FáQuintaLinhaToolStripMenuItem1.Checked = False
                DóQuintaLinhaToolStripMenuItem1.Checked = False
                FáTerceiraLinhaToolStripMenuItem1.Checked = False
                SolPrimeiraLinhaToolStripMenuItem1.Checked = False
                ToolStripMenuItem14.Checked = False
                ToolStripMenuItem16.Checked = False
                ToolStripMenuItem17.Checked = False
                ToolStripMenuItem18.Checked = False
                ToolStripMenuItem20.Checked = False
                ToolStripMenuItem19.Checked = False
                ToolStripMenuItem21.Checked = False
                ToolStripMenuItem22.Checked = False
                DóPrimeiraLinhaToolStripMenuItem3.Checked = False
                ToolStripMenuItem13.Checked = True
            End If
            If CheckBox4.Checked Then
                ToolStripMenuItem14.Checked = True
            Else
                ToolStripMenuItem14.Checked = False
            End If

            ExibirTrackBars()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub CheckBox4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox4.Click, Me.Load, ToolStripMenuItem14.Click
        Try

            If CheckBox4.Checked OrElse ToolStripMenuItem14.Checked Then
                DóPrimeiraLinhaToolStripMenuItem1.Checked = False
                DóSegundaLinhaToolStripMenuItem1.Checked = False
                DóTerceiraLinhaToolStripMenuItem1.Checked = False
                DóQuartaLinhaToolStripMenuItem1.Checked = False
                FáQuintaLinhaToolStripMenuItem1.Checked = False
                DóQuintaLinhaToolStripMenuItem1.Checked = False
                FáTerceiraLinhaToolStripMenuItem1.Checked = False
                SolPrimeiraLinhaToolStripMenuItem1.Checked = False
                DóPrimeiraLinhaToolStripMenuItem3.Checked = False
                ToolStripMenuItem16.Checked = False
                ToolStripMenuItem17.Checked = False
                ToolStripMenuItem18.Checked = False
                ToolStripMenuItem20.Checked = False
                ToolStripMenuItem19.Checked = False
                ToolStripMenuItem21.Checked = False
                ToolStripMenuItem22.Checked = False
                ToolStripMenuItem14.Checked = True
            End If
            If CheckBox3.Checked Then
                ToolStripMenuItem13.Checked = True
            Else
                ToolStripMenuItem13.Checked = False
            End If

            ExibirTrackBars()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged

        Try

            If RadioButton1.Checked Then
                Cronometro.Stop()
                HoraFim = FormatDateTime(Now, DateFormat.LongTime)
                If Pontuação.Visible Then GravaPontuação()
                Timer1.Enabled = True
                CheckBox26.Enabled = False : GerarNotasAutomaticamenteNoIntervaloDeTempoToolStripMenuItem.Enabled = False
                PictureBox15.Visible = False
                PictureBox16.Visible = True
                PictureBox17.Visible = False
                PictureBox18.Visible = True
                PictureBox14.Visible = False

                CorretoIncorreto = Nothing
                'atualiza CorretoIncorreto na tela
                Dim Rect7 As New Rectangle(342, 184, 349, 61)
                Me.Invalidate(Rect7)

                Pontuação.Text = ""
                Erros.Text = ""
                Acertos.Text = ""
                Pontuação.Visible = False
                Erros.Visible = False
                Acertos.Visible = False
                vv = 0
                zz = 0
                ww = 1
                ProgressBar1.Visible = False
                ProgressBar1.Value = 0
                ToolStripMenuItem7.Enabled = True

            End If
            ToolStripMenuItem44.Checked = RadioButton1.Checked

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged

        Try

            AtualizaRegiões()

            If RadioButton2.Checked Then
                Cronometro.Reset()
                Cronometro.Start()
                HoraInicio = FormatDateTime(Now, DateFormat.LongTime)
                Timer1.Enabled = False
                CheckBox26.Enabled = True : GerarNotasAutomaticamenteNoIntervaloDeTempoToolStripMenuItem.Enabled = True
                PictureBox15.Visible = True
                PictureBox16.Visible = False
                PictureBox17.Visible = True
                PictureBox18.Visible = False
                PictureBox14.Visible = True
                Pontuação.Visible = True
                Erros.Visible = True
                Acertos.Visible = True
                CorTecla.Left = -100
                CorTecla2.Left = -100
                bb = 0
                ProgressBar1.Visible = True
                CheckBox9.Checked = False
                ToolStripMenuItem7.Enabled = False

                ValorTotalDeNotas = CInt(TextBox1.Text)
                ExecutaLoopNotasAleatórias()
            End If

            ToolStripMenuItem45.Checked = RadioButton2.Checked

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub ExibePontuação()

        Try

            If bb = 1 Then
                Pontuação.Text = FormatNumber(((zz / vv) * 100), 15)
            Else
                If vv = 0 Then
                    vv += 1
                    Pontuação.Text = FormatNumber(((zz / vv) * 100), 2)
                    vv -= 1
                Else
                    Pontuação.Text = FormatNumber(((zz / vv) * 100), 2)
                End If
            End If

            If vv = 0 Then
                vv += 1
                ProgressBar1.Value = CInt(FormatNumber(((zz / vv) * 100), 2))
                vv -= 1
            Else
                ProgressBar1.Value = CInt(FormatNumber(((zz / vv) * 100), 2))
            End If

            Erros.Text = "Erros: " & (vv - zz)
            If Erros.Text = "Erros: -1" Then Erros.Text = "Erros: 0"
            Acertos.Text = "Acertos: " & zz

            If Pontuação.Text = "0,50" Then bb = 1

            If CDbl(Pontuação.Text) < 95 AndAlso CDbl(Pontuação.Text) >= 90 Then
                Pontuação.ForeColor = Color.FromArgb(10, 213, 10)
            ElseIf CDbl(Pontuação.Text) < 90 AndAlso CDbl(Pontuação.Text) >= 80 Then
                Pontuação.ForeColor = Color.FromArgb(136, 255, 136)
            ElseIf CDbl(Pontuação.Text) < 80 AndAlso CDbl(Pontuação.Text) >= 70 Then
                Pontuação.ForeColor = Color.FromArgb(255, 255, 132)
            ElseIf CDbl(Pontuação.Text) < 70 AndAlso CDbl(Pontuação.Text) >= 60 Then
                Pontuação.ForeColor = Color.FromArgb(255, 255, 10)
            ElseIf CDbl(Pontuação.Text) < 60 AndAlso CDbl(Pontuação.Text) >= 40 Then
                Pontuação.ForeColor = Color.FromArgb(255, 80, 80)
            ElseIf CDbl(Pontuação.Text) < 40 AndAlso CDbl(Pontuação.Text) >= 20 Then
                Pontuação.ForeColor = Color.FromArgb(255, 40, 40)
            ElseIf CDbl(Pontuação.Text) < 20 AndAlso CDbl(Pontuação.Text) >= 0 Then
                Pontuação.ForeColor = Color.FromArgb(255, 10, 10)
            Else
                Pontuação.ForeColor = Color.FromArgb(10, 166, 10)
            End If

            Pontuação.Text = Pontuação.Text & "%"

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub PictureBox20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox20.Click

        Try

            Ajuda.BackgroundImage = My.Resources.Ajuda
            Ajuda.BringToFront()
            Ajuda.Left = 0
            Ajuda.Top = 35
            Ajuda.Visible = Not Ajuda.Visible
            If Ajuda.Visible = True Then
                ProgressBar1.Visible = False
            Else
                If RadioButton2.Checked Then
                    ProgressBar1.Visible = True
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub Ajuda_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Ajuda.Click
        Try

            Ajuda.Visible = Not Ajuda.Visible
            If Ajuda.Visible = True Then
                ProgressBar1.Visible = False
            Else
                If RadioButton2.Checked Then
                    ProgressBar1.Visible = True
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Ligado1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Ligado1.Click
        My.Settings.NovoValorExibirNomeDaNota = False
        ExibirLigadoDesligadoNomeDaNota()
    End Sub

    Private Sub Desligado1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Desligado1.Click
        My.Settings.NovoValorExibirNomeDaNota = True
        ExibirLigadoDesligadoNomeDaNota()
    End Sub

    Private Sub PictureBox17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox17.Click
        RadioButton1.Checked = True
    End Sub

    Private Sub PictureBox18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox18.Click
        RadioButton2.Checked = True
    End Sub

    Private Sub PictureBox16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox16.Click
        RadioButton2.Checked = True
    End Sub

    Private Sub PictureBox15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox15.Click
        RadioButton1.Checked = True
    End Sub

    Private Sub Fechar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Fechar.Click, ToolStripMenuItem49.Click
        Try

            Me.Close()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Fechar_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles Fechar.MouseHover
        Fechar.Image = My.Resources.Fechar_Ativo
    End Sub

    Private Sub Fechar_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Fechar.MouseLeave
        Fechar.Image = My.Resources.Fechar_Inativo
    End Sub

    Private Sub Minimizar_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Minimizar.MouseLeave
        Minimizar.Image = My.Resources.Minimizar_Inativo
    End Sub

    Private Sub Minimizar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Minimizar.Click, ToolStripMenuItem48.Click

        Me.WindowState = CType(1, FormWindowState) '1 é para minimizar

    End Sub

    Private Sub Minimizar_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles Minimizar.MouseHover
        Minimizar.Image = My.Resources.Minimizar_Ativo
    End Sub

    Private Sub Pausar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Pausar.Click
        Try

            Continuar.BringToFront()
            ToolStripMenuItem6.Text = "Continuar"
            ToolStripMenuItem6.Image = My.Resources.Continuar
            Timer1.Enabled = False

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Continuar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Continuar.Click
        Try

            Pausar.BringToFront()
            ToolStripMenuItem6.Text = "Pausar"
            ToolStripMenuItem6.Image = My.Resources.Pausar
            Timer1.Interval = 1
            Timer1.Enabled = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub TeclasPiano_MouseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles P1.MouseUp, P3.MouseUp, P4.MouseUp, P6.MouseUp, P8.MouseUp, P9.MouseUp, P11.MouseUp, P13.MouseUp, P15.MouseUp, P16.MouseUp, P18.MouseUp, P20.MouseUp _
, P21.MouseUp, P23.MouseUp, P25.MouseUp, P27.MouseUp, P28.MouseUp, P30.MouseUp, P32.MouseUp, P33.MouseUp, P35.MouseUp, P37.MouseUp, P39.MouseUp, P40.MouseUp, P42.MouseUp, P44.MouseUp, P45.MouseUp, P47.MouseUp, P49.MouseUp, P51.MouseUp, P52.MouseUp, P54.MouseUp, P56.MouseUp, P57.MouseUp, P59.MouseUp, P61.MouseUp, P63.MouseUp, P64.MouseUp, P66.MouseUp, P68.MouseUp, P69.MouseUp, P71.MouseUp, P73.MouseUp, P75.MouseUp, P76.MouseUp _
, P78.MouseUp, P80.MouseUp, P81.MouseUp, P83.MouseUp, P85.MouseUp, P87.MouseUp, P88.MouseUp, P2.MouseUp, P5.MouseUp, P7.MouseUp, P10.MouseUp, P12.MouseUp, P14.MouseUp, P17.MouseUp, P19.MouseUp, P22.MouseUp, P24.MouseUp, P26.MouseUp, P29.MouseUp, P31.MouseUp, P34.MouseUp, P36.MouseUp, P38.MouseUp, P41.MouseUp, P43.MouseUp, P46.MouseUp, P48.MouseUp, P50.MouseUp, P53.MouseUp, P55.MouseUp, P58.MouseUp, P60.MouseUp _
, P62.MouseUp, P65.MouseUp, P67.MouseUp, P70.MouseUp, P72.MouseUp, P74.MouseUp, P77.MouseUp, P79.MouseUp, P82.MouseUp, P84.MouseUp, P86.MouseUp, CorTecla.MouseUp, CorTecla2.MouseUp

        Try

            ' Obtém referência ao picturebox que invocou este método.

            Dim pbox As PictureBox = DirectCast(sender, PictureBox)
            If pbox.Width = 13 Then
                pbox.Image = My.Resources.Tecla_Preta
            Else
                pbox.Image = My.Resources.Tecla_Branca
            End If
            CorTecla.Visible = False
            CorTecla2.Visible = False

            If RadioButton2.Checked Then
                If Not CheckBox26.Checked Then
                    ExecutaLoopNotasAleatórias()
                Else
                    Timer1.Enabled = True
                End If
            End If
            If pbox.Name <> "CorTecla" AndAlso pbox.Name <> "CorTecla2" Then
                Try
                    MidiPlayer.Play(New NoteOff(0, 1, CByte(CInt(pbox.Name.Replace("P", "")) + 20), CByte(NumericUpDown2.Value)))
                Catch ex As Exception

                    'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
                Finally
                    'esta parte do bloco é executada independentemente se algum erro acontecer ou não

                End Try
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub TeclasPianoPressionadas(ByVal sender As Object, ByVal e As System.EventArgs) Handles P1.MouseDown, P3.MouseDown, P4.MouseDown, P6.MouseDown, P8.MouseDown, P9.MouseDown, P11.MouseDown, P13.MouseDown, P15.MouseDown, P16.MouseDown, P18.MouseDown, P20.MouseDown _
, P21.MouseDown, P23.MouseDown, P25.MouseDown, P27.MouseDown, P28.MouseDown, P30.MouseDown, P32.MouseDown, P33.MouseDown, P35.MouseDown, P37.MouseDown, P39.MouseDown, P40.MouseDown, P42.MouseDown, P44.MouseDown, P45.MouseDown, P47.MouseDown, P49.MouseDown, P51.MouseDown, P52.MouseDown, P54.MouseDown, P56.MouseDown, P57.MouseDown, P59.MouseDown, P61.MouseDown, P63.MouseDown, P64.MouseDown, P66.MouseDown, P68.MouseDown, P69.MouseDown, P71.MouseDown, P73.MouseDown, P75.MouseDown, P76.MouseDown _
, P78.MouseDown, P80.MouseDown, P81.MouseDown, P83.MouseDown, P85.MouseDown, P87.MouseDown, P88.MouseDown, P2.MouseDown, P5.MouseDown, P7.MouseDown, P10.MouseDown, P12.MouseDown, P14.MouseDown, P17.MouseDown, P19.MouseDown, P22.MouseDown, P24.MouseDown, P26.MouseDown, P29.MouseDown, P31.MouseDown, P34.MouseDown, P36.MouseDown, P38.MouseDown, P41.MouseDown, P43.MouseDown, P46.MouseDown, P48.MouseDown, P50.MouseDown, P53.MouseDown, P55.MouseDown, P58.MouseDown, P60.MouseDown _
, P62.MouseDown, P65.MouseDown, P67.MouseDown, P70.MouseDown, P72.MouseDown, P74.MouseDown, P77.MouseDown, P79.MouseDown, P82.MouseDown, P84.MouseDown, P86.MouseDown

        Try

            ' Obtém referência ao picturebox que invocou este método.
            Dim pbox As PictureBox = DirectCast(sender, PictureBox)
            If pbox.Name = "P1" OrElse pbox.Name = "P4" OrElse pbox.Name = "P9" OrElse pbox.Name = "P16" OrElse pbox.Name = "P21" OrElse pbox.Name = "P28" OrElse pbox.Name = "P33" OrElse pbox.Name = "P40" OrElse pbox.Name = "P45" OrElse pbox.Name = "P52" OrElse pbox.Name = "P57" OrElse pbox.Name = "P64" OrElse pbox.Name = "P69" OrElse pbox.Name = "P76" OrElse pbox.Name = "P81" Then
                pbox.Image = My.Resources.Tecla_BrancaPressionada
            ElseIf pbox.Name = "P3" OrElse pbox.Name = "P8" OrElse pbox.Name = "P15" OrElse pbox.Name = "P20" OrElse pbox.Name = "P27" OrElse pbox.Name = "P32" OrElse pbox.Name = "P39" OrElse pbox.Name = "P44" OrElse pbox.Name = "P51" OrElse pbox.Name = "P56" OrElse pbox.Name = "P63" OrElse pbox.Name = "P68" OrElse pbox.Name = "P75" OrElse pbox.Name = "P80" OrElse pbox.Name = "P87" Then
                pbox.Image = My.Resources.Tecla_BrancaPressionada3
            ElseIf pbox.Name = "P6" OrElse pbox.Name = "P11" OrElse pbox.Name = "P13" OrElse pbox.Name = "P18" OrElse pbox.Name = "P23" OrElse pbox.Name = "P25" OrElse pbox.Name = "P30" OrElse pbox.Name = "P35" OrElse pbox.Name = "P37" OrElse pbox.Name = "P42" OrElse pbox.Name = "P47" OrElse pbox.Name = "P49" OrElse pbox.Name = "P54" OrElse pbox.Name = "P59" OrElse pbox.Name = "P61" OrElse pbox.Name = "P66" OrElse pbox.Name = "P71" OrElse pbox.Name = "P73" OrElse pbox.Name = "P78" OrElse pbox.Name = "P83" OrElse pbox.Name = "P85" Then
                pbox.Image = My.Resources.Tecla_BrancaPressionada2
            ElseIf pbox.Name = "P88" Then
                pbox.Image = My.Resources.Tecla_BrancaPressionada4
            Else
                pbox.Image = My.Resources.Tecla_PretaPressionada
            End If

            NomeTeclaFinal = pbox.Name

            If CheckBox11.Checked Then
                If ComboBox1.Text = "Piano Bösendorfer" Then
                    mp3 = "Sons\Notas\Piano Bösendorfer\" & pbox.Name & ".mp3"
                    'TocaSom()
                ElseIf ComboBox1.Text = "Strings (Trio)" Then
                    mp3 = "Sons\Notas\Strings (Trio)\" & pbox.Name & ".mp3"
                    'TocaSom()
                Else
                    Try
                        MidiPlayer.Play(New ProgramChange(0, 1, CType(ComboBox1.Text.Substring(0, 3), GeneralMidiInstruments)))
                        MidiPlayer.Play(New NoteOn(0, 1, CByte(CInt(pbox.Name.Replace("P", "")) + 20), CByte(NumericUpDown2.Value)))
                    Catch ex As Exception

                        'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
                    Finally
                        'esta parte do bloco é executada independentemente se algum erro acontecer ou não

                    End Try
                End If
            End If



            If CInt(TextBox1.Text) = 1 Then
                DefineValorCorretoIncorreto()
                MostraNotaCorreta()
                ExibePontuação()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub DefineValorCorretoIncorreto()

        Try

            w = 2 'valor sera "Incorreto" a não ser que se acerte a nota
            If RadioButton2.Checked AndAlso Pauta.Checked AndAlso NotaGenérica.Checked Then
                If NomeCifra(1) = "A" AndAlso (o = "A0" OrElse o = "A1" OrElse o = "A2" OrElse o = "A3" OrElse o = "A4" OrElse o = "A5" OrElse o = "A6" OrElse o = "A7") Then
                    w = 1
                ElseIf NomeCifra(1) = "A#-Bb" AndAlso (o = "A#0" OrElse o = "A#1" OrElse o = "A#2" OrElse o = "A#3" OrElse o = "A#4" OrElse o = "A#5" OrElse o = "A#6" OrElse o = "A#7" OrElse _
                     o = "Bb0" OrElse o = "Bb1" OrElse o = "Bb2" OrElse o = "Bb3" OrElse o = "Bb4" OrElse o = "Bb5" OrElse o = "Bb6" OrElse o = "Bb7") Then
                    w = 1
                ElseIf NomeCifra(1) = "B-Cb" AndAlso (o = "B0" OrElse o = "B1" OrElse o = "B2" OrElse o = "B3" OrElse o = "B4" OrElse o = "B5" OrElse o = "B6" OrElse o = "B7" OrElse _
                    o = "Cb1" OrElse o = "Cb2" OrElse o = "Cb3" OrElse o = "Cb4" OrElse o = "Cb5" OrElse o = "Cb6" OrElse o = "Cb7" OrElse o = "Cb8") Then
                    w = 1
                ElseIf NomeCifra(1) = "C-B#" AndAlso (o = "C1" OrElse o = "C2" OrElse o = "C3" OrElse o = "C4" OrElse o = "C5" OrElse o = "C6" OrElse o = "C7" OrElse o = "C8" OrElse _
                    o = "B#0" OrElse o = "B#1" OrElse o = "B#2" OrElse o = "B#3" OrElse o = "B#4" OrElse o = "B#5" OrElse o = "B#6" OrElse o = "B#7") Then
                    w = 1
                ElseIf NomeCifra(1) = "C#-Db" AndAlso (o = "C#1" OrElse o = "C#2" OrElse o = "C#3" OrElse o = "C#4" OrElse o = "C#5" OrElse o = "C#6" OrElse o = "C#7" OrElse o = "C#8" OrElse _
                    o = "Db1" OrElse o = "Db2" OrElse o = "Db3" OrElse o = "Db4" OrElse o = "Db5" OrElse o = "Db6" OrElse o = "Db7") Then
                    w = 1
                ElseIf NomeCifra(1) = "D" AndAlso (o = "D1" OrElse o = "D2" OrElse o = "D3" OrElse o = "D4" OrElse o = "D5" OrElse o = "D6" OrElse o = "D7") Then
                    w = 1
                ElseIf NomeCifra(1) = "D#-Eb" AndAlso (o = "D#1" OrElse o = "D#2" OrElse o = "D#3" OrElse o = "D#4" OrElse o = "D#5" OrElse o = "D#6" OrElse o = "D#7" OrElse _
                    o = "Eb1" OrElse o = "Eb2" OrElse o = "Eb3" OrElse o = "Eb4" OrElse o = "Eb5" OrElse o = "Eb6" OrElse o = "Eb7") Then
                    w = 1
                ElseIf NomeCifra(1) = "E-Fb" AndAlso (o = "E1" OrElse o = "E2" OrElse o = "E3" OrElse o = "E4" OrElse o = "E5" OrElse o = "E6" OrElse o = "E7" OrElse _
                    o = "Fb1" OrElse o = "Fb2" OrElse o = "Fb3" OrElse o = "Fb4" OrElse o = "Fb5" OrElse o = "Fb6" OrElse o = "Fb7") Then
                    w = 1
                ElseIf NomeCifra(1) = "F-E#" AndAlso (o = "F1" OrElse o = "F2" OrElse o = "F3" OrElse o = "F4" OrElse o = "F5" OrElse o = "F6" OrElse o = "F7" OrElse _
                    o = "E#1" OrElse o = "E#2" OrElse o = "E#3" OrElse o = "E#4" OrElse o = "E#5" OrElse o = "E#6" OrElse o = "E#7") Then
                    w = 1
                ElseIf NomeCifra(1) = "F#-Gb" AndAlso (o = "F#1" OrElse o = "F#2" OrElse o = "F#3" OrElse o = "F#4" OrElse o = "F#5" OrElse o = "F#6" OrElse o = "F#7" OrElse _
                    o = "Gb1" OrElse o = "Gb2" OrElse o = "Gb3" OrElse o = "Gb4" OrElse o = "Gb5" OrElse o = "Gb6" OrElse o = "Gb7") Then
                    w = 1
                ElseIf NomeCifra(1) = "G" AndAlso (o = "G1" OrElse o = "G2" OrElse o = "G3" OrElse o = "G4" OrElse o = "G5" OrElse o = "G6" OrElse o = "G7") Then
                    w = 1
                ElseIf NomeCifra(1) = "G#-Ab" AndAlso (o = "G#1" OrElse o = "G#2" OrElse o = "G#3" OrElse o = "G#4" OrElse o = "G#5" OrElse o = "G#6" OrElse o = "G#7" OrElse _
                    o = "Ab0" OrElse o = "Ab1" OrElse o = "Ab2" OrElse o = "Ab3" OrElse o = "Ab4" OrElse o = "Ab5" OrElse o = "Ab6" OrElse o = "Ab7") Then
                    w = 1
                End If
            ElseIf (RadioButton2.Checked AndAlso Pauta.Checked AndAlso NotaExata.Checked) OrElse (RadioButton2.Checked AndAlso CifrasDeAcordes.Checked) Then
                If CInt(TextBox1.Text) = 1 AndAlso NomeTeclaFinal = "P" & kkFinal(0) Then
                    w = 1
                ElseIf CInt(TextBox1.Text) = 2 AndAlso (NomeTeclaFinal = "P" & kkFinal(0) & "P" & kkFinal(1) OrElse NomeTeclaFinal = "P" & kkFinal(1) & "P" & kkFinal(0)) Then
                    w = 1
                ElseIf CInt(TextBox1.Text) = 3 AndAlso (NomeTeclaFinal = "P" & kkFinal(0) & "P" & kkFinal(1) & "P" & kkFinal(2) OrElse NomeTeclaFinal = "P" & kkFinal(0) & "P" & kkFinal(2) & "P" & kkFinal(1) OrElse NomeTeclaFinal = "P" & kkFinal(1) & "P" & kkFinal(0) & "P" & kkFinal(2) OrElse NomeTeclaFinal = "P" & kkFinal(1) & "P" & kkFinal(2) & "P" & kkFinal(0) OrElse NomeTeclaFinal = "P" & kkFinal(2) & "P" & kkFinal(0) & "P" & kkFinal(1) OrElse NomeTeclaFinal = "P" & kkFinal(2) & "P" & kkFinal(1) & "P" & kkFinal(0)) Then
                    w = 1
                ElseIf CInt(TextBox1.Text) = 4 AndAlso (NomeTeclaFinal = "P" & kkFinal(0) & "P" & kkFinal(1) & "P" & kkFinal(2) & "P" & kkFinal(3) OrElse _
                        NomeTeclaFinal = "P" & kkFinal(0) & "P" & kkFinal(1) & "P" & kkFinal(3) & "P" & kkFinal(2) OrElse _
                        NomeTeclaFinal = "P" & kkFinal(0) & "P" & kkFinal(2) & "P" & kkFinal(1) & "P" & kkFinal(3) OrElse _
                        NomeTeclaFinal = "P" & kkFinal(0) & "P" & kkFinal(2) & "P" & kkFinal(3) & "P" & kkFinal(1) OrElse _
                        NomeTeclaFinal = "P" & kkFinal(0) & "P" & kkFinal(3) & "P" & kkFinal(1) & "P" & kkFinal(2) OrElse _
                        NomeTeclaFinal = "P" & kkFinal(0) & "P" & kkFinal(3) & "P" & kkFinal(2) & "P" & kkFinal(1) OrElse _
                        NomeTeclaFinal = "P" & kkFinal(1) & "P" & kkFinal(0) & "P" & kkFinal(2) & "P" & kkFinal(3) OrElse _
                        NomeTeclaFinal = "P" & kkFinal(1) & "P" & kkFinal(0) & "P" & kkFinal(3) & "P" & kkFinal(2) OrElse _
                        NomeTeclaFinal = "P" & kkFinal(1) & "P" & kkFinal(2) & "P" & kkFinal(0) & "P" & kkFinal(3) OrElse _
                        NomeTeclaFinal = "P" & kkFinal(1) & "P" & kkFinal(2) & "P" & kkFinal(3) & "P" & kkFinal(0) OrElse _
                        NomeTeclaFinal = "P" & kkFinal(1) & "P" & kkFinal(3) & "P" & kkFinal(0) & "P" & kkFinal(2) OrElse _
                        NomeTeclaFinal = "P" & kkFinal(1) & "P" & kkFinal(3) & "P" & kkFinal(2) & "P" & kkFinal(0) OrElse _
                        NomeTeclaFinal = "P" & kkFinal(2) & "P" & kkFinal(0) & "P" & kkFinal(1) & "P" & kkFinal(3) OrElse _
                        NomeTeclaFinal = "P" & kkFinal(2) & "P" & kkFinal(0) & "P" & kkFinal(3) & "P" & kkFinal(1) OrElse _
                        NomeTeclaFinal = "P" & kkFinal(2) & "P" & kkFinal(1) & "P" & kkFinal(0) & "P" & kkFinal(3) OrElse _
                        NomeTeclaFinal = "P" & kkFinal(2) & "P" & kkFinal(1) & "P" & kkFinal(3) & "P" & kkFinal(0) OrElse _
                        NomeTeclaFinal = "P" & kkFinal(2) & "P" & kkFinal(3) & "P" & kkFinal(0) & "P" & kkFinal(1) OrElse _
                        NomeTeclaFinal = "P" & kkFinal(2) & "P" & kkFinal(3) & "P" & kkFinal(1) & "P" & kkFinal(0) OrElse _
                        NomeTeclaFinal = "P" & kkFinal(3) & "P" & kkFinal(0) & "P" & kkFinal(1) & "P" & kkFinal(2) OrElse _
                        NomeTeclaFinal = "P" & kkFinal(3) & "P" & kkFinal(0) & "P" & kkFinal(2) & "P" & kkFinal(1) OrElse _
                        NomeTeclaFinal = "P" & kkFinal(3) & "P" & kkFinal(1) & "P" & kkFinal(0) & "P" & kkFinal(2) OrElse _
                        NomeTeclaFinal = "P" & kkFinal(3) & "P" & kkFinal(1) & "P" & kkFinal(2) & "P" & kkFinal(0) OrElse _
                        NomeTeclaFinal = "P" & kkFinal(3) & "P" & kkFinal(2) & "P" & kkFinal(0) & "P" & kkFinal(1) OrElse _
                        NomeTeclaFinal = "P" & kkFinal(3) & "P" & kkFinal(2) & "P" & kkFinal(1) & "P" & kkFinal(0)) Then
                    w = 1
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub MostraNotaCorreta()

        Try

            If RadioButton2.Checked Then

                If w = 1 Then
                    CorretoIncorreto = My.Resources.Correto
                ElseIf w = 2 Then
                    CorretoIncorreto = My.Resources.Incorreto
                End If

                'atualiza CorretoIncorreto na tela
                Dim Rect7 As New Rectangle(342, 184, 349, 61)
                Me.Invalidate(Rect7)


                CorTecla.Visible = False
                CorTecla2.Visible = False


                If w = 2 Then
                    CorTecla.Image = My.Resources.Tecla_Incorreta
                    CorTecla2.Image = My.Resources.Tecla_PretaIncorreta
                    CorTecla.Visible = True
                    CorTecla2.Visible = True
                    If ww = 1 Then
                        vv += 1
                        'ExibePontuação()
                        ww = 2
                        aa = "Incorreto"
                    End If

                ElseIf w = 1 Then
                    CorTecla.Image = My.Resources.Tecla_Correta
                    CorTecla2.Image = My.Resources.Tecla_PretaCorreta
                    CorTecla.Visible = True
                    CorTecla2.Visible = True
                    If ww = 1 Then
                        zz += 1
                        vv += 1
                        'ExibePontuação()
                        ww = 2
                        aa = "Correto"
                    End If
                End If

                If kk = 0 Then kk = 1
                If kk = 2 OrElse kk = 5 OrElse kk = 7 OrElse kk = 10 OrElse kk = 12 OrElse kk = 14 OrElse kk = 17 OrElse kk = 19 OrElse kk = 22 OrElse kk = 24 OrElse kk = 26 OrElse kk = 29 OrElse kk = 31 OrElse kk = 34 OrElse kk = 36 OrElse kk = 38 OrElse kk = 41 OrElse kk = 43 OrElse kk = 46 OrElse kk = 48 OrElse kk = 50 OrElse kk = 53 OrElse kk = 55 OrElse kk = 58 OrElse kk = 60 OrElse kk = 62 OrElse kk = 65 OrElse kk = 67 OrElse kk = 70 OrElse kk = 72 OrElse kk = 74 OrElse kk = 77 OrElse kk = 79 OrElse kk = 82 OrElse kk = 84 OrElse kk = 86 Then
                    CorTecla2.Location = PosiçãoTecla(kk - 1, 0)
                Else
                    CorTecla.Location = PosiçãoTecla(kk - 1, 0)
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub TeclasPianoPressionadas2(ByVal sender As Object, ByVal e As System.EventArgs) Handles CorTecla.MouseDown, CorTecla2.MouseDown
        ColorirEDefinirQualSomSeraTocado()
    End Sub

    Private Sub TeclasPiano_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles P1.MouseHover, P3.MouseHover, P4.MouseHover, P6.MouseHover, P8.MouseHover, P9.MouseHover, P11.MouseHover, P13.MouseHover, P15.MouseHover, P16.MouseHover, P18.MouseHover, P20.MouseHover _
, P21.MouseHover, P23.MouseHover, P25.MouseHover, P27.MouseHover, P28.MouseHover, P30.MouseHover, P32.MouseHover, P33.MouseHover, P35.MouseHover, P37.MouseHover, P39.MouseHover, P40.MouseHover, P42.MouseHover, P44.MouseHover, P45.MouseHover, P47.MouseHover, P49.MouseHover, P51.MouseHover, P52.MouseHover, P54.MouseHover, P56.MouseHover, P57.MouseHover, P59.MouseHover, P61.MouseHover, P63.MouseHover, P64.MouseHover, P66.MouseHover, P68.MouseHover, P69.MouseHover, P71.MouseHover, P73.MouseHover, P75.MouseHover, P76.MouseHover _
, P78.MouseHover, P80.MouseHover, P81.MouseHover, P83.MouseHover, P85.MouseHover, P87.MouseHover, P88.MouseHover, P2.MouseHover, P5.MouseHover, P7.MouseHover, P10.MouseHover, P12.MouseHover, P14.MouseHover, P17.MouseHover, P19.MouseHover, P22.MouseHover, P24.MouseHover, P26.MouseHover, P29.MouseHover, P31.MouseHover, P34.MouseHover, P36.MouseHover, P38.MouseHover, P41.MouseHover, P43.MouseHover, P46.MouseHover, P48.MouseHover, P50.MouseHover, P53.MouseHover, P55.MouseHover, P58.MouseHover, P60.MouseHover _
, P62.MouseHover, P65.MouseHover, P67.MouseHover, P70.MouseHover, P72.MouseHover, P74.MouseHover, P77.MouseHover, P79.MouseHover, P82.MouseHover, P84.MouseHover, P86.MouseHover, CorTecla.MouseHover, CorTecla2.MouseHover

        Try

            If RadioButton1.Checked AndAlso My.Settings.NovoValorExibirNomeDaNota = True Then
                ' Obtém referência ao picturebox que invocou este método.
                pbox2 = DirectCast(sender, PictureBox)

                ToolTipNotas()
                ToolTipNota(0) = ToolTipNota2

                'atualiza ToolTipNota(0)
                Dim Rect8 As New Rectangle(479, 558, 104, 58)
                Me.Invalidate(Rect8)
            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub TeclasPiano_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles P1.MouseLeave, P3.MouseLeave, P4.MouseLeave, P6.MouseLeave, P8.MouseLeave, P9.MouseLeave, P11.MouseLeave, P13.MouseLeave, P15.MouseLeave, P16.MouseLeave, P18.MouseLeave, P20.MouseLeave _
, P21.MouseLeave, P23.MouseLeave, P25.MouseLeave, P27.MouseLeave, P28.MouseLeave, P30.MouseLeave, P32.MouseLeave, P33.MouseLeave, P35.MouseLeave, P37.MouseLeave, P39.MouseLeave, P40.MouseLeave, P42.MouseLeave, P44.MouseLeave, P45.MouseLeave, P47.MouseLeave, P49.MouseLeave, P51.MouseLeave, P52.MouseLeave, P54.MouseLeave, P56.MouseLeave, P57.MouseLeave, P59.MouseLeave, P61.MouseLeave, P63.MouseLeave, P64.MouseLeave, P66.MouseLeave, P68.MouseLeave, P69.MouseLeave, P71.MouseLeave, P73.MouseLeave, P75.MouseLeave, P76.MouseLeave _
, P78.MouseLeave, P80.MouseLeave, P81.MouseLeave, P83.MouseLeave, P85.MouseLeave, P87.MouseLeave, P88.MouseLeave, P2.MouseLeave, P5.MouseLeave, P7.MouseLeave, P10.MouseLeave, P12.MouseLeave, P14.MouseLeave, P17.MouseLeave, P19.MouseLeave, P22.MouseLeave, P24.MouseLeave, P26.MouseLeave, P29.MouseLeave, P31.MouseLeave, P34.MouseLeave, P36.MouseLeave, P38.MouseLeave, P41.MouseLeave, P43.MouseLeave, P46.MouseLeave, P48.MouseLeave, P50.MouseLeave, P53.MouseLeave, P55.MouseLeave, P58.MouseLeave, P60.MouseLeave _
, P62.MouseLeave, P65.MouseLeave, P67.MouseLeave, P70.MouseLeave, P72.MouseLeave, P74.MouseLeave, P77.MouseLeave, P79.MouseLeave, P82.MouseLeave, P84.MouseLeave, P86.MouseLeave, CorTecla.MouseLeave, CorTecla2.MouseLeave

        Try

            ToolTipNota(0) = Nothing
            'atualiza ToolTipNota(0)
            Dim Rect8 As New Rectangle(479, 558, 104, 58)
            Me.Invalidate(Rect8)

            yyy = 0

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub MemoNotes_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

        Try

            ' A tecla pressionada está na nossa lista? 
            If Not lista.Contains(e.KeyCode) Then
                ' Não! A adicionamos à lista.
                lista.Add(e.KeyCode)

                'If e.KeyCode = Keys.A Then mp3 = "Sons\Notas\P28.mp3" : TocaSom()
                'If e.KeyCode = Keys.S Then mp3 = "Sons\Notas\P30.mp3" : TocaSom()
                'If e.KeyCode = Keys.D Then mp3 = "Sons\Notas\P32.mp3" : TocaSom()
                'If e.KeyCode = Keys.F Then mp3 = "Sons\Notas\P33.mp3" : TocaSom()
                'If e.KeyCode = Keys.G Then mp3 = "Sons\Notas\P35.mp3" : TocaSom()

                If e.KeyCode = Keys.V Then CaptaçãoSom = True 'V de Voz
                If e.KeyCode = Keys.R Then 'R de Repetir
                    'If Sonoro.Checked AndAlso RadioButton2.Checked Then SoarNota()
                    If RadioButton2.Checked Then SoarNota()
                End If
                If e.KeyCode = Keys.M Then ReligarCaptaçãoSom() : IndicadorDeFrequencia.BringToFront() 'M de Microfone

                If CInt(TextBox1.Text) = 1 Then
                    If RadioButton2.Checked AndAlso NotaGenérica.Checked Then
                        If e.KeyCode = Keys.A OrElse e.KeyCode = Keys.F6 OrElse e.KeyCode = Keys.Divide Then
                            NomeCifra(1) = "A"
                        ElseIf e.KeyCode = Keys.B OrElse e.KeyCode = Keys.F7 OrElse e.KeyCode = Keys.Subtract Then
                            NomeCifra(1) = "B-Cb"
                        ElseIf e.KeyCode = Keys.C OrElse e.KeyCode = Keys.F1 OrElse e.KeyCode = Keys.NumPad1 Then
                            NomeCifra(1) = "C-B#"
                        ElseIf e.KeyCode = Keys.D OrElse e.KeyCode = Keys.F2 OrElse e.KeyCode = Keys.NumPad3 Then
                            NomeCifra(1) = "D"
                        ElseIf e.KeyCode = Keys.E OrElse e.KeyCode = Keys.F3 OrElse e.KeyCode = Keys.NumPad5 Then
                            NomeCifra(1) = "E-Fb"
                        ElseIf e.KeyCode = Keys.F OrElse e.KeyCode = Keys.F4 OrElse e.KeyCode = Keys.NumPad6 Then
                            NomeCifra(1) = "F-E#"
                        ElseIf e.KeyCode = Keys.G OrElse e.KeyCode = Keys.F5 OrElse e.KeyCode = Keys.NumPad8 Then
                            NomeCifra(1) = "G"
                        ElseIf e.KeyCode = Keys.NumPad2 Then
                            NomeCifra(1) = "C#-Db"
                        ElseIf e.KeyCode = Keys.NumPad4 Then
                            NomeCifra(1) = "D#-Eb"
                        ElseIf e.KeyCode = Keys.NumPad7 Then
                            NomeCifra(1) = "F#-Gb"
                        ElseIf e.KeyCode = Keys.NumPad9 Then
                            NomeCifra(1) = "G#-Ab"
                        ElseIf e.KeyCode = Keys.Multiply Then
                            NomeCifra(1) = "A#-Bb"
                        End If

                        DefineValorCorretoIncorreto()
                        MostraNotaCorreta()
                        ExibePontuação()
                    End If
                End If
            End If

            If e.KeyCode = Keys.Escape Then RadioButton1.Checked = True
            If e.KeyCode = Keys.J Then RadioButton2.Checked = True 'J de Jogar

            If e.KeyCode = Keys.ControlKey Then TeclaControlPressionada = 1

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub MemoNotes_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyUp
        Try

            lista.Remove(e.KeyCode)
            CorTecla.Visible = False
            CorTecla2.Visible = False
            If RadioButton2.Checked AndAlso e.KeyCode <> Keys.V AndAlso e.KeyCode <> Keys.R AndAlso e.KeyCode <> Keys.M Then
                If Not CheckBox26.Checked Then
                    ExecutaLoopNotasAleatórias()
                Else
                    Timer1.Enabled = True
                End If
            End If

            If e.KeyCode = Keys.ControlKey Then
                TeclaControlPressionada = 2
                ToolTipNota(0) = Nothing
                SetaPautaTop = 0 : SetaTecladoLeft = 0 : RetanguloBranco = ""
                AtualizaSetasNaPautaETeclado()
            End If

            If e.KeyCode = Keys.V Then CaptaçãoSom = False

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub ColorirEDefinirQualSomSeraTocado()

        Try

            If RadioButton1.Checked Then
                If kk = 0 Then kk = 1
                If kk = 2 OrElse kk = 5 OrElse kk = 7 OrElse kk = 10 OrElse kk = 12 OrElse kk = 14 OrElse kk = 17 OrElse kk = 19 OrElse kk = 22 OrElse kk = 24 OrElse kk = 26 OrElse kk = 29 OrElse kk = 31 OrElse kk = 34 OrElse kk = 36 OrElse kk = 38 OrElse kk = 41 OrElse kk = 43 OrElse kk = 46 OrElse kk = 48 OrElse kk = 50 OrElse kk = 53 OrElse kk = 55 OrElse kk = 58 OrElse kk = 60 OrElse kk = 62 OrElse kk = 65 OrElse kk = 67 OrElse kk = 70 OrElse kk = 72 OrElse kk = 74 OrElse kk = 77 OrElse kk = 79 OrElse kk = 82 OrElse kk = 84 OrElse kk = 86 Then
                    CorTecla2.Location = PosiçãoTecla(kk - 1, 0)
                Else
                    CorTecla.Location = PosiçãoTecla(kk - 1, 0)
                End If
            End If


            If CheckBox9.Checked OrElse (RadioButton2.Checked AndAlso Sonoro.Checked) Then
                SoarNota()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub SoarNota()

        Try
            Try
                If NotaMidi <> 1000 Then MidiPlayer.Play(New NoteOff(0, 1, CByte(NotaMidi), CByte(NumericUpDown2.Value)))
            Catch ex As Exception

                'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
            Finally
                'esta parte do bloco é executada independentemente se algum erro acontecer ou não

            End Try


            If ComboBox1.Text = "Piano Bösendorfer" Then
                mp3 = "Sons\Notas\Piano Bösendorfer\P" & kk & ".mp3"
                'TocaSom()
            ElseIf ComboBox1.Text = "Strings (Trio)" Then
                mp3 = "Sons\Notas\Strings (Trio)\P" & kk & ".mp3"
                'TocaSom()
            Else
                If ComboBox1.Text <> "" Then
                    Timer3.Enabled = False
                    Try
                        MidiPlayer.Play(New ProgramChange(0, 1, CType(ComboBox1.Text.Substring(0, 3), GeneralMidiInstruments)))
                        MidiPlayer.Play(New NoteOn(0, 1, CByte(kk + 20), CByte(NumericUpDown2.Value)))
                    Catch ex As Exception

                        'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
                    Finally
                        'esta parte do bloco é executada independentemente se algum erro acontecer ou não

                    End Try

                    NotaMidi = kk + 20
                    Timer3.Enabled = True
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub CheckBox9_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox9.CheckedChanged, Me.Load
        Try

            If CheckBox9.Checked Then
                PictureBox22.Visible = True
                PictureBox23.Visible = False
                ToolStripMenuItem7.Image = My.Resources.Ligado2
            Else
                PictureBox22.Visible = False
                PictureBox23.Visible = True
                ToolStripMenuItem7.Image = My.Resources.Desligado2
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub CheckBox11_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox11.CheckedChanged, Me.Load
        Try

            If CheckBox11.Checked Then
                PictureBox24.Visible = True
                PictureBox25.Visible = False
                ToolStripMenuItem8.Image = My.Resources.Ligado2
            Else
                PictureBox24.Visible = False
                PictureBox25.Visible = True
                ToolStripMenuItem8.Image = My.Resources.Desligado2
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub PictureBox22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox22.Click
        Try

            CancelaSom()
            CheckBox9.Checked = False
            PictureBox22.Visible = False
            PictureBox23.Visible = True
            ToolStripMenuItem7.Image = My.Resources.Desligado2

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub PictureBox23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox23.Click
        Try

            If Not RadioButton2.Checked Then
                CheckBox9.Checked = True
                PictureBox22.Visible = True
                PictureBox23.Visible = False
                ToolStripMenuItem7.Image = My.Resources.Ligado2
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub PictureBox24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox24.Click
        Try

            CheckBox11.Checked = False
            PictureBox24.Visible = False
            PictureBox25.Visible = True
            ToolStripMenuItem8.Image = My.Resources.Desligado2

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub PictureBox25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox25.Click
        Try

            If Not CifrasDeAcordes.Checked Then
                CheckBox11.Checked = True
                PictureBox24.Visible = True
                PictureBox25.Visible = False
                ToolStripMenuItem8.Image = My.Resources.Ligado2
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub PictureBox26_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox26.MouseDown
        If e.Button = MouseButtons.Left Then
            spinUP()
            spin = 1
            Timer5.Enabled = True
        End If
    End Sub

    Private Sub PictureBox27_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox27.MouseDown
        If e.Button = MouseButtons.Left Then
            spinDOWN()
            spin = 2
            Timer5.Enabled = True
        End If
    End Sub

    Private Sub spinUP()

        Try

            If CDbl(Decimos.Text) <> 99 Then
                If CDbl(Decimos.Text) < 9 Then
                    Decimos.Text = "0" & CDbl(Decimos.Text) + 1
                Else
                    Decimos.Text = CStr(CDbl(Decimos.Text) + 1)
                End If
            Else
                If CDbl(Minutos.Text + Segundos.Text) <> 118 Then
                    Decimos.Text = "00"
                End If
                If CDbl(Segundos.Text) <> 59 Then
                    If CDbl(Segundos.Text) < 9 Then
                        Segundos.Text = "0" & CDbl(Segundos.Text) + 1
                    ElseIf CDbl(Segundos.Text) = 9 Then
                        Segundos.Text = "10"
                    Else
                        Segundos.Text = CStr(CDbl(Segundos.Text) + 1)
                    End If
                Else
                    If CDbl(Minutos.Text) <> 59 Then
                        Segundos.Text = "00"
                        If CDbl(Minutos.Text) < 9 Then
                            Minutos.Text = "0" & CDbl(Minutos.Text) + 1
                        ElseIf CDbl(Minutos.Text) = 9 Then
                            Minutos.Text = "10"
                        Else
                            Minutos.Text = CStr(CDbl(Minutos.Text) + 1)
                        End If
                    Else
                        Decimos.Text = "99"
                    End If
                End If
            End If
            Timer1.Interval = CInt(((CDbl(Minutos.Text) * 60) + CDbl(Segundos.Text) + (CDbl(Decimos.Text) / 100)) * 1000)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub spinDOWN()

        Try

            If Decimos.Text <> "00" Then
                If CDbl(Decimos.Text) <= 10 Then
                    Decimos.Text = "0" & CDbl(Decimos.Text) - 1
                Else
                    Decimos.Text = CStr(CDbl(Decimos.Text) - 1)
                End If
                If Minutos.Text = "00" AndAlso Segundos.Text = "00" AndAlso Decimos.Text = "00" Then Decimos.Text = "01"
            Else
                If Minutos.Text + Segundos.Text <> "00" Then
                    Decimos.Text = "99"
                End If
                If Segundos.Text <> "00" Then
                    If CDbl(Segundos.Text) <= 10 Then
                        Segundos.Text = "0" & CDbl(Segundos.Text) - 1
                    Else
                        Segundos.Text = CStr(CDbl(Segundos.Text) - 1)
                    End If
                Else
                    If Minutos.Text <> "00" Then
                        Segundos.Text = "59"
                        If CDbl(Minutos.Text) <= 10 Then
                            Minutos.Text = "0" & CDbl(Minutos.Text) - 1
                        Else
                            Minutos.Text = CStr(CDbl(Minutos.Text) - 1)
                        End If
                    End If
                End If
            End If

            Timer1.Interval = CInt(((CDbl(Minutos.Text) * 60) + CDbl(Segundos.Text) + (CDbl(Decimos.Text) / 100)) * 1000)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub Timer5_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer5.Tick
        If spin = 1 Then
            spinUP()
        Else
            spinDOWN()
        End If
    End Sub

    Private Sub PictureBox26_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox26.MouseUp
        Timer5.Enabled = False
    End Sub

    Private Sub PictureBox27_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox27.MouseUp
        Timer5.Enabled = False
    End Sub

    Private Sub Minutos_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Minutos.Leave
        Try

            If CDbl(Minutos.Text) > 59 Then Minutos.Text = "59"
            If CDbl(Minutos.Text) < 0 Then Minutos.Text = "00"
            If Minutos.Text = "00" AndAlso Segundos.Text = "00" AndAlso Decimos.Text = "00" Then Decimos.Text = "01"
            If Minutos.Text = "0" Then
                Minutos.Text = "00"
            ElseIf Minutos.Text = "1" Then
                Minutos.Text = "01"
            ElseIf Minutos.Text = "2" Then
                Minutos.Text = "02"
            ElseIf Minutos.Text = "3" Then
                Minutos.Text = "03"
            ElseIf Minutos.Text = "4" Then
                Minutos.Text = "04"
            ElseIf Minutos.Text = "5" Then
                Minutos.Text = "05"
            ElseIf Minutos.Text = "6" Then
                Minutos.Text = "06"
            ElseIf Minutos.Text = "7" Then
                Minutos.Text = "07"
            ElseIf Minutos.Text = "8" Then
                Minutos.Text = "08"
            ElseIf Minutos.Text = "9" Then
                Minutos.Text = "09"
            End If
            Timer1.Interval = CInt(((CDbl(Minutos.Text) * 60) + CDbl(Segundos.Text) + (CDbl(Decimos.Text) / 100)) * 1000)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Segundos_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Segundos.Leave
        Try

            If CDbl(Segundos.Text) > 59 Then Segundos.Text = "59"
            If CDbl(Segundos.Text) < 0 Then Segundos.Text = "00"
            If Minutos.Text = "00" AndAlso Segundos.Text = "00" AndAlso Decimos.Text = "00" Then Decimos.Text = "01"
            If Segundos.Text = "0" Then
                Segundos.Text = "00"
            ElseIf Segundos.Text = "1" Then
                Segundos.Text = "01"
            ElseIf Segundos.Text = "2" Then
                Segundos.Text = "02"
            ElseIf Segundos.Text = "3" Then
                Segundos.Text = "03"
            ElseIf Segundos.Text = "4" Then
                Segundos.Text = "04"
            ElseIf Segundos.Text = "5" Then
                Segundos.Text = "05"
            ElseIf Segundos.Text = "6" Then
                Segundos.Text = "06"
            ElseIf Segundos.Text = "7" Then
                Segundos.Text = "07"
            ElseIf Segundos.Text = "8" Then
                Segundos.Text = "08"
            ElseIf Segundos.Text = "9" Then
                Segundos.Text = "09"
            End If
            Timer1.Interval = CInt(((CDbl(Minutos.Text) * 60) + CDbl(Segundos.Text) + (CDbl(Decimos.Text) / 100)) * 1000)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Decimos_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Decimos.Leave
        Try

            If CDbl(Decimos.Text) > 99 Then Decimos.Text = "99"
            If CDbl(Decimos.Text) < 0 Then Decimos.Text = "00"
            If Minutos.Text = "00" AndAlso Segundos.Text = "00" AndAlso Decimos.Text = "00" Then Decimos.Text = "01"
            If Decimos.Text = "0" Then
                Decimos.Text = "00"
            ElseIf Decimos.Text = "1" Then
                Decimos.Text = "01"
            ElseIf Decimos.Text = "2" Then
                Decimos.Text = "02"
            ElseIf Decimos.Text = "3" Then
                Decimos.Text = "03"
            ElseIf Decimos.Text = "4" Then
                Decimos.Text = "04"
            ElseIf Decimos.Text = "5" Then
                Decimos.Text = "05"
            ElseIf Decimos.Text = "6" Then
                Decimos.Text = "06"
            ElseIf Decimos.Text = "7" Then
                Decimos.Text = "07"
            ElseIf Decimos.Text = "8" Then
                Decimos.Text = "08"
            ElseIf Decimos.Text = "9" Then
                Decimos.Text = "09"
            End If
            Timer1.Interval = CInt(((CDbl(Minutos.Text) * 60) + CDbl(Segundos.Text) + (CDbl(Decimos.Text) / 100)) * 1000)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub SolToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem13.Click
        CheckBox3.Checked = Not CheckBox3.Checked
    End Sub

    Private Sub FáToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem14.Click
        CheckBox4.Checked = Not CheckBox4.Checked
    End Sub

    Private Sub NãoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem44.Click
        RadioButton1.Checked = True
    End Sub

    Private Sub SimToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem45.Click
        RadioButton2.Checked = True
    End Sub

    Private Sub ToolStripMenuItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem7.Click

        Try

            If CheckBox9.Checked Then
                CheckBox9.Checked = False
                PictureBox22.Visible = False
                PictureBox23.Visible = True
                ToolStripMenuItem7.Image = My.Resources.Desligado2

            Else

                If Not RadioButton2.Checked Then
                    CheckBox9.Checked = True
                    PictureBox22.Visible = True
                    PictureBox23.Visible = False
                    ToolStripMenuItem7.Image = My.Resources.Ligado2
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub ToolStripMenuItem8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem8.Click

        Try

            If CheckBox11.Checked Then
                CheckBox11.Checked = False
                PictureBox24.Visible = False
                PictureBox25.Visible = True
                ToolStripMenuItem8.Image = My.Resources.Desligado2
            Else

                CheckBox11.Checked = True
                PictureBox24.Visible = True
                PictureBox25.Visible = False
                ToolStripMenuItem8.Image = My.Resources.Ligado2
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub ToolStripMenuItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem6.Click

        Try

            If ToolStripMenuItem6.Text = "Pausar" Then
                Continuar.BringToFront()
                Timer1.Enabled = False
                ToolStripMenuItem6.Text = "Continuar"
                ToolStripMenuItem6.Image = My.Resources.Continuar
            Else
                Pausar.BringToFront()
                Timer1.Enabled = True
                ToolStripMenuItem6.Text = "Pausar"
                ToolStripMenuItem6.Image = My.Resources.Pausar
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub NotaExata_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NotaExata.CheckedChanged, Me.Load
        Try

            If Not NotaExata.Checked AndAlso Not NotaGenérica.Checked Then
                NotaExata.Checked = True : ToolStripMenuItem46.Checked = True
                NotaGenérica.Checked = False : ToolStripMenuItem47.Checked = False
            ElseIf NotaExata.Checked Then
                NotaGenérica.Checked = False : ToolStripMenuItem47.Checked = False
                ToolStripMenuItem46.Checked = True
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub NotaGenérica_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NotaGenérica.CheckedChanged, Me.Load
        Try

            If Not NotaGenérica.Checked AndAlso Not NotaExata.Checked Then
                NotaGenérica.Checked = True : ToolStripMenuItem47.Checked = True
                NotaExata.Checked = False : ToolStripMenuItem46.Checked = False
            ElseIf NotaGenérica.Checked Then
                NotaExata.Checked = False : ToolStripMenuItem46.Checked = False
                ToolStripMenuItem47.Checked = True
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub NotaExataToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem46.Click
        NotaExata.Checked = True
    End Sub

    Private Sub CifraToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem47.Click
        NotaGenérica.Checked = True
    End Sub

    Private Sub CheckBox13_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Load, DóPrimeiraLinhaToolStripMenuItem1.Click, DóPrimeiraLinhaToolStripMenuItem3.Click

        Try

            If DóPrimeiraLinhaToolStripMenuItem1.Checked OrElse DóPrimeiraLinhaToolStripMenuItem3.Checked Then
                CheckBox3.Checked = False
                CheckBox4.Checked = False
                DóSegundaLinhaToolStripMenuItem1.Checked = False
                DóTerceiraLinhaToolStripMenuItem1.Checked = False
                DóQuartaLinhaToolStripMenuItem1.Checked = False
                FáQuintaLinhaToolStripMenuItem1.Checked = False
                DóQuintaLinhaToolStripMenuItem1.Checked = False
                FáTerceiraLinhaToolStripMenuItem1.Checked = False
                SolPrimeiraLinhaToolStripMenuItem1.Checked = False
                DóPrimeiraLinhaToolStripMenuItem3.Checked = True
                DóPrimeiraLinhaToolStripMenuItem1.Checked = True
                ToolStripMenuItem16.Checked = False
                ToolStripMenuItem17.Checked = False
                ToolStripMenuItem18.Checked = False
                ToolStripMenuItem20.Checked = False
                ToolStripMenuItem19.Checked = False
                ToolStripMenuItem21.Checked = False
                ToolStripMenuItem22.Checked = False
                ToolStripMenuItem13.Checked = False
                ToolStripMenuItem14.Checked = False
            End If

            ExibirTrackBars()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub CheckBox14_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Load, ToolStripMenuItem16.Click, DóSegundaLinhaToolStripMenuItem1.Click

        Try

            If DóSegundaLinhaToolStripMenuItem1.Checked OrElse ToolStripMenuItem16.Checked Then
                CheckBox3.Checked = False
                CheckBox4.Checked = False
                DóPrimeiraLinhaToolStripMenuItem1.Checked = False
                DóTerceiraLinhaToolStripMenuItem1.Checked = False
                DóQuartaLinhaToolStripMenuItem1.Checked = False
                FáQuintaLinhaToolStripMenuItem1.Checked = False
                DóQuintaLinhaToolStripMenuItem1.Checked = False
                FáTerceiraLinhaToolStripMenuItem1.Checked = False
                SolPrimeiraLinhaToolStripMenuItem1.Checked = False
                DóPrimeiraLinhaToolStripMenuItem3.Checked = False
                ToolStripMenuItem16.Checked = True
                DóSegundaLinhaToolStripMenuItem1.Checked = True
                ToolStripMenuItem17.Checked = False
                ToolStripMenuItem18.Checked = False
                ToolStripMenuItem20.Checked = False
                ToolStripMenuItem19.Checked = False
                ToolStripMenuItem21.Checked = False
                ToolStripMenuItem22.Checked = False
                ToolStripMenuItem13.Checked = False
                ToolStripMenuItem14.Checked = False
            End If

            ExibirTrackBars()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub CheckBox15_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Load, ToolStripMenuItem17.Click, DóTerceiraLinhaToolStripMenuItem1.Click

        Try

            If DóTerceiraLinhaToolStripMenuItem1.Checked OrElse ToolStripMenuItem17.Checked Then
                CheckBox3.Checked = False
                CheckBox4.Checked = False
                DóSegundaLinhaToolStripMenuItem1.Checked = False
                DóPrimeiraLinhaToolStripMenuItem1.Checked = False
                DóQuartaLinhaToolStripMenuItem1.Checked = False
                FáQuintaLinhaToolStripMenuItem1.Checked = False
                DóQuintaLinhaToolStripMenuItem1.Checked = False
                FáTerceiraLinhaToolStripMenuItem1.Checked = False
                SolPrimeiraLinhaToolStripMenuItem1.Checked = False
                DóPrimeiraLinhaToolStripMenuItem3.Checked = False
                ToolStripMenuItem16.Checked = False
                ToolStripMenuItem17.Checked = True
                DóTerceiraLinhaToolStripMenuItem1.Checked = True
                ToolStripMenuItem18.Checked = False
                ToolStripMenuItem20.Checked = False
                ToolStripMenuItem19.Checked = False
                ToolStripMenuItem21.Checked = False
                ToolStripMenuItem22.Checked = False
                ToolStripMenuItem13.Checked = False
                ToolStripMenuItem14.Checked = False
            End If

            ExibirTrackBars()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub CheckBox16_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Load, ToolStripMenuItem18.Click, DóQuartaLinhaToolStripMenuItem1.Click

        Try

            If DóQuartaLinhaToolStripMenuItem1.Checked OrElse ToolStripMenuItem18.Checked Then
                CheckBox3.Checked = False
                CheckBox4.Checked = False
                DóSegundaLinhaToolStripMenuItem1.Checked = False
                DóTerceiraLinhaToolStripMenuItem1.Checked = False
                DóPrimeiraLinhaToolStripMenuItem1.Checked = False
                FáQuintaLinhaToolStripMenuItem1.Checked = False
                DóQuintaLinhaToolStripMenuItem1.Checked = False
                FáTerceiraLinhaToolStripMenuItem1.Checked = False
                SolPrimeiraLinhaToolStripMenuItem1.Checked = False
                DóPrimeiraLinhaToolStripMenuItem3.Checked = False
                ToolStripMenuItem16.Checked = False
                ToolStripMenuItem17.Checked = False
                ToolStripMenuItem18.Checked = True
                DóQuartaLinhaToolStripMenuItem1.Checked = True
                ToolStripMenuItem20.Checked = False
                ToolStripMenuItem19.Checked = False
                ToolStripMenuItem21.Checked = False
                ToolStripMenuItem22.Checked = False
                ToolStripMenuItem13.Checked = False
                ToolStripMenuItem14.Checked = False
            End If

            ExibirTrackBars()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub CheckBox17_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Load, ToolStripMenuItem20.Click, FáTerceiraLinhaToolStripMenuItem1.Click

        Try

            If FáTerceiraLinhaToolStripMenuItem1.Checked OrElse ToolStripMenuItem20.Checked Then
                CheckBox3.Checked = False
                CheckBox4.Checked = False
                DóSegundaLinhaToolStripMenuItem1.Checked = False
                DóTerceiraLinhaToolStripMenuItem1.Checked = False
                DóQuartaLinhaToolStripMenuItem1.Checked = False
                DóPrimeiraLinhaToolStripMenuItem1.Checked = False
                DóQuintaLinhaToolStripMenuItem1.Checked = False
                FáQuintaLinhaToolStripMenuItem1.Checked = False
                SolPrimeiraLinhaToolStripMenuItem1.Checked = False
                DóPrimeiraLinhaToolStripMenuItem3.Checked = False
                ToolStripMenuItem16.Checked = False
                ToolStripMenuItem17.Checked = False
                ToolStripMenuItem18.Checked = False
                ToolStripMenuItem20.Checked = True
                FáTerceiraLinhaToolStripMenuItem1.Checked = True
                ToolStripMenuItem19.Checked = False
                ToolStripMenuItem21.Checked = False
                ToolStripMenuItem22.Checked = False
                ToolStripMenuItem13.Checked = False
                ToolStripMenuItem14.Checked = False
            End If

            ExibirTrackBars()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub CheckBox18_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Load, ToolStripMenuItem19.Click, DóQuintaLinhaToolStripMenuItem1.Click

        Try

            If DóQuintaLinhaToolStripMenuItem1.Checked OrElse ToolStripMenuItem19.Checked Then
                CheckBox3.Checked = False
                CheckBox4.Checked = False
                DóSegundaLinhaToolStripMenuItem1.Checked = False
                DóTerceiraLinhaToolStripMenuItem1.Checked = False
                DóQuartaLinhaToolStripMenuItem1.Checked = False
                DóPrimeiraLinhaToolStripMenuItem1.Checked = False
                FáQuintaLinhaToolStripMenuItem1.Checked = False
                FáTerceiraLinhaToolStripMenuItem1.Checked = False
                SolPrimeiraLinhaToolStripMenuItem1.Checked = False
                DóPrimeiraLinhaToolStripMenuItem3.Checked = False
                ToolStripMenuItem16.Checked = False
                ToolStripMenuItem17.Checked = False
                ToolStripMenuItem18.Checked = False
                ToolStripMenuItem20.Checked = False
                ToolStripMenuItem19.Checked = True
                DóQuintaLinhaToolStripMenuItem1.Checked = True
                ToolStripMenuItem21.Checked = False
                ToolStripMenuItem22.Checked = False
                ToolStripMenuItem13.Checked = False
                ToolStripMenuItem14.Checked = False
            End If

            ExibirTrackBars()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub CheckBox19_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Load, ToolStripMenuItem21.Click, FáQuintaLinhaToolStripMenuItem1.Click

        Try

            If FáQuintaLinhaToolStripMenuItem1.Checked OrElse ToolStripMenuItem21.Checked Then
                CheckBox3.Checked = False
                CheckBox4.Checked = False
                DóSegundaLinhaToolStripMenuItem1.Checked = False
                DóTerceiraLinhaToolStripMenuItem1.Checked = False
                DóQuartaLinhaToolStripMenuItem1.Checked = False
                DóPrimeiraLinhaToolStripMenuItem1.Checked = False
                DóQuintaLinhaToolStripMenuItem1.Checked = False
                FáTerceiraLinhaToolStripMenuItem1.Checked = False
                SolPrimeiraLinhaToolStripMenuItem1.Checked = False
                DóPrimeiraLinhaToolStripMenuItem3.Checked = False
                ToolStripMenuItem16.Checked = False
                ToolStripMenuItem17.Checked = False
                ToolStripMenuItem18.Checked = False
                ToolStripMenuItem20.Checked = False
                ToolStripMenuItem19.Checked = False
                ToolStripMenuItem21.Checked = True
                FáQuintaLinhaToolStripMenuItem1.Checked = True
                ToolStripMenuItem22.Checked = False
                ToolStripMenuItem13.Checked = False
                ToolStripMenuItem14.Checked = False
            End If

            ExibirTrackBars()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub CheckBox24_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Load, ToolStripMenuItem22.Click, SolPrimeiraLinhaToolStripMenuItem1.Click

        Try

            If SolPrimeiraLinhaToolStripMenuItem1.Checked OrElse ToolStripMenuItem22.Checked Then
                CheckBox3.Checked = False
                CheckBox4.Checked = False
                DóSegundaLinhaToolStripMenuItem1.Checked = False
                DóTerceiraLinhaToolStripMenuItem1.Checked = False
                DóQuartaLinhaToolStripMenuItem1.Checked = False
                DóPrimeiraLinhaToolStripMenuItem1.Checked = False
                DóQuintaLinhaToolStripMenuItem1.Checked = False
                FáTerceiraLinhaToolStripMenuItem1.Checked = False
                FáQuintaLinhaToolStripMenuItem1.Checked = False
                DóPrimeiraLinhaToolStripMenuItem3.Checked = False
                ToolStripMenuItem16.Checked = False
                ToolStripMenuItem17.Checked = False
                ToolStripMenuItem18.Checked = False
                ToolStripMenuItem20.Checked = False
                ToolStripMenuItem19.Checked = False
                ToolStripMenuItem21.Checked = False
                ToolStripMenuItem22.Checked = True
                SolPrimeiraLinhaToolStripMenuItem1.Checked = True
                ToolStripMenuItem13.Checked = False
                ToolStripMenuItem14.Checked = False
            End If

            ExibirTrackBars()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub ExibirTrackBars()

        Try

            If CheckBox3.Checked OrElse CheckBox4.Checked Then
                TrackBar4.Visible = True
                TrackBar2.Visible = True
            Else
                TrackBar4.Visible = False
                TrackBar2.Visible = False
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub DóPrimeiraLinhaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem15.Click
        DóPrimeiraLinhaToolStripMenuItem1.Checked = Not DóPrimeiraLinhaToolStripMenuItem1.Checked
    End Sub

    Private Sub DóSegundaLinhaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem16.Click
        DóSegundaLinhaToolStripMenuItem1.Checked = Not DóSegundaLinhaToolStripMenuItem1.Checked
    End Sub

    Private Sub DóTerceiraLinhaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem17.Click
        DóTerceiraLinhaToolStripMenuItem1.Checked = Not DóTerceiraLinhaToolStripMenuItem1.Checked
    End Sub

    Private Sub DóQuartaLinhaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem18.Click
        DóQuartaLinhaToolStripMenuItem1.Checked = Not DóQuartaLinhaToolStripMenuItem1.Checked
    End Sub

    Private Sub DóQuintaLinhaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem19.Click
        DóQuintaLinhaToolStripMenuItem1.Checked = Not DóQuintaLinhaToolStripMenuItem1.Checked
    End Sub

    Private Sub FáTerceiraLinhaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem20.Click
        FáTerceiraLinhaToolStripMenuItem1.Checked = Not FáTerceiraLinhaToolStripMenuItem1.Checked
    End Sub

    Private Sub FáQuintaLinhaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem21.Click
        FáQuintaLinhaToolStripMenuItem1.Checked = Not FáQuintaLinhaToolStripMenuItem1.Checked
    End Sub

    Private Sub SolPrimeiraLinhaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem22.Click
        SolPrimeiraLinhaToolStripMenuItem1.Checked = Not SolPrimeiraLinhaToolStripMenuItem1.Checked
    End Sub

    Private Sub AtualizaClaves(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles CheckBox3.MouseUp, CheckBox4.MouseUp, RadioButton1.MouseUp, RadioButton2.MouseUp, ToolStripMenuItem22.MouseUp, ToolStripMenuItem21.MouseUp, ToolStripMenuItem20.MouseUp, ToolStripMenuItem19.MouseUp, ToolStripMenuItem18.MouseUp, ToolStripMenuItem17.MouseUp, ToolStripMenuItem16.MouseUp, ToolStripMenuItem14.MouseUp, ToolStripMenuItem13.MouseUp, SolPrimeiraLinhaToolStripMenuItem1.MouseUp, FáTerceiraLinhaToolStripMenuItem1.MouseUp, FáQuintaLinhaToolStripMenuItem1.MouseUp, DóTerceiraLinhaToolStripMenuItem1.MouseUp, DóSegundaLinhaToolStripMenuItem1.MouseUp, DóQuintaLinhaToolStripMenuItem1.MouseUp, DóQuartaLinhaToolStripMenuItem1.MouseUp, DóPrimeiraLinhaToolStripMenuItem1.MouseUp, ToolStripMenuItem15.MouseUp, DóPrimeiraLinhaToolStripMenuItem3.MouseUp, Violão.MouseUp

        'Dim Rect As New Rectangle(0, 70, 325, 502)
        'Me.Invalidate(Rect)

        GerarImagemDeFundo()
        GerarImagemDeFundo2()

    End Sub

    Private Sub TrackBar1_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar1.Scroll, TrackBar2.Scroll, TrackBar4.Scroll, TrackBar5.Scroll, Button1.Click
        Dim Rect5 As New Rectangle(231, 155, 7, 400)
        Me.Invalidate(Rect5)
    End Sub

    Private Sub TrackBar1_Scroll_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar1.Scroll
        Try

            If TrackBar1.Value < TrackBar5.Value + 5 Then TrackBar1.Value = TrackBar5.Value + 5

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub TrackBar5_Scroll_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar5.Scroll
        Try

            If TrackBar5.Value > TrackBar1.Value - 5 Then TrackBar5.Value = TrackBar1.Value - 5

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub TrackBar4_Scroll_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar4.Scroll
        Try

            If TrackBar4.Value < TrackBar2.Value + 5 Then TrackBar4.Value = TrackBar2.Value + 5

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub TrackBar2_Scroll_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar2.Scroll
        Try

            If TrackBar2.Value > TrackBar4.Value - 5 Then TrackBar2.Value = TrackBar4.Value - 5

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Cms1_ItemClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles Cms1.ItemClicked

        Try

            Dim midiError As Integer
            Dim MenuItem As ToolStripMenuItem
            For Each MenuItem In Cms1.Items
                MenuItem.Checked = False
            Next
            MenuItem = CType(e.ClickedItem, ToolStripMenuItem)
            MenuItem.Checked = True

            ' On scanne le port Midi In
            midiError = WinMM.midiInOpen(hMidiIn, CInt(MenuItem.Tag), DelgMidiIn, 0, &H30000)
            midiError = WinMM.midiInStart(hMidiIn)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Protected Sub MidiInProc(ByVal MidiInHandle As Int32, _
                         ByVal NewMsg As Int32, _
                         ByVal Instance As Int32, _
                         ByVal wParam As Int32, _
                         ByVal lParam As Int32)
        Try

            If wParam > 255 Then
                'Trace.WriteLine("Msg " & wParam)
                Dim b() As Byte = BitConverter.GetBytes(wParam)
                canal = CByte(b(0) And &HF) ' recupere le canal
                Select Case b(0) And &HF0
                    Case &H90
                        If b(2) > 0 Then
                            Me.TouchOn(b(1), b(2))
                        Else
                            Me.TouchOff(b(1), b(2))
                        End If
                    Case &H80 And &HF0
                        Me.TouchOff(b(1), b(2))
                End Select
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub TouchOn(ByVal [Param] As Byte, ByVal [KeyVelocity] As Byte)

        Try

            If Me.KeyCol([Param]).InvokeRequired Then
                Me.KeyCol([Param]).Invoke(DelgParamON, New Object() {[Param], [KeyVelocity]})
            Else

                If CheckBox11.Checked Then
                    If ComboBox1.Text = "Piano Bösendorfer" Then
                        mp3 = "Sons\Notas\Piano Bösendorfer\P" & [Param] - 20 & ".mp3"
                        'TocaSom()
                    ElseIf ComboBox1.Text = "Strings (Trio)" Then
                        mp3 = "Sons\Notas\Strings (Trio)\P" & [Param] - 20 & ".mp3"
                        'TocaSom()
                    Else
                        Try
                            MidiPlayer.Play(New ProgramChange(0, 1, CType(ComboBox1.Text.Substring(0, 3), GeneralMidiInstruments)))
                            MidiPlayer.Play(New NoteOn(0, 1, CByte([Param]), CByte([KeyVelocity])))
                        Catch ex As Exception

                            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
                        Finally
                            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

                        End Try
                    End If
                End If


                If RadioButton2.Checked Then Timer2.Enabled = True
                'Me.KeyCol([Param]).Visible = True
                NomeCifra(1) = "" 'NomeTecla = ""
                Dim cControl() As Control = Me.Controls.Find("P" & [Param] - 20, True)
                Dim pb As PictureBox = DirectCast(cControl(0), PictureBox)
                If [Param] = 21 Then
                    pb.Image = My.Resources.Tecla_BrancaPressionada : NomeCifra(1) = "A"
                ElseIf [Param] = 23 OrElse [Param] = 35 OrElse [Param] = 47 OrElse [Param] = 59 OrElse [Param] = 71 OrElse [Param] = 83 OrElse [Param] = 95 OrElse [Param] = 107 Then
                    pb.Image = My.Resources.Tecla_BrancaPressionada3 : NomeCifra(1) = "B-Cb"
                    If CifrasDeAcordes.Checked Then [Param] = 23
                ElseIf [Param] = 24 OrElse [Param] = 36 OrElse [Param] = 48 OrElse [Param] = 60 OrElse [Param] = 72 OrElse [Param] = 84 OrElse [Param] = 96 Then
                    pb.Image = My.Resources.Tecla_BrancaPressionada : NomeCifra(1) = "C-B#"
                    If CifrasDeAcordes.Checked Then [Param] = 24
                ElseIf [Param] = 26 OrElse [Param] = 38 OrElse [Param] = 50 OrElse [Param] = 62 OrElse [Param] = 74 OrElse [Param] = 86 OrElse [Param] = 98 Then
                    pb.Image = My.Resources.Tecla_BrancaPressionada2 : NomeCifra(1) = "D"
                    If CifrasDeAcordes.Checked Then [Param] = 26
                ElseIf [Param] = 28 OrElse [Param] = 40 OrElse [Param] = 52 OrElse [Param] = 64 OrElse [Param] = 76 OrElse [Param] = 88 OrElse [Param] = 100 Then
                    pb.Image = My.Resources.Tecla_BrancaPressionada3 : NomeCifra(1) = "E-Fb"
                    If CifrasDeAcordes.Checked Then [Param] = 28
                ElseIf [Param] = 29 OrElse [Param] = 41 OrElse [Param] = 53 OrElse [Param] = 65 OrElse [Param] = 77 OrElse [Param] = 89 OrElse [Param] = 101 Then
                    pb.Image = My.Resources.Tecla_BrancaPressionada : NomeCifra(1) = "F-E#"
                    If CifrasDeAcordes.Checked Then [Param] = 29
                ElseIf [Param] = 31 OrElse [Param] = 43 OrElse [Param] = 55 OrElse [Param] = 67 OrElse [Param] = 79 OrElse [Param] = 91 OrElse [Param] = 103 Then
                    pb.Image = My.Resources.Tecla_BrancaPressionada2 : NomeCifra(1) = "G"
                    If CifrasDeAcordes.Checked Then [Param] = 31
                ElseIf [Param] = 33 OrElse [Param] = 45 OrElse [Param] = 57 OrElse [Param] = 69 OrElse [Param] = 81 OrElse [Param] = 93 OrElse [Param] = 105 Then
                    pb.Image = My.Resources.Tecla_BrancaPressionada2 : NomeCifra(1) = "A"
                    If CifrasDeAcordes.Checked Then [Param] = 21
                ElseIf [Param] = 108 Then
                    pb.Image = My.Resources.Tecla_BrancaPressionada4 : NomeCifra(1) = "C-B#"
                    If CifrasDeAcordes.Checked Then [Param] = 24
                Else
                    pb.Image = My.Resources.Tecla_PretaPressionada
                    If [Param] = 22 OrElse [Param] = 34 OrElse [Param] = 46 OrElse [Param] = 58 OrElse [Param] = 70 OrElse [Param] = 82 OrElse [Param] = 94 OrElse [Param] = 106 Then
                        NomeCifra(1) = "A#-Bb"
                        If CifrasDeAcordes.Checked Then [Param] = 22
                    ElseIf [Param] = 25 OrElse [Param] = 37 OrElse [Param] = 49 OrElse [Param] = 61 OrElse [Param] = 73 OrElse [Param] = 85 OrElse [Param] = 97 Then
                        NomeCifra(1) = "C#-Db"
                        If CifrasDeAcordes.Checked Then [Param] = 25
                    ElseIf [Param] = 27 OrElse [Param] = 39 OrElse [Param] = 51 OrElse [Param] = 63 OrElse [Param] = 75 OrElse [Param] = 87 OrElse [Param] = 99 Then
                        NomeCifra(1) = "D#-Eb"
                        If CifrasDeAcordes.Checked Then [Param] = 27
                    ElseIf [Param] = 30 OrElse [Param] = 42 OrElse [Param] = 54 OrElse [Param] = 66 OrElse [Param] = 78 OrElse [Param] = 90 OrElse [Param] = 102 Then
                        NomeCifra(1) = "F#-Gb"
                        If CifrasDeAcordes.Checked Then [Param] = 30
                    ElseIf [Param] = 32 OrElse [Param] = 44 OrElse [Param] = 56 OrElse [Param] = 68 OrElse [Param] = 80 OrElse [Param] = 92 OrElse [Param] = 104 Then
                        NomeCifra(1) = "G#-Ab"
                        If CifrasDeAcordes.Checked Then [Param] = 32
                    End If
                End If


                If NomeTecla(0) = "" Then
                    NomeTecla(0) = "P" & [Param] - 20
                    NomeTecla(1) = ""
                ElseIf CInt(TextBox1.Text) > 1 AndAlso NomeTecla(1) = "" Then
                    NomeTecla(1) = "P" & [Param] - 20
                    NomeTecla(2) = ""
                ElseIf CInt(TextBox1.Text) > 2 AndAlso NomeTecla(2) = "" Then
                    NomeTecla(2) = "P" & [Param] - 20
                    NomeTecla(3) = ""
                ElseIf CInt(TextBox1.Text) > 3 AndAlso NomeTecla(3) = "" Then
                    NomeTecla(3) = "P" & [Param] - 20
                End If


                If RadioButton2.Checked AndAlso NomeTecla(CInt(TextBox1.Text) - 1) <> "" Then
                    NomeTeclaFinal = NomeTecla(0) & NomeTecla(1) & NomeTecla(2) & NomeTecla(3)
                    DefineValorCorretoIncorreto()
                    MostraNotaCorreta()
                    ExibePontuação()
                    NomeTecla(0) = "" : NomeTecla(1) = "" : NomeTecla(2) = "" : NomeTecla(3) = ""

                    If RadioButton2.Checked AndAlso NomeTecla(0) = "" Then
                        If Not CheckBox26.Checked Then
                            ExecutaLoopNotasAleatórias()
                        Else
                            Timer1.Enabled = True
                        End If
                        Timer2.Enabled = False
                    End If
                End If

            End If
            ' Me.KeyCol([Param]).BackColor = ccolor(canal)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub TouchOff(ByVal [Param] As Byte, ByVal [KeyVelocity] As Byte)
        Try

            If Me.KeyCol([Param]).InvokeRequired Then
                Me.KeyCol([Param]).Invoke(DelgParamOff, New Object() {[Param], [KeyVelocity]})
            Else
                'Me.KeyCol([Param]).Visible = False
                Dim cControl() As Control = Me.Controls.Find("P" & [Param] - 20, True)
                Dim pb As PictureBox = DirectCast(cControl(0), PictureBox)

                If [Param] = 22 OrElse [Param] = 34 OrElse [Param] = 46 OrElse [Param] = 58 OrElse [Param] = 70 OrElse [Param] = 82 OrElse [Param] = 94 OrElse [Param] = 106 _
                OrElse [Param] = 25 OrElse [Param] = 37 OrElse [Param] = 49 OrElse [Param] = 61 OrElse [Param] = 73 OrElse [Param] = 85 OrElse [Param] = 97 _
                OrElse [Param] = 27 OrElse [Param] = 39 OrElse [Param] = 51 OrElse [Param] = 63 OrElse [Param] = 75 OrElse [Param] = 87 OrElse [Param] = 99 _
                OrElse [Param] = 30 OrElse [Param] = 42 OrElse [Param] = 54 OrElse [Param] = 66 OrElse [Param] = 78 OrElse [Param] = 90 OrElse [Param] = 102 _
                OrElse [Param] = 32 OrElse [Param] = 44 OrElse [Param] = 56 OrElse [Param] = 68 OrElse [Param] = 80 OrElse [Param] = 92 OrElse [Param] = 104 Then
                    pb.Image = My.Resources.Tecla_Preta
                Else
                    pb.Image = My.Resources.Tecla_Branca
                End If
            End If
            Try
                MidiPlayer.Play(New NoteOff(0, 1, CByte([Param]), CByte([KeyVelocity])))
            Catch ex As Exception

                'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
            Finally
                'esta parte do bloco é executada independentemente se algum erro acontecer ou não

            End Try

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    'Private Sub TouchON(ByVal Num As Integer)
    '    KeyCol(Num).Visible = True
    'End Sub
    'Private Sub TouchOFF(ByVal Num As Integer)
    '    KeyCol(Num).Visible = False
    'End Sub

    Private Sub DefineNomeMenu(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SolMaiorMiMenorFToolStripMenuItem.Click, SiMaiorToolStripMenuItem.Click, RéMaiorSiMenorFCToolStripMenuItem.Click, RébMaiorSibMenorToolStripMenuItem.Click, MiMaiorDóMenorFCGDToolStripMenuItem.Click, LáMaiorFáMenorFCGToolStripMenuItem.Click, GbBbEbAbDbGbCbToolStripMenuItem.Click, FDmBbToolStripMenuItem.Click, FáMaiorRéMenorFCGDAEToolStripMenuItem.Click, EbCmBbEbAbToolStripMenuItem.Click, DóMaiorLáMenorToolStripMenuItem.Click, BbGmBbEbToolStripMenuItem.Click, AbMaiorToolStripMenuItem.Click, NenhumToolStripMenuItem.Click, NenhumToolStripMenuItem1.Click, GEmFToolStripMenuItem.Click, GbEbmBbEbAbDbGbCbToolStripMenuItem.Click, FDmFCGDAEToolStripMenuItem.Click, FDmBbToolStripMenuItem1.Click, ECmFCGDToolStripMenuItem.Click, EbCmBbEbAbToolStripMenuItem1.Click, DBmFCToolStripMenuItem.Click, DbBbmBbEbAbDbGbToolStripMenuItem.Click, CAmToolStripMenuItem.Click, BGmFCGDAToolStripMenuItem.Click, BbGmBbEbToolStripMenuItem1.Click, AFmFCGToolStripMenuItem.Click, AbFmBbEbAbDbToolStripMenuItem.Click, CbAbmBbEbAbDbGbCbFbToolStripMenuItem.Click, CAmFCGDAEBToolStripMenuItem.Click, CbAbmBbEbAbDbGbCbFbToolStripMenuItem1.Click, CAmFCGDAEBToolStripMenuItem1.Click

        Try

            Dim Menu As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
            NomeDoMenu = Menu.Text
            My.Settings.NovoValorNomeDoMenu = Menu.Text
            TomToolStripMenuItem.Image = Menu.Image
            If NomeDoMenu = "Nenhum" Then
                Acidentes.Enabled = True
                My.Settings.NovoValorCoresArmadura = False
            Else
                Acidentes.Enabled = False
            End If
            'Dim Rect As New Rectangle(775, 170, 52, 33)
            'Me.Invalidate(Rect)
            'Dim Rect2 As New Rectangle(0, 70, 325, 500)
            'Me.Invalidate(Rect2)

            GerarImagemDeFundo()
            GerarImagemDeFundo2()


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub CheckBox26_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox26.CheckedChanged
        Try

            If CheckBox26.Checked Then
                ExecutaLoopNotasAleatórias()
            Else
                Timer1.Enabled = False
            End If
            GerarNotasAutomaticamenteNoIntervaloDeTempoToolStripMenuItem.Checked = CheckBox26.Checked

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub GerarNotasAutomaticamenteNoIntervaloDeTempoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GerarNotasAutomaticamenteNoIntervaloDeTempoToolStripMenuItem.Click
        CheckBox26.Checked = Not CheckBox26.Checked
    End Sub

    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox2.Click
        Try

            If Not Sonoro.Checked Then
                If CInt(TextBox1.Text) < 4 Then TextBox1.Text = CStr(CDbl(TextBox1.Text) + 1)
                ValorTotalDeNotas = CInt(TextBox1.Text)
                If RadioButton2.Checked Then ExecutaLoopNotasAleatórias()
                CheckedQuantidadeDeNotas()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click
        Try

            If Not Sonoro.Checked Then
                If CInt(TextBox1.Text) > 1 Then TextBox1.Text = CStr(CDbl(TextBox1.Text) - 1)
                ValorTotalDeNotas = CInt(TextBox1.Text)
                If RadioButton2.Checked Then ExecutaLoopNotasAleatórias()
                CheckedQuantidadeDeNotas()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Try

            NomeTeclaFinal = NomeTecla(0) & NomeTecla(1) & NomeTecla(2) & NomeTecla(3)
            DefineValorCorretoIncorreto()
            MostraNotaCorreta()
            ExibePontuação()
            NomeTecla(0) = "" : NomeTecla(1) = "" : NomeTecla(2) = "" : NomeTecla(3) = ""
            If Not CheckBox26.Checked Then
                ExecutaLoopNotasAleatórias()
            Else
                Timer1.Enabled = True
            End If
            Timer2.Enabled = False

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub GravaPontuação()

        Try

            Timer1.Enabled = False
            If MsgBox("Deseja gravar esta pontuação?", MsgBoxStyle.YesNo, "Atenção") = MsgBoxResult.Yes Then

                If Pontuação.Text = "" Then
                    ValorPercentual = "0,000"
                Else
                    ValorPercentual = FormatNumber(((zz / vv) * 100), 3)
                End If

                Dim elapsed As TimeSpan = Cronometro.Elapsed

                If vv = 0 Then
                    TempoMedio = "0,000"
                Else
                    TempoMedio = FormatNumber(((Cronometro.ElapsedMilliseconds / vv) / 1000), 3)
                End If

                ' Create our string that holds our new Person information.
                Dim cr As String = Environment.NewLine

                Dim ModoJogo As String = ""
                Dim ModoJogo2 As String = ""
                If NotaExata.Checked Then
                    ModoJogo = "Pauta: Nota Exata"
                ElseIf NotaGenérica.Checked Then
                    ModoJogo = "Pauta: Nota Genérica"
                ElseIf CifrasDeAcordes.Checked Then
                    ModoJogo = "Cifras de Acordes"
                End If

                If Visual.Checked And Sonoro.Checked Then
                    ModoJogo2 = "(Visual e Sonoro)"
                ElseIf Visual.Checked And Not Sonoro.Checked Then
                    ModoJogo2 = "(Visual)"
                ElseIf Not Visual.Checked And Sonoro.Checked Then
                    ModoJogo2 = "(Sonoro)"
                End If

                ModoJogo = ModoJogo & " " & ModoJogo2

                Dim xdoc As XDocument = XDocument.Load("HistóricoPontuação.xml")
                Dim elemento As System.Xml.Linq.XElement = xdoc.Descendants("Pontuação").Last

                Dim NotasSelecionadas As String
                NotasSelecionadas = ""
                If My.Settings.NovoValorExibirNotas(0) = "True" Then NotasSelecionadas = NotasSelecionadas & "C "
                If My.Settings.NovoValorExibirNotas(1) = "True" Then NotasSelecionadas = NotasSelecionadas & "D "
                If My.Settings.NovoValorExibirNotas(2) = "True" Then NotasSelecionadas = NotasSelecionadas & "E "
                If My.Settings.NovoValorExibirNotas(3) = "True" Then NotasSelecionadas = NotasSelecionadas & "F "
                If My.Settings.NovoValorExibirNotas(4) = "True" Then NotasSelecionadas = NotasSelecionadas & "G "
                If My.Settings.NovoValorExibirNotas(5) = "True" Then NotasSelecionadas = NotasSelecionadas & "A "
                If My.Settings.NovoValorExibirNotas(6) = "True" Then NotasSelecionadas = NotasSelecionadas & "B "


                Dim NovaPontuação As String = _
                    "<Pontuação>" & cr & _
                    "    <ID>" & Format(CDbl(elemento.Descendants("ID").Value) + 1, "000000") & "</ID>" & cr & _
                    "    <Data>" & FormatDateTime(Now, DateFormat.ShortDate) & "</Data>" & cr & _
                    "    <HoraInicio>" & HoraInicio & "</HoraInicio>" & cr & _
                    "    <HoraFim>" & HoraFim & "</HoraFim>" & cr & _
                    "    <Percentual>" & ValorPercentual & "</Percentual>" & cr & _
                    "    <Acertos>" & zz & "</Acertos>" & cr & _
                    "    <Erros>" & (vv - zz) & "</Erros>" & cr & _
                    "    <Total>" & vv & "</Total>" & cr & _
                    "    <Tempo>" & String.Format("{0:00}:{1:00}:{2:00}.{3:000}", _
                                               Math.Floor(elapsed.TotalHours), _
                                               elapsed.Minutes, _
                                               elapsed.Seconds, _
                                               elapsed.Milliseconds) & "</Tempo>" & cr & _
                    "    <TempoMédio>" & TempoMedio & "</TempoMédio>" & cr & _
                    "    <Tom>" & NomeDoMenu & "</Tom>" & cr & _
                    "    <Clave>" & NomeClave & "</Clave>" & cr & _
                    "    <ModoDeJogo>" & ModoJogo & "</ModoDeJogo>" & cr & _
                    "    <QtdeNotas>" & TextBox1.Text & "</QtdeNotas>" & cr & _
                    "    <NotasSelecionadas>" & NotasSelecionadas & "</NotasSelecionadas>" & cr & _
                    "</Pontuação>"


                ' Load the XmlDocument.
                Dim xd As New XmlDocument()
                xd.Load("HistóricoPontuação.xml")

                ' Create a new XmlDocumentFragment for our document.
                Dim docFrag As XmlDocumentFragment = xd.CreateDocumentFragment()
                ' The Xml for this fragment is our newPerson string.
                docFrag.InnerXml = NovaPontuação
                ' The root element of our file is found using
                ' the DocumentElement property of the XmlDocument.
                Dim root As XmlNode = xd.DocumentElement
                ' Append our new Person to the root element.
                root.AppendChild(docFrag)

                ' Save the Xml.
                xd.Save("HistóricoPontuação.xml")
            Else
                Exit Sub
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub Pontuação_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Pontuação.Click, ExibirHistóricoDaPontuaçãoToolStripMenuItem.Click

        Try

            HistóricoPontuação.Show()
            Dim xmlFile As XmlReader
            xmlFile = XmlReader.Create("HistóricoPontuação.xml", New XmlReaderSettings())
            Dim ds As New DataSet
            ds.ReadXml(xmlFile)

            Dim dv As DataView 'ordenar ID em ordem decrescente
            dv = New DataView(ds.Tables(0), "", "ID Desc", DataViewRowState.CurrentRows)

            HistóricoPontuação.DataGridView1.DataSource = dv
            HistóricoPontuação.DataGridView1.Columns(0).Width = 60
            HistóricoPontuação.DataGridView1.Columns(1).Width = 80
            HistóricoPontuação.DataGridView1.Columns(2).Width = 80
            HistóricoPontuação.DataGridView1.Columns(3).Width = 80
            HistóricoPontuação.DataGridView1.Columns(4).Width = 80
            HistóricoPontuação.DataGridView1.Columns(5).Width = 60
            HistóricoPontuação.DataGridView1.Columns(6).Width = 60
            HistóricoPontuação.DataGridView1.Columns(7).Width = 60
            HistóricoPontuação.DataGridView1.Columns(9).Width = 70
            HistóricoPontuação.DataGridView1.Columns(12).Width = 150
            HistóricoPontuação.DataGridView1.Columns(13).Width = 60



            GerarGráfico()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Public Sub GerarGráfico()

        Try

            HistóricoPontuação.Chart1.Series("Percentual Acertos").Points.Clear()
            HistóricoPontuação.Chart1.Series("Tempo Médio").Points.Clear()
            HistóricoPontuação.Chart1.Annotations.Clear()


            Dim value1, value2 As String
            Dim QtdeLinhas, Período, Valor As Integer
            Valor = 0


            If CStr(HistóricoPontuação.ComboBox1.Text) = "" OrElse CStr(HistóricoPontuação.ComboBox1.Text) = "Todos" OrElse CDbl(HistóricoPontuação.ComboBox1.Text) > HistóricoPontuação.DataGridView1.RowCount Then
                Período = HistóricoPontuação.DataGridView1.RowCount
            Else
                Período = CInt(HistóricoPontuação.ComboBox1.Text)
            End If


            QtdeLinhas = Período



            Do While QtdeLinhas > 0
                value1 = CStr(HistóricoPontuação.DataGridView1.Item(4, QtdeLinhas - 1).Value)
                value2 = CStr(HistóricoPontuação.DataGridView1.Item(9, QtdeLinhas - 1).Value)
                HistóricoPontuação.Chart1.Series("Percentual Acertos").Points.AddY(value1.Replace(",", "."))
                HistóricoPontuação.Chart1.Series("Tempo Médio").Points.AddY(value2.Replace(",", "."))

                If Valor = 0 Then
                    HistóricoPontuação.Chart1.Series("Percentual Acertos").Points(Valor).AxisLabel = CStr(HistóricoPontuação.DataGridView1.Item(1, QtdeLinhas - 1).Value)
                Else
                    If CStr(HistóricoPontuação.DataGridView1.Item(1, QtdeLinhas - 1).Value) <> CStr(HistóricoPontuação.DataGridView1.Item(1, QtdeLinhas).Value) Then
                        HistóricoPontuação.Chart1.Series("Percentual Acertos").Points(Valor).AxisLabel = CStr(HistóricoPontuação.DataGridView1.Item(1, QtdeLinhas - 1).Value)
                    Else
                        HistóricoPontuação.Chart1.Series("Percentual Acertos").Points(Valor).AxisLabel = " "
                    End If
                End If


                If CDbl(value1) < 95 AndAlso CDbl(value1) >= 90 Then
                    HistóricoPontuação.Chart1.Series("Percentual Acertos").Points(Valor).MarkerColor = Color.FromArgb(10, 213, 10)
                ElseIf CDbl(value1) < 90 AndAlso CDbl(value1) >= 80 Then
                    HistóricoPontuação.Chart1.Series("Percentual Acertos").Points(Valor).MarkerColor = Color.FromArgb(136, 255, 136)
                ElseIf CDbl(value1) < 80 AndAlso CDbl(value1) >= 70 Then
                    HistóricoPontuação.Chart1.Series("Percentual Acertos").Points(Valor).MarkerColor = Color.FromArgb(255, 255, 132)
                ElseIf CDbl(value1) < 70 AndAlso CDbl(value1) >= 60 Then
                    HistóricoPontuação.Chart1.Series("Percentual Acertos").Points(Valor).MarkerColor = Color.FromArgb(255, 255, 10)
                ElseIf CDbl(value1) < 60 AndAlso CDbl(value1) >= 40 Then
                    HistóricoPontuação.Chart1.Series("Percentual Acertos").Points(Valor).MarkerColor = Color.FromArgb(255, 80, 80)
                ElseIf CDbl(value1) < 40 AndAlso CDbl(value1) >= 20 Then
                    HistóricoPontuação.Chart1.Series("Percentual Acertos").Points(Valor).MarkerColor = Color.FromArgb(255, 40, 40)
                ElseIf CDbl(value1) < 20 AndAlso CDbl(value1) >= 0 Then
                    HistóricoPontuação.Chart1.Series("Percentual Acertos").Points(Valor).MarkerColor = Color.FromArgb(255, 10, 10)
                Else
                    HistóricoPontuação.Chart1.Series("Percentual Acertos").Points(Valor).MarkerColor = Color.FromArgb(10, 166, 10)
                End If


                QtdeLinhas -= 1
                Valor += 1
            Loop

            Dim font As New Font(HistóricoPontuação.DataGridView1.DefaultCellStyle.Font.FontFamily, 10, FontStyle.Bold)

            HistóricoPontuação.DataGridView1.Columns(4).DefaultCellStyle.Font = font


            Dim pointValueName As String = "Y"
            Dim valueToFind As [Double] = 0
            ' Find maximum
            Dim maxPoint As DataPoint = HistóricoPontuação.Chart1.Series("Percentual Acertos").Points.FindMaxByValue(pointValueName)
            valueToFind = maxPoint.GetValueByName(pointValueName)
            'Color all the points with the specified value
            For Each dataPoint As DataPoint In HistóricoPontuação.Chart1.Series("Percentual Acertos").Points.FindAllByValue(valueToFind, pointValueName)
                dataPoint.MarkerStyle = MarkerStyle.Star5
                dataPoint.MarkerSize = 15

                Dim anotação1 As New CalloutAnnotation()
                With anotação1
                    .AnchorDataPoint = dataPoint
                    .Text = "Maior" + ControlChars.Lf + "nota"
                    .BackColor = Color.FromArgb(15, 255, 0, 0)
                    .SmartLabelStyle.Enabled = False
                    .CalloutStyle = CalloutStyle.RoundedRectangle
                End With
                HistóricoPontuação.Chart1.Annotations.Add(anotação1)
            Next

            ' Find minimum
            Dim minPoint As DataPoint = HistóricoPontuação.Chart1.Series("Tempo Médio").Points.FindMinByValue(pointValueName)
            valueToFind = minPoint.GetValueByName(pointValueName)

            'Color all the points with the specified value
            For Each dataPoint As DataPoint In HistóricoPontuação.Chart1.Series("Tempo Médio").Points.FindAllByValue(valueToFind, pointValueName)
                dataPoint.MarkerStyle = MarkerStyle.Star5
                dataPoint.MarkerSize = 15

                Dim anotação2 As New CalloutAnnotation()
                With anotação2
                    .AnchorDataPoint = dataPoint
                    .Text = "Menor" + ControlChars.Lf + "tempo"
                    .BackColor = Color.FromArgb(15, 255, 0, 0)
                    .SmartLabelStyle.Enabled = False
                    .CalloutStyle = CalloutStyle.RoundedRectangle
                End With
                HistóricoPontuação.Chart1.Annotations.Add(anotação2)
            Next

            HistóricoPontuação.Chart1.ChartAreas(0).AxisX.ScaleView.ZoomReset(0)

            If Período > 20 Then HistóricoPontuação.Chart1.ChartAreas("ChartArea1").AxisX.ScaleView.Zoom(Período - 20, Período) 'define o ínicio e o fim do zoom no gráfico

            Dim mean1 As Double = HistóricoPontuação.Chart1.DataManipulator.Statistics.Mean("Percentual Acertos")

            HistóricoPontuação.Chart1.ChartAreas(0).AxisY.StripLines(0).IntervalOffset = CInt(mean1) - 1
            HistóricoPontuação.Chart1.ChartAreas(0).AxisY.StripLines(0).StripWidth = 2
            HistóricoPontuação.Chart1.ChartAreas(0).AxisY.StripLines(0).BackColor = Color.FromArgb(150, 255, 100, 0)
            HistóricoPontuação.Chart1.ChartAreas(0).AxisY.StripLines(0).Text = "Média" + ControlChars.Lf + CStr(CInt(mean1))
            HistóricoPontuação.Chart1.ChartAreas(0).AxisY.StripLines(0).TextLineAlignment = StringAlignment.Far

            Dim mean2 As Double = HistóricoPontuação.Chart1.DataManipulator.Statistics.Mean("Tempo Médio")

            HistóricoPontuação.Chart1.ChartAreas(0).AxisY.StripLines(1).IntervalOffset = (120 / (5 / mean2)) - 1.4
            HistóricoPontuação.Chart1.ChartAreas(0).AxisY.StripLines(1).StripWidth = 2
            HistóricoPontuação.Chart1.ChartAreas(0).AxisY.StripLines(1).BackColor = Color.FromArgb(150, 0, 100, 255)
            HistóricoPontuação.Chart1.ChartAreas(0).AxisY.StripLines(1).Text = "Média" + ControlChars.Lf + CStr(FormatNumber(mean2, 3))
            HistóricoPontuação.Chart1.ChartAreas(0).AxisY.StripLines(1).TextLineAlignment = StringAlignment.Far

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub ApagarHistóricoDaPontuaçãoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ApagarHistóricoDaPontuaçãoToolStripMenuItem.Click
        Try

            If MsgBox("Deseja realmente apagar o histórico?", MsgBoxStyle.YesNo, "Atenção") = MsgBoxResult.Yes Then
                Dim writer As New XmlTextWriter("HistóricoPontuação.xml", System.Text.Encoding.UTF8)
                writer.WriteStartDocument(True)
                writer.Formatting = Formatting.Indented
                writer.Indentation = 2
                writer.WriteStartElement("HistóricoPontuação")
                'createNode("", "", "", writer)
                writer.WriteEndElement()
                writer.WriteEndDocument()
                writer.Close()
            Else
                Exit Sub
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub ToolStripMenuItem37_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem37.Click
        TextBox1.Text = "1"
        CheckedQuantidadeDeNotas()
    End Sub

    Private Sub ToolStripMenuItem50_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem50.Click
        TextBox1.Text = "2"
        CheckedQuantidadeDeNotas()
    End Sub

    Private Sub ToolStripMenuItem51_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem51.Click
        TextBox1.Text = "3"
        CheckedQuantidadeDeNotas()
    End Sub

    Private Sub ToolStripMenuItem52_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem52.Click
        TextBox1.Text = "4"
        CheckedQuantidadeDeNotas()
    End Sub

    Private Sub CheckedQuantidadeDeNotas() Handles Me.Load

        Try

            If TextBox1.Text = "1" Then
                ToolStripMenuItem37.Checked = True
                ToolStripMenuItem50.Checked = False
                ToolStripMenuItem51.Checked = False
                ToolStripMenuItem52.Checked = False
            ElseIf TextBox1.Text = "2" Then
                ToolStripMenuItem37.Checked = False
                ToolStripMenuItem50.Checked = True
                ToolStripMenuItem51.Checked = False
                ToolStripMenuItem52.Checked = False
            ElseIf TextBox1.Text = "3" Then
                ToolStripMenuItem37.Checked = False
                ToolStripMenuItem50.Checked = False
                ToolStripMenuItem51.Checked = True
                ToolStripMenuItem52.Checked = False
            ElseIf TextBox1.Text = "4" Then
                ToolStripMenuItem37.Checked = False
                ToolStripMenuItem50.Checked = False
                ToolStripMenuItem51.Checked = False
                ToolStripMenuItem52.Checked = True
            End If

            If CDbl(TextBox1.Text) > 1 Then 'se for mais do que 1 nota, então opção disponível será somente para NOTA EXATA
                NotaGenérica.Checked = False
                NotaGenérica.Enabled = False
                NotaExata.Checked = True
            Else
                NotaGenérica.Enabled = True
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub Visual_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Visual.CheckedChanged, Sonoro.CheckedChanged
        If Not Visual.Checked AndAlso Not Sonoro.Checked Then
            Visual.Checked = True
            Sonoro.Checked = False
        End If
    End Sub

    Private Sub Sonoro_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Sonoro.MouseUp
        If Sonoro.Checked Then
            TextBox1.Text = "1"
            ValorTotalDeNotas = CInt(TextBox1.Text)
            CheckedQuantidadeDeNotas()
        End If
    End Sub

    Private Sub RepetirSom_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RepetirSom.Click
        'If Sonoro.Checked AndAlso RadioButton2.Checked Then SoarNota()
        If RadioButton2.Checked Then SoarNota()
    End Sub

    Private Sub Pauta_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Pauta.CheckedChanged, Me.Load
        If Not Pauta.Checked AndAlso Not CifrasDeAcordes.Checked Then
            Pauta.Checked = True : PautaToolStripMenuItem.Checked = True
            CifrasDeAcordes.Checked = False : CifraToolStripMenuItem1.Checked = False
        ElseIf Pauta.Checked Then
            CifrasDeAcordes.Checked = False : CifraToolStripMenuItem1.Checked = False
            PautaToolStripMenuItem.Checked = True
        End If
    End Sub

    Private Sub PautaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PautaToolStripMenuItem.Click
        Pauta.Checked = True
    End Sub

    Private Sub CifraToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CifraToolStripMenuItem1.Click
        CifrasDeAcordes.Checked = True
    End Sub

    Private Sub CifrasDeAcordes_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CifrasDeAcordes.CheckedChanged
        Try


            If Not CifrasDeAcordes.Checked AndAlso Not Pauta.Checked Then
                CifrasDeAcordes.Checked = True : CifraToolStripMenuItem1.Checked = True
                Pauta.Checked = False : PautaToolStripMenuItem.Checked = False
            ElseIf CifrasDeAcordes.Checked Then
                CheckBox11.Checked = False
                Pauta.Checked = False : PautaToolStripMenuItem.Checked = False
                CifraToolStripMenuItem1.Checked = True
            End If

            'atualiza notas na tela
            'Dim Rect2 As New Rectangle(121, 130, 84, 433)
            'Me.Invalidate(Rect2)
            'Dim Rect3 As New Rectangle(350, 260, 300, 170)
            'Me.Invalidate(Rect3)

            GerarImagemDeFundo()
            GerarImagemDeFundo2()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub PictureBox3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox3.Click
        OpcoesMemonotes.ShowDialog()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click, AlternarModoAbrangênciaToolStripMenuItem.Click
        Try

            If My.Settings.NovoValorAbrangenciaDasNotas = True Then
                My.Settings.NovoValorAbrangenciaDasNotas = False
            Else
                My.Settings.NovoValorAbrangenciaDasNotas = True
            End If

            LayoutAbrangenciaDasNotas()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub LayoutAbrangenciaDasNotas() Handles Me.Load

        Try

            If My.Settings.NovoValorAbrangenciaDasNotas = True Then

                AlternarModoAbrangênciaToolStripMenuItem.Checked = False
                If TrackBar1.Value < 47 Then TrackBar1.Value = 47
                If TrackBar5.Value > 43 Then TrackBar5.Value = 43
                If TrackBar4.Value < 20 Then TrackBar4.Value = 20
                If TrackBar2.Value > 16 Then TrackBar2.Value = 16

                With TrackBar1
                    .Location = New Point(269, 158)
                    .Height = 127
                    .Width = 28
                    .Maximum = 67
                    .Minimum = 47
                End With


                With TrackBar5
                    .Location = New Point(269, 278)
                    .Height = 82
                    .Width = 28
                    .Maximum = 43
                    .Minimum = 32
                End With


                With TrackBar4
                    .Location = New Point(269, 373)
                    .Height = 82
                    .Width = 28
                    .Maximum = 31
                    .Minimum = 20
                End With

                With TrackBar2
                    .Location = New Point(269, 448)
                    .Height = 102
                    .Width = 28
                    .Maximum = 16
                    .Minimum = 1
                End With
            Else

                AlternarModoAbrangênciaToolStripMenuItem.Checked = True

                With TrackBar1
                    .Location = New Point(256, 158)
                    .Height = 177
                    .Width = 28
                    .Maximum = 67
                    .Minimum = 37
                End With

                With TrackBar5
                    .Location = New Point(285, 183)
                    .Height = 177
                    .Width = 28
                    .Maximum = 62
                    .Minimum = 32
                End With


                With TrackBar4
                    .Location = New Point(256, 373)
                    .Height = 152
                    .Width = 28
                    .Maximum = 31
                    .Minimum = 6
                End With

                With TrackBar2
                    .Location = New Point(285, 398)
                    .Height = 152
                    .Width = 28
                    .Maximum = 26
                    .Minimum = 1
                End With

            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub CancelaSom()
        Try
            MidiPlayer.Play(New NoteOff(0, 1, CByte(NotaMidi), CByte(NumericUpDown2.Value)))
        Catch ex As Exception

            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Timer3_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer3.Tick
        Try

            CancelaSom()
            Timer3.Enabled = False

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click, MaisOpçõesToolStripMenuItem.Click
        Try

            OpcoesMemonotes.ShowDialog()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Public Sub AtualizaRegiões()
        Try

            Dim Rect1 As New Rectangle(0, 70, 325, 502)
            Invalidate(Rect1)
            Dim Rect2 As New Rectangle(89, 707, 885, 52)
            Invalidate(Rect2)
            Dim Rect3 As New Rectangle(350, 260, 300, 170)
            Invalidate(Rect3)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub QuantidadeDeLoopsQuePoderãoSerExecutados()
        Try

            Timer1.Enabled = False
            ToolStripMenuItem6.Text = "Continuar"
            ToolStripMenuItem6.Image = My.Resources.Continuar

            MsgBox("Esta configuração está gerando um looping infinito." & vbCrLf & _
                   "Soluções: aumente um pouco a área de abrangência, ou selecione mais notas para serem exibidas." & vbCrLf & _
                    vbCrLf & _
                    "Após fazer estes ajustes clique no botão ""Continuar""", MsgBoxStyle.Information, "Atenção")
            Continuar.BringToFront()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub


    '====================================================================================
    '====================================================================================
    '====================================================================================
    '====================================================================================
    '====================================================================================
    '====================================================================================

    Private m_isListenning As Boolean = False

    Public ReadOnly Property IsListenning() As Boolean
        Get
            Return m_isListenning
        End Get

    End Property

    Public Sub New()
        InitializeComponent()
    End Sub

    Private frequencyInfoSource As FrequencyInfoSource

    Public Sub StopListenning()

        Try

            m_isListenning = False
            frequencyInfoSource.[Stop]()
            RemoveHandler frequencyInfoSource.FrequencyDetected, New EventHandler(Of FrequencyDetectedEventArgs)(AddressOf frequencyInfoSource_FrequencyDetected)
            frequencyInfoSource = Nothing


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Public Sub StartListenning(ByVal device As SoundCaptureDevice)

        Try

            m_isListenning = True
            frequencyInfoSource = New SoundFrequencyInfoSource(device)
            AddHandler frequencyInfoSource.FrequencyDetected, New EventHandler(Of FrequencyDetectedEventArgs)(AddressOf frequencyInfoSource_FrequencyDetected)
            frequencyInfoSource.Listen()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub frequencyInfoSource_FrequencyDetected(ByVal sender As Object, ByVal e As FrequencyDetectedEventArgs)

        Try

            If InvokeRequired Then
                BeginInvoke(New EventHandler(Of FrequencyDetectedEventArgs)(AddressOf frequencyInfoSource_FrequencyDetected), sender, e)
            Else
                UpdateFrequecyDisplays(e.Frequency)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub IniciaTimerParaReativarCaptacaoSom()
        Try

            contadorReativaCaptacao = 0

            'If CaptaçãoSom = False Then 'so executa o codigo a seguir se captacao do som nao estiver ativado
            'inicia o timer, se acaso uma nota ficar muito tempo parada, indicando "congelamento", fará com que a captacao do som seja reiniciada automaticamente
            'If Not Timer6.Enabled Then NotaCantadaArmazenada2 = frequencyTextBox.Text

            'If frequencyTextBox.Text = NotaCantadaArmazenada2 Then
            'If Not Timer6.Enabled Then
            'Timer6.Enabled = True
            'End If
            'Else
            'Timer6.Enabled = False
            'End If
            'Else
            'Timer6.Enabled = False
            'End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub UpdateFrequecyDisplays(ByVal frequency As Double)

        Try

            If frequency > 0 Then

                SinalDetectado = True


                'por algum motivo, quando LineInput 2 é escolhido, a frequencia fica uma oitava abaixo do normal
                'a linha abaixo corrige este problema
                If My.Settings.NovoValorLineInput(2) = "2" Then frequency *= 2


                CurrentFrequency = frequency
                IndicadorDeFrequencia.FrequenciesScale1.SignalDetected = True
                IndicadorDeFrequencia.FrequenciesScale1.CurrentFrequency = frequency

                frequencyTextBox.Enabled = True
                frequencyTextBox.Text = frequency.ToString("f3")

                Dim noteName As String = Nothing
                FindClosestNote(frequency, closestFrequency, noteName)
                IndicadorDeFrequencia.FrequenciesScale1.ClosestFrequency = closestFrequency

                IndicadorDeFrequencia.FrequenciesScale1.NoteName = noteName & Octave
                closeFrequencyTextBox.Enabled = True
                closeFrequencyTextBox.Text = closestFrequency.ToString("f3")
                noteNameTextBox.Enabled = True
                noteNameTextBox.Text = noteName & Octave




                IniciaTimerParaReativarCaptacaoSom()




                'TextBox2.Text = NomeTecla(0) & "      P" & kkFinal(0)


                'Ajusta Nota Cantada para Exibir na Pauta
                If noteName & Octave = "A0" OrElse noteName & Octave = "A#0" Then
                    NotaCantadaAjustadaParaPauta = 16
                ElseIf noteName & Octave = "B0" OrElse noteName & Octave = "B#0" OrElse noteName & Octave = "Bb0" Then
                    NotaCantadaAjustadaParaPauta = 17


                ElseIf noteName & Octave = "C1" OrElse noteName & Octave = "C#1" OrElse noteName & Octave = "Cb1" Then
                    NotaCantadaAjustadaParaPauta = 18
                ElseIf noteName & Octave = "D1" OrElse noteName & Octave = "D#1" OrElse noteName & Octave = "Db1" Then
                    NotaCantadaAjustadaParaPauta = 19
                ElseIf noteName & Octave = "E1" OrElse noteName & Octave = "E#1" OrElse noteName & Octave = "Eb1" Then
                    NotaCantadaAjustadaParaPauta = 20
                ElseIf noteName & Octave = "F1" OrElse noteName & Octave = "F#1" OrElse noteName & Octave = "Fb1" Then
                    NotaCantadaAjustadaParaPauta = 21
                ElseIf noteName & Octave = "G1" OrElse noteName & Octave = "G#1" OrElse noteName & Octave = "Gb1" Then
                    NotaCantadaAjustadaParaPauta = 22
                ElseIf noteName & Octave = "A1" OrElse noteName & Octave = "A#1" OrElse noteName & Octave = "Ab1" Then
                    NotaCantadaAjustadaParaPauta = 23
                ElseIf noteName & Octave = "B1" OrElse noteName & Octave = "B#1" OrElse noteName & Octave = "Bb1" Then
                    NotaCantadaAjustadaParaPauta = 24


                ElseIf noteName & Octave = "C2" OrElse noteName & Octave = "C#2" OrElse noteName & Octave = "Cb2" Then
                    NotaCantadaAjustadaParaPauta = 25
                ElseIf noteName & Octave = "D2" OrElse noteName & Octave = "D#2" OrElse noteName & Octave = "Db2" Then
                    NotaCantadaAjustadaParaPauta = 26
                ElseIf noteName & Octave = "E2" OrElse noteName & Octave = "E#2" OrElse noteName & Octave = "Eb2" Then
                    NotaCantadaAjustadaParaPauta = 27
                ElseIf noteName & Octave = "F2" OrElse noteName & Octave = "F#2" OrElse noteName & Octave = "Fb2" Then
                    NotaCantadaAjustadaParaPauta = 28
                ElseIf noteName & Octave = "G2" OrElse noteName & Octave = "G#2" OrElse noteName & Octave = "Gb2" Then
                    NotaCantadaAjustadaParaPauta = 29
                ElseIf noteName & Octave = "A2" OrElse noteName & Octave = "A#2" OrElse noteName & Octave = "Ab2" Then
                    NotaCantadaAjustadaParaPauta = 30
                ElseIf noteName & Octave = "B2" OrElse noteName & Octave = "B#2" OrElse noteName & Octave = "Bb2" Then
                    NotaCantadaAjustadaParaPauta = 31


                ElseIf noteName & Octave = "C3" OrElse noteName & Octave = "C#3" OrElse noteName & Octave = "Cb3" Then
                    NotaCantadaAjustadaParaPauta = 32
                ElseIf noteName & Octave = "D3" OrElse noteName & Octave = "D#3" OrElse noteName & Octave = "Db3" Then
                    NotaCantadaAjustadaParaPauta = 33
                ElseIf noteName & Octave = "E3" OrElse noteName & Octave = "E#3" OrElse noteName & Octave = "Eb3" Then
                    NotaCantadaAjustadaParaPauta = 34
                ElseIf noteName & Octave = "F3" OrElse noteName & Octave = "F#3" OrElse noteName & Octave = "Fb3" Then
                    NotaCantadaAjustadaParaPauta = 35
                ElseIf noteName & Octave = "G3" OrElse noteName & Octave = "G#3" OrElse noteName & Octave = "Gb3" Then
                    NotaCantadaAjustadaParaPauta = 36
                ElseIf noteName & Octave = "A3" OrElse noteName & Octave = "A#3" OrElse noteName & Octave = "Ab3" Then
                    NotaCantadaAjustadaParaPauta = 37
                ElseIf noteName & Octave = "B3" OrElse noteName & Octave = "B#3" OrElse noteName & Octave = "Bb3" Then
                    NotaCantadaAjustadaParaPauta = 38


                ElseIf noteName & Octave = "C4" OrElse noteName & Octave = "C#4" OrElse noteName & Octave = "Cb4" Then
                    NotaCantadaAjustadaParaPauta = 39
                ElseIf noteName & Octave = "D4" OrElse noteName & Octave = "D#4" OrElse noteName & Octave = "Db4" Then
                    NotaCantadaAjustadaParaPauta = 40
                ElseIf noteName & Octave = "E4" OrElse noteName & Octave = "E#4" OrElse noteName & Octave = "Eb4" Then
                    NotaCantadaAjustadaParaPauta = 41
                ElseIf noteName & Octave = "F4" OrElse noteName & Octave = "F#4" OrElse noteName & Octave = "Fb4" Then
                    NotaCantadaAjustadaParaPauta = 42
                ElseIf noteName & Octave = "G4" OrElse noteName & Octave = "G#4" OrElse noteName & Octave = "Gb4" Then
                    NotaCantadaAjustadaParaPauta = 43
                ElseIf noteName & Octave = "A4" OrElse noteName & Octave = "A#4" OrElse noteName & Octave = "Ab4" Then
                    NotaCantadaAjustadaParaPauta = 44
                ElseIf noteName & Octave = "B4" OrElse noteName & Octave = "B#4" OrElse noteName & Octave = "Bb4" Then
                    NotaCantadaAjustadaParaPauta = 45


                ElseIf noteName & Octave = "C5" OrElse noteName & Octave = "C#5" OrElse noteName & Octave = "Cb5" Then
                    NotaCantadaAjustadaParaPauta = 46
                ElseIf noteName & Octave = "D5" OrElse noteName & Octave = "D#5" OrElse noteName & Octave = "Db5" Then
                    NotaCantadaAjustadaParaPauta = 47
                ElseIf noteName & Octave = "E5" OrElse noteName & Octave = "E#5" OrElse noteName & Octave = "Eb5" Then
                    NotaCantadaAjustadaParaPauta = 48
                ElseIf noteName & Octave = "F5" OrElse noteName & Octave = "F#5" OrElse noteName & Octave = "Fb5" Then
                    NotaCantadaAjustadaParaPauta = 49
                ElseIf noteName & Octave = "G5" OrElse noteName & Octave = "G#5" OrElse noteName & Octave = "Gb5" Then
                    NotaCantadaAjustadaParaPauta = 50
                ElseIf noteName & Octave = "A5" OrElse noteName & Octave = "A#5" OrElse noteName & Octave = "Ab5" Then
                    NotaCantadaAjustadaParaPauta = 51
                ElseIf noteName & Octave = "B5" OrElse noteName & Octave = "B#5" OrElse noteName & Octave = "Bb5" Then
                    NotaCantadaAjustadaParaPauta = 52


                ElseIf noteName & Octave = "C6" OrElse noteName & Octave = "C#6" OrElse noteName & Octave = "Cb6" Then
                    NotaCantadaAjustadaParaPauta = 53
                ElseIf noteName & Octave = "D6" OrElse noteName & Octave = "D#6" OrElse noteName & Octave = "Db6" Then
                    NotaCantadaAjustadaParaPauta = 54
                ElseIf noteName & Octave = "E6" OrElse noteName & Octave = "E#5" OrElse noteName & Octave = "Eb6" Then
                    NotaCantadaAjustadaParaPauta = 55
                ElseIf noteName & Octave = "F6" OrElse noteName & Octave = "F#6" OrElse noteName & Octave = "Fb6" Then
                    NotaCantadaAjustadaParaPauta = 56
                ElseIf noteName & Octave = "G6" OrElse noteName & Octave = "G#6" OrElse noteName & Octave = "Gb6" Then
                    NotaCantadaAjustadaParaPauta = 57
                ElseIf noteName & Octave = "A6" OrElse noteName & Octave = "A#6" OrElse noteName & Octave = "Ab6" Then
                    NotaCantadaAjustadaParaPauta = 58
                ElseIf noteName & Octave = "B6" OrElse noteName & Octave = "B#6" OrElse noteName & Octave = "Bb6" Then
                    NotaCantadaAjustadaParaPauta = 59


                ElseIf noteName & Octave = "C7" OrElse noteName & Octave = "C#7" OrElse noteName & Octave = "Cb7" Then
                    NotaCantadaAjustadaParaPauta = 60
                ElseIf noteName & Octave = "D7" OrElse noteName & Octave = "D#7" OrElse noteName & Octave = "Db7" Then
                    NotaCantadaAjustadaParaPauta = 61
                ElseIf noteName & Octave = "E7" OrElse noteName & Octave = "E#7" OrElse noteName & Octave = "Eb7" Then
                    NotaCantadaAjustadaParaPauta = 62
                ElseIf noteName & Octave = "F7" OrElse noteName & Octave = "F#7" OrElse noteName & Octave = "Fb7" Then
                    NotaCantadaAjustadaParaPauta = 63
                ElseIf noteName & Octave = "G7" OrElse noteName & Octave = "G#7" OrElse noteName & Octave = "Gb7" Then
                    NotaCantadaAjustadaParaPauta = 64
                ElseIf noteName & Octave = "A7" OrElse noteName & Octave = "A#7" OrElse noteName & Octave = "Ab7" Then
                    NotaCantadaAjustadaParaPauta = 65
                ElseIf noteName & Octave = "B7" OrElse noteName & Octave = "B#7" OrElse noteName & Octave = "Bb7" Then
                    NotaCantadaAjustadaParaPauta = 66


                ElseIf noteName & Octave = "C8" OrElse noteName & Octave = "Cb8" Then
                    NotaCantadaAjustadaParaPauta = 67

                End If



                'atualiza indicador de frequencia na pauta
                Dim Rect2 As New Rectangle(191, 29, 63, 556)
                Me.Invalidate(Rect2)

                If (CaptaçãoSom = True AndAlso My.Settings.NovoValorFonteCaptaçãoÁudio(0) = "True") OrElse My.Settings.NovoValorFonteCaptaçãoÁudio(1) = "True" Then
                    'a variável CaptaçãoSom fará com que este código seja executado somente se a tecla V (Voz) estiver pressionada
                    'a não ser que a opção escolhida seja "Violão" ao invés de "Voz"
                    If Not Timer4.Enabled Then NotaCantadaArmazenada = NomeTecla(0)

                    If NomeTecla(0) = NotaCantadaArmazenada AndAlso sliderBrush1 Is ActiveSliderBrush1 Then
                        If Not Timer4.Enabled Then

                            'se captação de áudio for VOZ então a resposta tem que ser mantida estável durante
                            '1 segundo, para que o sistema passe para a próxima nota do exercício
                            'caso captação seja VIOLÃO então o tempo necessário pode ser menor
                            If My.Settings.NovoValorFonteCaptaçãoÁudio(0) = "True" Then
                                Timer4.Interval = CInt(My.Settings.NovoValorTimers(12))
                            Else
                                Timer4.Interval = CInt(My.Settings.NovoValorTimers(13))
                            End If


                            Timer4.Enabled = True


                        End If
                    Else
                        Timer4.Enabled = False
                    End If
                Else
                    Timer4.Enabled = False
                End If

            Else
                SinalDetectado = False
                IndicadorDeFrequencia.FrequenciesScale1.SignalDetected = False

                frequencyTextBox.Enabled = False
                closeFrequencyTextBox.Enabled = False
                noteNameTextBox.Enabled = False

                CorTecla3.Left = -100
                CorTecla4.Left = -100

                IniciaTimerParaReativarCaptacaoSom()

            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Shared NoteNames As String() = {"A", "A#", "B", "C", "C#", "D", "D#", "E", "F", "F#", "G", "G#"}
    Shared ToneStep As Double = Math.Pow(2, 1.0 / 12)
    Shared Octave As Integer = 0

    Private Sub FindClosestNote(ByVal frequency As Double, _
 <Out()> ByRef closestFrequency As Double, <Out()> ByRef noteName As String)

        Try

            Const AFrequency As Double = 440.0
            Const ToneIndexOffsetToPositives As Integer = 120

            Dim toneIndex As Integer = CInt(Math.Truncate(Math.Round(Math.Log(frequency / AFrequency, ToneStep))))


            If NomeDoMenu = "C - Am" OrElse _
                NomeDoMenu = "G - Em (F#)" OrElse _
                NomeDoMenu = "D - Bm (F# - C#)" OrElse _
                NomeDoMenu = "A - F#m (F# - C# - G#)" OrElse _
                NomeDoMenu = "E - C#m (F# - C# - G# - D#)" OrElse _
                NomeDoMenu = "B - G#m (F# - C# - G# - D# - A#)" Then

                NoteNames = {"A", "A#", "B", "C", "C#", "D", "D#", "E", "F", "F#", "G", "G#"}

            ElseIf NomeDoMenu = "F# - D#m (F# - C# - G# - D# - A# - E#)" Then

                NoteNames = {"A", "A#", "B", "C", "C#", "D", "D#", "E", "E#", "F#", "G", "G#"}

            ElseIf NomeDoMenu = "C# - A#m (F# - C# - G# - D# - A# - E# - B#)" Then

                NoteNames = {"A", "A#", "B", "B#", "C#", "D", "D#", "E", "E#", "F#", "G", "G#"}

            ElseIf NomeDoMenu = "Cb - Abm (Bb - Eb - Ab - Db - Gb - Cb - Fb)" Then

                NoteNames = {"A", "Bb", "Cb", "C", "Db", "D", "Eb", "Fb", "F", "Gb", "G", "Ab"}

            ElseIf NomeDoMenu = "Gb - Ebm (Bb - Eb - Ab - Db - Gb - Cb)" Then

                NoteNames = {"A", "Bb", "Cb", "C", "Db", "D", "Eb", "E", "F", "Gb", "G", "Ab"}

            ElseIf NomeDoMenu = "Db - Bbm (Bb - Eb - Ab - Db - Gb)" OrElse _
               NomeDoMenu = "Ab - Fm (Bb - Eb - Ab - Db)" OrElse _
               NomeDoMenu = "Eb - Cm (Bb - Eb - Ab)" OrElse _
               NomeDoMenu = "Bb - Gm (Bb - Eb)" OrElse _
               NomeDoMenu = "F - Dm (Bb)" Then

                NoteNames = {"A", "Bb", "B", "C", "Db", "D", "Eb", "E", "F", "Gb", "G", "Ab"}

            End If

            noteName = NoteNames((ToneIndexOffsetToPositives + toneIndex) Mod NoteNames.Length)
            closestFrequency = Math.Pow(ToneStep, toneIndex) * AFrequency

            ToneIndex2 = toneIndex
            toneIndex = (toneIndex + (12 * 5)) - 2
            Octave = 0
            Do While toneIndex > 12
                toneIndex -= 12
                Octave += 1
            Loop

            NomeTecla(0) = "P" & ToneIndex2 + 49 'a variável NomeTecla irá armazenar tanto o que for inserido via teclado, quanto via microfone
            Dim NotaCantadaID As Integer = ToneIndex2 + 49


            CorTecla3.Left = -100
            CorTecla4.Left = -100
            If NotaCantadaID = 2 OrElse NotaCantadaID = 5 OrElse NotaCantadaID = 7 OrElse NotaCantadaID = 10 OrElse NotaCantadaID = 12 OrElse NotaCantadaID = 14 OrElse NotaCantadaID = 17 OrElse NotaCantadaID = 19 OrElse NotaCantadaID = 22 OrElse NotaCantadaID = 24 OrElse NotaCantadaID = 26 OrElse NotaCantadaID = 29 OrElse NotaCantadaID = 31 OrElse NotaCantadaID = 34 OrElse NotaCantadaID = 36 OrElse NotaCantadaID = 38 OrElse NotaCantadaID = 41 OrElse NotaCantadaID = 43 OrElse NotaCantadaID = 46 OrElse NotaCantadaID = 48 OrElse NotaCantadaID = 50 OrElse NotaCantadaID = 53 OrElse NotaCantadaID = 55 OrElse NotaCantadaID = 58 OrElse NotaCantadaID = 60 OrElse NotaCantadaID = 62 OrElse NotaCantadaID = 65 OrElse NotaCantadaID = 67 OrElse NotaCantadaID = 70 OrElse NotaCantadaID = 72 OrElse NotaCantadaID = 74 OrElse NotaCantadaID = 77 OrElse NotaCantadaID = 79 OrElse NotaCantadaID = 82 OrElse NotaCantadaID = 84 OrElse NotaCantadaID = 86 Then
                CorTecla4.Location = PosiçãoTecla(NotaCantadaID - 1, 0)
            Else
                CorTecla3.Location = PosiçãoTecla(NotaCantadaID - 1, 0)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Public device As SoundCaptureDevice = Nothing
    Private Sub listenButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles listenButton.Click

        Try

            Using form As New SelectDeviceForm()
                If form.ShowDialog() = DialogResult.OK Then
                    device = form.SelectedDevice
                Else
                    device = Nothing
                End If
            End Using

            If device IsNot Nothing Then
                StartListenning(device)
                UpdateListenStopButtons()

                CorTecla3.Visible = True
                CorTecla4.Visible = True

                IndicadorDeFrequencia.Show()
                Timer6.Enabled = True
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub stopButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles stopButton.Click
        Try

            StopListenning()
            UpdateListenStopButtons()


            CorTecla3.Left = -100
            CorTecla4.Left = -100
            CorTecla3.Visible = False
            CorTecla4.Visible = False


            GerarImagemDeFundo()
            GerarImagemDeFundo2()


            IndicadorDeFrequencia.Close()
            Timer6.Enabled = False

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Public Sub UpdateListenStopButtons()

        Try

            listenButton.Enabled = Not m_isListenning
            stopButton.Enabled = m_isListenning

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub Timer4_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer4.Tick
        Try

            Timer4.Enabled = False

            If RadioButton2.Checked AndAlso NomeTecla(CInt(TextBox1.Text) - 1) <> "" Then
                NomeTeclaFinal = NomeTecla(0) & NomeTecla(1) & NomeTecla(2) & NomeTecla(3)
                DefineValorCorretoIncorreto()

                If (My.Settings.NovoValorFonteCaptaçãoÁudio(0) = "True") OrElse (My.Settings.NovoValorFonteCaptaçãoÁudio(1) = "True" AndAlso w = 1) Then
                    'só executa este código se Captação Áudio for VOZ, ou se for VIOLÃO desde que a resposta esteja correta (w = 1 é resposta correta, w = 2 é resposta incorreta)
                    'no modo captação áudio VIOLÃO o sistema não indicará respostas incorretas, e somente passará para o próximo exercício quando acertar a nota,
                    'já no modo VOZ o sistema indicará as notas corretas e incorretas, e irá para o próximo exercício tanto quando acertar ou quando errar a nota
                    MostraNotaCorreta()
                    ExibePontuação()
                End If

                NomeTecla(0) = "" : NomeTecla(1) = "" : NomeTecla(2) = "" : NomeTecla(3) = ""

                If RadioButton2.Checked AndAlso NomeTecla(0) = "" Then
                    If Not CheckBox26.Checked Then
                        If (My.Settings.NovoValorFonteCaptaçãoÁudio(0) = "True") OrElse (My.Settings.NovoValorFonteCaptaçãoÁudio(1) = "True" AndAlso w = 1) Then
                            'só executa este código se Captação Áudio for VOZ, ou se for VIOLÃO desde que a resposta esteja correta (w = 1 é resposta correta, w = 2 é resposta incorreta)
                            'no modo captação áudio VIOLÃO o sistema não indicará respostas incorretas, e somente passará para o próximo exercício quando acertar a nota,
                            'já no modo VOZ o sistema indicará as notas corretas e incorretas, e irá para o próximo exercício tanto quando acertar ou quando errar a nota
                            ExecutaLoopNotasAleatórias()
                        End If
                    Else
                        Timer1.Enabled = True
                    End If

                    Timer2.Enabled = False
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub PictureBox4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox4.Click
        ReligarCaptaçãoSom()
        IndicadorDeFrequencia.BringToFront()
    End Sub

    Public Sub ReligarCaptaçãoSom()
        Try
            If stopButton.Enabled Then
                If device IsNot Nothing Then
                    StopListenning()
                    StartListenning(device)
                    UpdateListenStopButtons()
                End If

            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub TrackBar3_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar3.Scroll
        IndicadorDeFrequencia.Invalidate()
    End Sub


    Dim contadorReativaCaptacao As Integer = 0
    Private Sub Timer6_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer6.Tick
        Try
            contadorReativaCaptacao += 1

            If contadorReativaCaptacao = 3 Then 'se o contador atingir 3 (passou-se 3 segundos), entao reativa captacao do som
                ReligarCaptaçãoSom()
                contadorReativaCaptacao = 0
                'Timer6.Enabled = False
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Voz_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Voz.CheckedChanged
        Try
            If Voz.Checked Then
                For i = 0 To 4
                    My.Settings.NovoValorTransposição(i) = "False"
                Next
                My.Settings.NovoValorMaxFreq = 3951.1 'B7 última nota Si do teclado, acima deste valor está dando erro index out of bounds
                CheckBox4.Enabled = True
            End If
            My.Settings.NovoValorFonteCaptaçãoÁudio(0) = CStr(Voz.Checked)
            LineInputChanged()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Violão_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Violão.CheckedChanged
        Try
            If Violão.Checked Then

                For i = 0 To 4
                    My.Settings.NovoValorTransposição(i) = "False"
                Next
                My.Settings.NovoValorTransposição(1) = "True"
                My.Settings.NovoValorMaxFreq = 1174.66 'D6 

                If CheckBox4.Checked Then CheckBox3.Checked = True
                CheckBox4.Checked = False
                CheckBox4.Enabled = False

            End If
            My.Settings.NovoValorFonteCaptaçãoÁudio(1) = CStr(Violão.Checked)
            LineInputChanged()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub LineInputVoz_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LineInputVoz1.Click, LineInputVoz2.Click
        If LineInputVoz1.Checked Then
            My.Settings.NovoValorLineInput(0) = "1"
        Else
            My.Settings.NovoValorLineInput(0) = "2"
        End If
        If Voz.Checked Then LineInputChanged()
    End Sub

    Private Sub LineInputViolão_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LineInputViolão1.Click, LineInputViolão2.Click
        If LineInputViolão1.Checked Then
            My.Settings.NovoValorLineInput(1) = "1"
        Else
            My.Settings.NovoValorLineInput(1) = "2"
        End If
        If Violão.Checked Then LineInputChanged()
    End Sub

    Private Sub LineInputChanged()

        Try

            If Voz.Checked Then
                My.Settings.NovoValorLineInput(2) = My.Settings.NovoValorLineInput(0)
            Else
                My.Settings.NovoValorLineInput(2) = My.Settings.NovoValorLineInput(1)
            End If
            ReligarCaptaçãoSom()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub


    '====================================================================================
    '====================================================================================
    '====================================================================================
    '====================================================================================
    '====================================================================================
    '====================================================================================

End Class