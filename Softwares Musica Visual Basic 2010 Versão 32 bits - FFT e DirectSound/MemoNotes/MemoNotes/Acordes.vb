Option Strict On
Option Explicit On
Imports System.Drawing.Bitmap
Imports System.Drawing.Imaging

Public Class Acordes
    Inherits PerPixelAlphaForm 'para mecher no form no modo design é preciso mudar para 'Inherits Form'
    'e depois de ajustado voltar para 'Inherits PerPixelAlphaForm' e clicar para mudar o inherit
    Public TransAmount As Byte = 255

    Private lista As New List(Of Keys)
    Dim FamiliaAcorde, TonalidadeAcorde, BaixoDoAcordeTerça, BaixoDoAcordeTerçaMenor, BaixoDoAcordeQuarta, BaixoDoAcordeQuinta As String
    Dim BaixoDoAcordeQuintaDiminuida, BaixoDoAcordeQuintaAumentada, BaixoDoAcordeSexta, BaixoDoAcordeSétima, BaixoDoAcordeSétimaMaior As String
    Dim Fonte As New Font("Arial", 12, FontStyle.Italic, GraphicsUnit.Pixel)
    Dim Fonte2 As New Font("Arial", 8, FontStyle.Regular, GraphicsUnit.Pixel)
    Dim Fonte3 As New Font("Times New Roman", 9, FontStyle.Regular, GraphicsUnit.Pixel)
    Dim Fonte4 As New Font("Times New Roman", 20, FontStyle.Bold, GraphicsUnit.Pixel)
    Dim Fonte5 As New Font("Arial", 11, FontStyle.Italic, GraphicsUnit.Pixel)
    Dim Fonte6 As New Font("Times New Roman", 11, FontStyle.Regular, GraphicsUnit.Pixel)
    Dim CorFonte As SolidBrush
    Dim CorFonte2 As SolidBrush = New SolidBrush(Color.FromArgb(0, 0, 0))
    Dim CorFonte3 As SolidBrush = New SolidBrush(Color.FromArgb(34, 177, 76))
    Dim CorFonte4 As SolidBrush
    Dim CorFonte5 As SolidBrush
    Dim CorFonte6 As SolidBrush
    Dim Cor As SolidBrush
    Dim CorBolinha As SolidBrush
    Dim CorPestana As Pen
    Dim Seta As Pen
    Dim Rect1 As Pen = New Pen(Color.FromArgb(200, 255, 0, 0), 2)
    Dim Rect2 As Pen = New Pen(Color.FromArgb(0, 0, 0), 4)
    Dim PosiçãoNotaFundamental As Pen
    Dim LinhaBracinhoViolão1A As Pen
    Dim LinhaBracinhoViolão1B As Pen
    Dim LinhaBracinhoViolão2 As Pen
    Dim LinhaDivisória As Pen = New Pen(Color.FromArgb(0, 0, 0), 2)
    Dim Linha As Pen
    Dim CorNotasBraçoViolão As SolidBrush = New SolidBrush(Color.Transparent)

    Dim AjusteLeft(6), AjusteTopo, AjustePosição(48), Index, FatorEscala, LarguraImagemCopiada, d, f, Pestana, LinhaFinalDaSeta(48), NumeraçãoTrastes, _
        AjusteLeftIntervalos, ValorEsquerda, ValorTopo, AjustePosiçãoTonalidadeDoAcorde, NotasAcordeIndiceLinha, _
        NotasAcordeIndiceColuna, AcordeAleatório, TonicaLinha, TonicaColuna As Integer

    Dim ArrayAcordePadrão(48, 8, 6), ImagemAcorde(48), IntervalosAcorde(27, 6), StringSubstituição(1) As String

    Dim ImagemCopiada(2) As Bitmap

    Dim modoJogo As Boolean

    Dim IntervalosAcorde2(,) As String = {{"", "3", "6", "9", "5", "7M", "3"}, _
                                        {"", "4", "7", "b3", "b6", "T", "4"}, _
                                        {"", "b5", "7M", "3", "6", "b9", "b5"}, _
                                        {"", "5", "T", "4", "7", "9", "5"}, _
                                        {"", "b6", "b9", "b5", "7M", "b3", "b6"}, _
                                        {"", "6", "9", "5", "T", "3", "6"}, _
                                        {"", "7", "b3", "b6", "b9", "4", "7"}, _
                                        {"", "7M", "3", "6", "9", "b5", "7M"}, _
                                        {"", "T", "4", "7", "b3", "5", "T"}, _
                                        {"", "b9", "b5", "7M", "3", "b6", "b9"}, _
                                        {"", "9", "5", "T", "4", "6", "9"}, _
                                        {"", "b3", "b6", "b9", "b5", "7", "b3"}, _
                                        {"", "3", "6", "9", "5", "7M", "3"}, _
                                        {"", "4", "7", "b3", "b6", "T", "4"}, _
                                        {"", "b5", "7M", "3", "6", "b9", "b5"}, _
                                        {"", "5", "T", "4", "7", "9", "5"}, _
                                        {"", "b6", "b9", "b5", "7M", "b3", "b6"}, _
                                        {"", "6", "9", "5", "T", "3", "6"}, _
                                        {"", "7", "b3", "b6", "b9", "4", "7"}, _
                                        {"", "7M", "3", "6", "9", "b5", "7M"}, _
                                        {"", "T", "4", "7", "b3", "5", "T"}, _
                                        {"", "b9", "b5", "7M", "3", "b6", "b9"}, _
                                        {"", "9", "5", "T", "4", "6", "9"}, _
                                        {"", "b3", "b6", "b9", "b5", "7", "b3"}, _
                                        {"", "3", "6", "9", "5", "7M", "3"}, _
                                        {"", "4", "7", "b3", "b6", "T", "4"}, _
                                        {"", "b5", "7M", "3", "6", "b9", "b5"}, _
                                        {"", "5", "T", "4", "7", "9", "5"}}

    Dim NotasAcorde(,) As String = {{"", "C", "Db", "D", "D#", "Eb", "E", "F", "F#", "Gb", "G", "G#", "Ab", "A", "Bb", "B"}, _
                                    {"", "D", "Eb", "E", "E#", "F", "F#", "G", "G#", "Ab", "A", "A#", "Bb", "B", "C", "C#"}, _
                                    {"", "E", "F", "F#", "G", "G", "G#", "A", "A#", "Bb", "B", "B#", "C", "C#", "D", "D#"}, _
                                    {"", "F", "Gb", "G", "G#", "Ab", "A", "Bb", "B", "Cb", "C", "C#", "Db", "D", "Eb", "E"}, _
                                    {"", "G", "Ab", "A", "A#", "Bb", "B", "C", "C#", "Db", "D", "D#", "Eb", "E", "F", "F#"}, _
                                    {"", "A", "Bb", "B", "B#", "C", "C#", "D", "D#", "Eb", "E", "E#", "F", "F#", "G", "G#"}, _
                                    {"", "B", "C", "C#", "D", "D", "D#", "E", "E#", "F", "F#", "G", "G", "G#", "A", "A#"}}

    Dim NotasBraçoViolão(,) As String = {{"", "E", "A", "D", "G", "B", "E"}, _
                                        {"", "F", "", "", "", "C", "F"}, _
                                        {"", "", "B", "E", "A", "", ""}, _
                                        {"", "G", "C", "F", "", "D", "G"}, _
                                        {"", "", "", "", "B", "", ""}, _
                                        {"", "A", "D", "G", "C", "E", "A"}, _
                                        {"", "", "", "", "", "F", ""}, _
                                        {"", "B", "E", "A", "D", "", "B"}, _
                                        {"", "C", "F", "", "", "G", "C"}, _
                                        {"", "", "", "B", "E", "", ""}, _
                                        {"", "D", "G", "C", "F", "A", "D"}, _
                                        {"", "", "", "", "", "", ""}, _
                                        {"", "E", "A", "D", "G", "B", "E"}, _
                                        {"", "F", "", "", "", "C", "F"}, _
                                        {"", "", "B", "E", "A", "", ""}, _
                                        {"", "G", "C", "F", "", "D", "G"}, _
                                        {"", "", "", "", "B", "", ""}, _
                                        {"", "A", "D", "G", "C", "E", "A"}, _
                                        {"", "", "", "", "", "F", ""}, _
                                        {"", "B", "E", "A", "D", "", "B"}, _
                                        {"", "C", "F", "", "", "G", "C"}, _
                                        {"", "", "", "B", "E", "", ""}, _
                                        {"", "D", "G", "C", "F", "A", "D"}}


    Private Sub CenterTextAt(ByVal gr As Graphics, ByVal txt As _
    String, ByVal x As Single, ByVal y As Single)
        ' Mark the center for debugging.
        'gr.DrawLine(Pens.Red, x - 10, y, x + 10, y)
        'gr.DrawLine(Pens.Red, x, y - 10, x, y + 10)

        ' Make a StringFormat object that centers.
        Dim sf As New StringFormat
        sf.LineAlignment = StringAlignment.Center
        sf.Alignment = StringAlignment.Center

        ' Draw the text.
        gr.DrawString(txt, Fonte6, CorFonte6, x, y, sf)
        sf.Dispose()
    End Sub

    Private Sub CenterTextAt_2(ByVal gr As Graphics, ByVal txt As _
    String, ByVal x As Single, ByVal y As Single)
        ' Mark the center for debugging.
        'gr.DrawLine(Pens.Red, x - 10, y, x + 10, y)
        'gr.DrawLine(Pens.Red, x, y - 10, x, y + 10)

        ' Make a StringFormat object that centers.
        Dim sf As New StringFormat
        sf.LineAlignment = StringAlignment.Center
        sf.Alignment = StringAlignment.Center

        ' Draw the text.
        gr.DrawString(txt, Fonte3, CorFonte6, x, y, sf)
        sf.Dispose()
    End Sub

    Public Sub Desenhar()

        Try

            CorFonte5 = New SolidBrush(My.Settings.NovoValorNumeraçãoTrastes) 'New SolidBrush(Color.FromArgb(150, 150, 150))
            LinhaBracinhoViolão2 = New Pen(My.Settings.NovoValorLinhasDiagramaAcordes, 1) 'New Pen(Color.FromArgb(200, 200, 200), 1)
            LinhaBracinhoViolão1A = New Pen(My.Settings.NovoValorLinhasDiagramaAcordes, 3) 'New Pen(Color.FromArgb(200, 200, 200), 3)
            LinhaBracinhoViolão1B = New Pen(My.Settings.NovoValorAcordesMaisUsados, 3) 'New Pen(Color.FromArgb(153, 217, 234), 3)
            CorBolinha = New SolidBrush(My.Settings.NovoValorBolinhaAcordes) 'New SolidBrush(Color.FromArgb(0, 0, 0))
            CorFonte = New SolidBrush(My.Settings.NovoValorNumeraçãoDedilhados1) 'New SolidBrush(Color.FromArgb(0, 0, 0))
            CorFonte4 = New SolidBrush(My.Settings.NovoValorNumeraçãoDedilhados2) 'New SolidBrush(Color.FromArgb(255, 255, 255))
            PosiçãoNotaFundamental = New Pen(My.Settings.NovoValorNotaDeReferênciaAcorde, 1) 'New Pen(Color.FromArgb(255, 0, 0), 1)
            CorPestana = New Pen(My.Settings.NovoValorCorPestanas, 3) 'New Pen(Color.FromArgb(0, 0, 0), 3)
            Seta = New Pen(My.Settings.NovoValorCorPestanas, 5) 'New Pen(Color.FromArgb(0, 0, 0), 5)
            CorFonte6 = New SolidBrush(My.Settings.NovoValorCorDasCifras) 'New SolidBrush(Color.FromArgb(0, 0, 0))

            ' Make a StringFormat object that centers.
            Dim sf As New StringFormat
            sf.LineAlignment = StringAlignment.Center
            sf.Alignment = StringAlignment.Center



            If modoJogo = True Then
                AcordeAleatório = 0
                Do While AcordeAleatório = 0
                    ' gera o array de bytes randômico de 4 bytes...
                    Dim randomNumber(3) As Byte

                    ' Create a new instance of the RNGCryptoServiceProvider.
                    Dim Gen As New Security.Cryptography.RNGCryptoServiceProvider()

                    ' Fill the array with a random value.
                    Gen.GetBytes(randomNumber)

                    ' calcula o número baseado no valor máximo
                    AcordeAleatório = Math.Abs(BitConverter.ToInt32(randomNumber, 0)) Mod (Index)

                    'não exibirá acordes invertidos se a opção "Exibir acordes invertidos" estiver desmarcada
                    If My.Settings.NovoValorExibirAcordesInvertidos = False AndAlso ArrayAcordePadrão(AcordeAleatório, 8, 4) = "x" Then AcordeAleatório = 0
                    If My.Settings.NovoValorExibirAcordesMaisUsadosModoJogo = True AndAlso ArrayAcordePadrão(AcordeAleatório, 8, 1) <> "*" Then AcordeAleatório = 0

                Loop
            End If



            Dim Cont As Integer
            If OpçõesTelaAcordes.Visible = False AndAlso TelaAcordesModoJogo.Visible = False Then 'não copiará imagens enquanto o formulário Opções estiver aberto
                Cont = 2
            Else
                Cont = 1
            End If

            For i = 1 To Cont
                Dim gr As Graphics
                Dim FaceBit As Bitmap


                If i = 1 Then
                    FaceBit = New Bitmap(My.Resources.Acordes)
                    gr = Graphics.FromImage(FaceBit)
                    gr.TextRenderingHint = TextRenderingHint.ClearTypeGridFit
                    If modoJogo = False Then CenterTextAt(gr, FamiliaAcorde, CInt(Me.Width / 2), 160)
                Else
                    FatorEscala = My.Settings.NovoValorFatorEscala
                    FaceBit = New Bitmap(My.Resources.Acordes, My.Resources.Acordes.Width * FatorEscala, My.Resources.Acordes.Height * FatorEscala)
                    gr = Graphics.FromImage(FaceBit)
                    gr.ScaleTransform(FatorEscala, FatorEscala)
                    gr.FillRectangle(Brushes.White, 0, 0, My.Resources.Acordes.Width * FatorEscala, My.Resources.Acordes.Height * FatorEscala)
                    gr.TextRenderingHint = TextRenderingHint.AntiAlias
                End If

                gr.FillRectangle(New SolidBrush(My.Settings.NovaCorIntervalos), CorIntervalos.Left + 2, CorIntervalos.Top + 2, 8, 8)


                If My.Settings.NovoValorDesenharBolinhasNotasAcordes = True Then gr.DrawImage(My.Resources.CorretoRitmo, Label3.Left + 2, Label3.Top - 1, 11, 11)
                If My.Settings.NovoValorDesenharAcordesMaisUsados = True Then gr.DrawImage(My.Resources.CorretoRitmo, Label4.Left + 2, Label4.Top - 1, 11, 11)
                If My.Settings.NovoValorLocalizaçãoFundamentalDeReferenciaNoAcorde = True Then gr.DrawImage(My.Resources.CorretoRitmo, Label5.Left + 2, Label5.Top - 1, 11, 11)
                If My.Settings.NovoValorExibirNomeIntervalos = True Then gr.DrawImage(My.Resources.CorretoRitmo, Label8.Left + 2, Label8.Top - 1, 11, 11)
                If My.Settings.NovoValorCopiarImagemDosAcordes = True Then gr.DrawImage(My.Resources.CorretoRitmo, Label9.Left + 2, Label9.Top - 1, 11, 11)
                If My.Settings.NovoValorExibirNumerosDedilhado = True Then gr.DrawImage(My.Resources.CorretoRitmo, Label17.Left + 2, Label17.Top - 1, 11, 11)
                If My.Settings.NovoValorExibirNomeDasNotasAcordes = True Then gr.DrawImage(My.Resources.CorretoRitmo, Label18.Left + 2, Label18.Top - 1, 11, 11)
                If My.Settings.NovoValorSalvarAcordes = True Then gr.DrawImage(My.Resources.CorretoRitmo, Label19.Left + 2, Label19.Top - 1, 11, 11)

                Seta.StartCap = LineCap.ArrowAnchor
                PosiçãoNotaFundamental.DashStyle = DashStyle.Dash


                If FamiliaAcorde = "Tríades maiores – 4 cordas com inversões" OrElse NumeraçãoFamiliaAcorde = 62 OrElse _
                    FamiliaAcorde = "Tríades menores – 4 cordas com inversões" OrElse NumeraçãoFamiliaAcorde = 63 OrElse _
                    FamiliaAcorde = "Tríades aumentadas – 4 cordas com inversões" OrElse NumeraçãoFamiliaAcorde = 64 OrElse _
                    FamiliaAcorde = "Tríades com quarta – 4 cordas com inversões" OrElse NumeraçãoFamiliaAcorde = 65 Then

                    gr.DrawLine(LinhaDivisória, 283, 172, 283, 519)
                    gr.DrawLine(LinhaDivisória, 538, 172, 538, 519)

                End If

                gr.SmoothingMode = SmoothingMode.AntiAlias

                If My.Settings.NovoValorExibirIntervalosBraçoViolão = True Then
                    gr.DrawEllipse(New Pen(Brushes.Red, 2), Label24.Left, Label24.Top, 12, 12)
                ElseIf My.Settings.NovoValorExibirNotasBraçoViolão = True Then
                    gr.DrawEllipse(New Pen(Brushes.Red, 2), Label21.Left, Label21.Top, 12, 12)
                End If

                LarguraImagemCopiada = 0

                For a = 1 To 48 'percorre os 48 arrays

                    If (modoJogo = True AndAlso a = AcordeAleatório) OrElse modoJogo = False Then 'se estiver no modo jogo desenha somente o acorde aleatório, caso contrário desenha todos os acordes

                        If ArrayAcordePadrão(a, 0, 0) = "Fim" Then
                            LarguraImagemCopiada = ((a - 1) * 128) - 32
                            If LarguraImagemCopiada > 992 Then LarguraImagemCopiada = 992
                            a = 100 'sairá do loop caso o array esteja vazio
                        ElseIf ImagemAcorde(a) IsNot Nothing AndAlso ImagemAcorde(a) IsNot " " Then


                            Pestana = 0
                            If LinhaFinalDaSeta(a) < 6 Then
                                f = 35
                            Else
                                f = 31
                            End If



                            If ArrayAcordePadrão(a, 1, 0) = "Traste3" Then
                                AjustePosição(0) = 2
                            ElseIf ArrayAcordePadrão(a, 1, 0) = "Traste4" Then
                                AjustePosição(0) = 3
                            ElseIf ArrayAcordePadrão(a, 1, 0) = "Traste5" Then
                                AjustePosição(0) = 4
                            ElseIf ArrayAcordePadrão(a, 1, 0) = "Traste6" Then
                                AjustePosição(0) = 5
                            ElseIf ArrayAcordePadrão(a, 1, 0) = "Traste7" Then
                                AjustePosição(0) = 6
                            ElseIf ArrayAcordePadrão(a, 1, 0) = "Traste8" Then
                                AjustePosição(0) = 7
                            ElseIf ArrayAcordePadrão(a, 1, 0) = "Traste9" Then
                                AjustePosição(0) = 8
                            ElseIf ArrayAcordePadrão(a, 1, 0) = "Traste10" Then
                                AjustePosição(0) = 9
                            ElseIf ArrayAcordePadrão(a, 1, 0) = "Traste11" Then
                                AjustePosição(0) = 10
                            ElseIf ArrayAcordePadrão(a, 1, 0) = "Traste12" Then
                                AjustePosição(0) = 11
                            Else
                                AjustePosição(0) = 0
                            End If


                            AjustePosição(0) += AjustePosição(a)

                            If a >= 1 AndAlso a <= 8 Then
                                AjusteTopo = 0
                            ElseIf a >= 9 AndAlso a <= 16 Then
                                AjusteTopo = 118
                            ElseIf a >= 17 AndAlso a <= 24 Then
                                AjusteTopo = 236
                            ElseIf a >= 25 AndAlso a <= 32 Then
                                AjusteTopo = 354
                            ElseIf a >= 33 AndAlso a <= 40 Then
                                AjusteTopo = 472
                            Else
                                AjusteTopo = 590
                            End If


                            If a = 1 OrElse a = 9 OrElse a = 17 OrElse a = 25 OrElse a = 33 OrElse a = 41 Then
                                AjusteLeft(0) = 0 : AjusteLeft(1) = 5 : AjusteLeft(2) = 2 : AjusteLeft(3) = 8 : AjusteLeft(4) = 10
                            ElseIf a = 2 OrElse a = 10 OrElse a = 18 OrElse a = 26 OrElse a = 34 OrElse a = 42 Then
                                AjusteLeft(0) = 128 : AjusteLeft(1) = 133 : AjusteLeft(2) = 130 : AjusteLeft(3) = 136 : AjusteLeft(4) = 138
                            ElseIf a = 3 OrElse a = 11 OrElse a = 19 OrElse a = 27 OrElse a = 35 OrElse a = 43 Then
                                AjusteLeft(0) = 256 : AjusteLeft(1) = 261 : AjusteLeft(2) = 258 : AjusteLeft(3) = 264 : AjusteLeft(4) = 266
                            ElseIf a = 4 OrElse a = 12 OrElse a = 20 OrElse a = 28 OrElse a = 36 OrElse a = 44 Then
                                AjusteLeft(0) = 384 : AjusteLeft(1) = 389 : AjusteLeft(2) = 386 : AjusteLeft(3) = 392 : AjusteLeft(4) = 394
                            ElseIf a = 5 OrElse a = 13 OrElse a = 21 OrElse a = 29 OrElse a = 37 OrElse a = 45 Then
                                AjusteLeft(0) = 512 : AjusteLeft(1) = 517 : AjusteLeft(2) = 514 : AjusteLeft(3) = 520 : AjusteLeft(4) = 522
                            ElseIf a = 6 OrElse a = 14 OrElse a = 22 OrElse a = 30 OrElse a = 38 OrElse a = 46 Then
                                AjusteLeft(0) = 640 : AjusteLeft(1) = 645 : AjusteLeft(2) = 642 : AjusteLeft(3) = 648 : AjusteLeft(4) = 650
                            ElseIf a = 7 OrElse a = 15 OrElse a = 23 OrElse a = 31 OrElse a = 39 OrElse a = 47 Then
                                AjusteLeft(0) = 768 : AjusteLeft(1) = 773 : AjusteLeft(2) = 770 : AjusteLeft(3) = 776 : AjusteLeft(4) = 778
                            ElseIf a = 8 OrElse a = 16 OrElse a = 24 OrElse a = 32 OrElse a = 40 OrElse a = 48 Then
                                AjusteLeft(0) = 896 : AjusteLeft(1) = 901 : AjusteLeft(2) = 898 : AjusteLeft(3) = 904 : AjusteLeft(4) = 906
                            End If


                            Linha = LinhaBracinhoViolão1A
                            If My.Settings.NovoValorDesenharAcordesMaisUsados = True AndAlso ArrayAcordePadrão(a, 8, 1) = "*" Then
                                Linha = LinhaBracinhoViolão1B
                            End If


                            gr.SmoothingMode = SmoothingMode.None

                            If TelaAcordesModoJogo.Visible = True AndAlso a = AcordeAleatório Then
                                If TelaAcordesModoJogo.Label1.Text = "Correto" Then
                                    gr.FillRectangle(Brushes.LimeGreen, 53 + AjusteLeft(0), 168 + AjusteTopo, 76, 12)
                                ElseIf TelaAcordesModoJogo.Label1.Text = "Incorreto" Then
                                    gr.FillRectangle(Brushes.OrangeRed, 53 + AjusteLeft(0), 168 + AjusteTopo, 76, 12)
                                End If
                            End If

                            'desenha os bracinhos do violão
                            'linhas horizontais
                            gr.DrawLine(LinhaBracinhoViolão2, 53 + AjusteLeft(0), 195 + AjusteTopo, 128 + AjusteLeft(0), 195 + AjusteTopo)
                            gr.DrawLine(LinhaBracinhoViolão2, 53 + AjusteLeft(0), 210 + AjusteTopo, 128 + AjusteLeft(0), 210 + AjusteTopo)
                            gr.DrawLine(LinhaBracinhoViolão2, 53 + AjusteLeft(0), 225 + AjusteTopo, 128 + AjusteLeft(0), 225 + AjusteTopo)
                            gr.DrawLine(LinhaBracinhoViolão2, 53 + AjusteLeft(0), 240 + AjusteTopo, 128 + AjusteLeft(0), 240 + AjusteTopo)
                            gr.DrawLine(LinhaBracinhoViolão2, 53 + AjusteLeft(0), 255 + AjusteTopo, 128 + AjusteLeft(0), 255 + AjusteTopo)
                            gr.DrawLine(LinhaBracinhoViolão2, 53 + AjusteLeft(0), 270 + AjusteTopo, 128 + AjusteLeft(0), 270 + AjusteTopo)
                            'linhas verticais
                            gr.DrawLine(LinhaBracinhoViolão2, 53 + AjusteLeft(0), 179 + AjusteTopo, 53 + AjusteLeft(0), 270 + AjusteTopo)
                            gr.DrawLine(LinhaBracinhoViolão2, 68 + AjusteLeft(0), 179 + AjusteTopo, 68 + AjusteLeft(0), 270 + AjusteTopo)
                            gr.DrawLine(LinhaBracinhoViolão2, 83 + AjusteLeft(0), 179 + AjusteTopo, 83 + AjusteLeft(0), 270 + AjusteTopo)
                            gr.DrawLine(LinhaBracinhoViolão2, 98 + AjusteLeft(0), 179 + AjusteTopo, 98 + AjusteLeft(0), 270 + AjusteTopo)
                            gr.DrawLine(LinhaBracinhoViolão2, 113 + AjusteLeft(0), 179 + AjusteTopo, 113 + AjusteLeft(0), 270 + AjusteTopo)
                            gr.DrawLine(LinhaBracinhoViolão2, 128 + AjusteLeft(0), 179 + AjusteTopo, 128 + AjusteLeft(0), 270 + AjusteTopo)


                            gr.DrawLine(Linha, 53 + AjusteLeft(0), 180 + AjusteTopo, 128 + AjusteLeft(0), 180 + AjusteTopo) 'esta linha horizontal precisa ser desenhada por último

                            If ImagemAcorde(a) IsNot Nothing Then
                                If modoJogo = False OrElse (modoJogo = True AndAlso My.Settings.NovoValorTrocarAcordesNoIntervaloDeTempo = True) Then
                                    CenterTextAt(gr, ImagemAcorde(a), 91 + AjusteLeft(0), 174 + AjusteTopo)
                                    If ArrayAcordePadrão(a, 8, 5) IsNot Nothing Then
                                        gr.FillRectangle(New SolidBrush(Color.FromArgb(125, 255, 255, 255)), 57 + AjusteLeft(0), 251 + AjusteTopo, 68, 17)
                                        CenterTextAt_2(gr, ArrayAcordePadrão(a, 8, 5), 91 + AjusteLeft(0), 260 + AjusteTopo)
                                    End If
                                End If
                            End If

                            gr.SmoothingMode = SmoothingMode.AntiAlias

                            AjusteLeft(6) = 0
                            d = 0
                            For b = 0 To 6  'percorre as 6 colunas de cada array
                                For c = 0 To 8 'percorre as 8 linhas do array
                                    If ArrayAcordePadrão(a, c, b) <> "" Then

                                        If c < 7 Then
                                            If ArrayAcordePadrão(a, c, b) = "-1" OrElse ArrayAcordePadrão(a, c, b) = "0" OrElse ArrayAcordePadrão(a, c, b) = "1" OrElse ArrayAcordePadrão(a, c, b) = "2" OrElse ArrayAcordePadrão(a, c, b) = "3" OrElse ArrayAcordePadrão(a, c, b) = "4" OrElse ArrayAcordePadrão(a, c, b) = "5" OrElse _
                                               ArrayAcordePadrão(a, c, b) = "1T" OrElse ArrayAcordePadrão(a, c, b) = "2T" OrElse ArrayAcordePadrão(a, c, b) = "3T" OrElse ArrayAcordePadrão(a, c, b) = "4T" OrElse ArrayAcordePadrão(a, c, b) = "5T" OrElse
                                               ArrayAcordePadrão(a, c, b) = "1S" OrElse ArrayAcordePadrão(a, c, b) = "2S" OrElse ArrayAcordePadrão(a, c, b) = "3S" OrElse ArrayAcordePadrão(a, c, b) = "4S" OrElse ArrayAcordePadrão(a, c, b) = "5S" Then

                                                If (CInt(Replace(Replace(ArrayAcordePadrão(a, c, b), "T", ""), "S", "")) + CInt(ArrayAcordePadrão(a, 8, 3))) <> 0 AndAlso ((c * 15) + 166 + AjusteTopo + CInt(ArrayAcordePadrão(a, 8, 2))) > (180 + AjusteTopo) Then
                                                    If My.Settings.NovoValorDesenharBolinhasNotasAcordes = True Then
                                                        If My.Settings.NovoValorExibirNumerosDedilhado = True Then
                                                            gr.FillEllipse(CorBolinha, (b * 15) + 31 + AjusteLeft(0), (c * 15) + 166 + AjusteTopo + CInt(ArrayAcordePadrão(a, 8, 2)), 13, 13)
                                                            gr.DrawString(CStr(CInt(Replace(Replace(ArrayAcordePadrão(a, c, b), "T", ""), "S", "")) + CInt(ArrayAcordePadrão(a, 8, 3))), Fonte5, CorFonte4, (b * 15) + 32 + AjusteLeft(0), (c * 15) + 166 + AjusteTopo + CInt(ArrayAcordePadrão(a, 8, 2)))
                                                        Else
                                                            gr.FillEllipse(CorBolinha, (b * 15) + 32 + AjusteLeft(0), (c * 15) + 167 + AjusteTopo + CInt(ArrayAcordePadrão(a, 8, 2)), 11, 11)
                                                        End If
                                                    Else
                                                        If My.Settings.NovoValorExibirNumerosDedilhado = True Then gr.DrawString(CStr(CInt(Replace(Replace(ArrayAcordePadrão(a, c, b), "T", ""), "S", "")) + CInt(ArrayAcordePadrão(a, 8, 3))), Fonte, CorFonte, (b * 15) + 32 + AjusteLeft(0), (c * 15) + 166 + AjusteTopo + CInt(ArrayAcordePadrão(a, 8, 2)))
                                                    End If
                                                End If

                                                'Desenha círculo indicando a localização da tônica de referência do acorde
                                                If My.Settings.NovoValorLocalizaçãoFundamentalDeReferenciaNoAcorde = True Then
                                                    Dim AjusteTopoInversão As Integer = AjusteTopo
                                                    If c = 0 OrElse ((c * 15) + 166 + AjusteTopo + CInt(ArrayAcordePadrão(a, 8, 2))) < (180 + AjusteTopo) Then AjusteTopo += 7 'ajuste nos casos da nota de referência for uma corda solta

                                                    If ArrayAcordePadrão(a, c, b) = "1T" OrElse ArrayAcordePadrão(a, c, b) = "2T" OrElse ArrayAcordePadrão(a, c, b) = "3T" OrElse ArrayAcordePadrão(a, c, b) = "4T" Then
                                                        gr.DrawEllipse(PosiçãoNotaFundamental, (b * 15) + 31 + AjusteLeft(0), (c * 15) + 166 + AjusteTopo + CInt(ArrayAcordePadrão(a, 8, 2)), 13, 13)
                                                    ElseIf ArrayAcordePadrão(a, c, b) = "1S" OrElse ArrayAcordePadrão(a, c, b) = "2S" OrElse ArrayAcordePadrão(a, c, b) = "3S" OrElse ArrayAcordePadrão(a, c, b) = "4S" Then
                                                        gr.FillEllipse(New SolidBrush(My.Settings.NovoValorNotaDeReferênciaAcorde), (b * 15) + 35 + AjusteLeft(0), (c * 15) + 177 + AjusteTopoInversão + CInt(ArrayAcordePadrão(a, 8, 2)), 6, 6)
                                                        gr.FillEllipse(Brushes.White, (b * 15) + 36 + AjusteLeft(0), (c * 15) + 178 + AjusteTopoInversão + CInt(ArrayAcordePadrão(a, 8, 2)), 4, 4)
                                                    End If

                                                    If c = 0 OrElse ((c * 15) + 166 + AjusteTopo + CInt(ArrayAcordePadrão(a, 8, 2))) < (180 + AjusteTopo) Then AjusteTopo -= 7
                                                End If
                                            ElseIf ArrayAcordePadrão(a, c, b) = "P" OrElse ArrayAcordePadrão(a, c, b) = "PT" OrElse ArrayAcordePadrão(a, c, b) = "T" OrElse ArrayAcordePadrão(a, c, b) = "S" OrElse ArrayAcordePadrão(a, c, b) = "PS" Then
                                                If (ArrayAcordePadrão(a, c, b) <> "S" AndAlso ArrayAcordePadrão(a, c, b) <> "T") AndAlso ((c * 15) + 173 + AjusteTopo + CInt(ArrayAcordePadrão(a, 8, 2))) > (180 + AjusteTopo) Then
                                                    gr.SmoothingMode = SmoothingMode.None
                                                    gr.DrawLine(CorPestana, (b * 15) + 30 + AjusteLeft(1), (c * 15) + 173 + AjusteTopo + CInt(ArrayAcordePadrão(a, 8, 2)), (b * 15) + f + AjusteLeft(3) + ((LinhaFinalDaSeta(a) - b) * 15), (c * 15) + 173 + AjusteTopo + CInt(ArrayAcordePadrão(a, 8, 2)))
                                                    gr.SmoothingMode = SmoothingMode.AntiAlias
                                                    gr.DrawLine(Seta, (b * 15) + 30 + AjusteLeft(2), (c * 15) + 173 + AjusteTopo + CInt(ArrayAcordePadrão(a, 8, 2)), (b * 15) + 30 + AjusteLeft(4), (c * 15) + 173 + AjusteTopo + CInt(ArrayAcordePadrão(a, 8, 2)))
                                                End If
                                                Pestana = c

                                                'Desenha círculo indicando a localização da tônica de referência do acorde
                                                If My.Settings.NovoValorLocalizaçãoFundamentalDeReferenciaNoAcorde = True Then
                                                    If ((c * 15) + 166 + AjusteTopo + CInt(ArrayAcordePadrão(a, 8, 2))) < (180 + AjusteTopo) Then AjusteTopo += 7 'ajuste nos casos da nota de referência for uma corda solta

                                                    If ArrayAcordePadrão(a, c, b) = "PT" OrElse ArrayAcordePadrão(a, c, b) = "T" Then
                                                        gr.DrawEllipse(PosiçãoNotaFundamental, (b * 15) + 31 + AjusteLeft(0), (c * 15) + 166 + AjusteTopo + CInt(ArrayAcordePadrão(a, 8, 2)), 13, 13)
                                                    ElseIf ArrayAcordePadrão(a, c, b) = "PS" OrElse ArrayAcordePadrão(a, c, b) = "S" Then
                                                        gr.FillEllipse(New SolidBrush(My.Settings.NovoValorNotaDeReferênciaAcorde), (b * 15) + 35 + AjusteLeft(0), (c * 15) + 170 + AjusteTopo + CInt(ArrayAcordePadrão(a, 8, 2)), 6, 6)
                                                        gr.FillEllipse(Brushes.White, (b * 15) + 36 + AjusteLeft(0), (c * 15) + 171 + AjusteTopo + CInt(ArrayAcordePadrão(a, 8, 2)), 4, 4)
                                                    End If

                                                    If ((c * 15) + 166 + AjusteTopo + CInt(ArrayAcordePadrão(a, 8, 2))) < (180 + AjusteTopo) Then AjusteTopo -= 7
                                                End If
                                            ElseIf ArrayAcordePadrão(a, c, b) = "Traste1" OrElse ArrayAcordePadrão(a, c, b) = "Traste3" OrElse ArrayAcordePadrão(a, c, b) = "Traste4" OrElse _
                                            ArrayAcordePadrão(a, c, b) = "Traste5" OrElse ArrayAcordePadrão(a, c, b) = "Traste6" OrElse _
                                            ArrayAcordePadrão(a, c, b) = "Traste7" OrElse ArrayAcordePadrão(a, c, b) = "Traste8" OrElse _
                                            ArrayAcordePadrão(a, c, b) = "Traste9" OrElse ArrayAcordePadrão(a, c, b) = "Traste10" OrElse _
                                            ArrayAcordePadrão(a, c, b) = "Traste11" OrElse ArrayAcordePadrão(a, c, b) = "Traste12" Then


                                                If ArrayAcordePadrão(a, 8, 6) = "" Then
                                                    NumeraçãoTrastes = CInt(Replace(ArrayAcordePadrão(a, c, b), "Traste", "")) + AjustePosiçãoTonalidadeDoAcorde
                                                    If ArrayAcordePadrão(a, 8, 2) = "15" Then
                                                        NumeraçãoTrastes -= 1
                                                    ElseIf ArrayAcordePadrão(a, 8, 2) = "-15" Then
                                                        NumeraçãoTrastes += 1
                                                    ElseIf ArrayAcordePadrão(a, 8, 2) = "-45" Then
                                                        NumeraçãoTrastes += 3
                                                    End If
                                                Else
                                                    NumeraçãoTrastes = CInt(ArrayAcordePadrão(a, 8, 6))
                                                End If


                                                If NumeraçãoTrastes >= 12 Then
                                                    NumeraçãoTrastes -= 12
                                                ElseIf NumeraçãoTrastes < 0 Then
                                                    NumeraçãoTrastes += 12
                                                ElseIf NumeraçãoTrastes = 0 Then
                                                    NumeraçãoTrastes += 1
                                                End If


                                                For ContadorTraste = 0 To 5
                                                    AjusteLeft(5) = 0
                                                    If NumeraçãoTrastes + ContadorTraste < 10 Then
                                                        AjusteLeft(5) = 4
                                                    End If
                                                    gr.DrawString(CStr(NumeraçãoTrastes + ContadorTraste), Fonte2, CorFonte5, (b * 15) + 37 + AjusteLeft(1) + AjusteLeft(5) + AjusteLeft(6), (c * 15) + 168 + AjusteTopo + (ContadorTraste * 15))
                                                Next

                                                If My.Settings.NovoValorExibirNotasBraçoViolão = True Then
                                                    For ContadorNotasBraçoViolãoLinha = 0 To 6
                                                        For ContadorNotasBraçoViolãoColuna = 0 To 6
                                                            CorNotasBraçoViolão = New SolidBrush(Color.Transparent)
                                                            If NotasBraçoViolão(NumeraçãoTrastes - 1 + ContadorNotasBraçoViolãoLinha, ContadorNotasBraçoViolãoColuna) = "C" Then
                                                                CorNotasBraçoViolão = New SolidBrush(Color.FromArgb(150, Color.Black))
                                                            ElseIf NotasBraçoViolão(NumeraçãoTrastes - 1 + ContadorNotasBraçoViolãoLinha, ContadorNotasBraçoViolãoColuna) = "D" Then
                                                                CorNotasBraçoViolão = New SolidBrush(Color.FromArgb(150, Color.Green))
                                                            ElseIf NotasBraçoViolão(NumeraçãoTrastes - 1 + ContadorNotasBraçoViolãoLinha, ContadorNotasBraçoViolãoColuna) = "E" Then
                                                                CorNotasBraçoViolão = New SolidBrush(Color.FromArgb(150, Color.SkyBlue))
                                                            ElseIf NotasBraçoViolão(NumeraçãoTrastes - 1 + ContadorNotasBraçoViolãoLinha, ContadorNotasBraçoViolãoColuna) = "F" Then
                                                                CorNotasBraçoViolão = New SolidBrush(Color.FromArgb(150, Color.Red))
                                                            ElseIf NotasBraçoViolão(NumeraçãoTrastes - 1 + ContadorNotasBraçoViolãoLinha, ContadorNotasBraçoViolãoColuna) = "G" Then
                                                                CorNotasBraçoViolão = New SolidBrush(Color.FromArgb(150, Color.Yellow))
                                                            ElseIf NotasBraçoViolão(NumeraçãoTrastes - 1 + ContadorNotasBraçoViolãoLinha, ContadorNotasBraçoViolãoColuna) = "A" Then
                                                                CorNotasBraçoViolão = New SolidBrush(Color.FromArgb(150, Color.Orange))
                                                            ElseIf NotasBraçoViolão(NumeraçãoTrastes - 1 + ContadorNotasBraçoViolãoLinha, ContadorNotasBraçoViolãoColuna) = "B" Then
                                                                CorNotasBraçoViolão = New SolidBrush(Color.FromArgb(150, Color.Pink))
                                                            End If

                                                            If CorNotasBraçoViolão.Color <> Color.Transparent Then
                                                                If ContadorNotasBraçoViolãoLinha = 0 AndAlso NumeraçãoTrastes = 1 Then
                                                                    gr.FillEllipse(CorNotasBraçoViolão, (ContadorNotasBraçoViolãoColuna * 15) + 32 + AjusteLeft(0), (ContadorNotasBraçoViolãoLinha * 15) + 174 + AjusteTopo, 11, 11)
                                                                    gr.DrawEllipse(Pens.Black, (ContadorNotasBraçoViolãoColuna * 15) + 32 + AjusteLeft(0), (ContadorNotasBraçoViolãoLinha * 15) + 174 + AjusteTopo, 11, 11)
                                                                ElseIf ContadorNotasBraçoViolãoLinha <> 0 Then
                                                                    gr.FillEllipse(CorNotasBraçoViolão, (ContadorNotasBraçoViolãoColuna * 15) + 32 + AjusteLeft(0), (ContadorNotasBraçoViolãoLinha * 15) + 167 + AjusteTopo, 11, 11)
                                                                    gr.DrawEllipse(Pens.Black, (ContadorNotasBraçoViolãoColuna * 15) + 32 + AjusteLeft(0), (ContadorNotasBraçoViolãoLinha * 15) + 167 + AjusteTopo, 11, 11)
                                                                End If
                                                            End If
                                                        Next
                                                    Next
                                                ElseIf My.Settings.NovoValorExibirIntervalosBraçoViolão = True Then
                                                    For ContadorNotasBraçoViolãoLinha = 0 To 6
                                                        For ContadorNotasBraçoViolãoColuna = 0 To 6

                                                            If ArrayAcordePadrão(a, ContadorNotasBraçoViolãoLinha, ContadorNotasBraçoViolãoColuna) = "T" OrElse ArrayAcordePadrão(a, ContadorNotasBraçoViolãoLinha, ContadorNotasBraçoViolãoColuna) = "1T" OrElse ArrayAcordePadrão(a, ContadorNotasBraçoViolãoLinha, ContadorNotasBraçoViolãoColuna) = "2T" OrElse ArrayAcordePadrão(a, ContadorNotasBraçoViolãoLinha, ContadorNotasBraçoViolãoColuna) = "3T" OrElse ArrayAcordePadrão(a, ContadorNotasBraçoViolãoLinha, ContadorNotasBraçoViolãoColuna) = "4T" OrElse ArrayAcordePadrão(a, ContadorNotasBraçoViolãoLinha, ContadorNotasBraçoViolãoColuna) = "PT" Then
                                                                TonicaLinha = ContadorNotasBraçoViolãoLinha
                                                                TonicaColuna = ContadorNotasBraçoViolãoColuna
                                                            End If

                                                        Next
                                                    Next

                                                    Dim AjusteTopoIntervalos As Integer = 0
                                                    If TonicaColuna = 1 OrElse TonicaColuna = 6 Then
                                                        AjusteTopoIntervalos = ((8 - TonicaLinha) * 15)
                                                    ElseIf TonicaColuna = 2 Then
                                                        AjusteTopoIntervalos = ((3 - TonicaLinha) * 15)
                                                    ElseIf TonicaColuna = 3 Then
                                                        AjusteTopoIntervalos = ((10 - TonicaLinha) * 15)
                                                    ElseIf TonicaColuna = 4 Then
                                                        AjusteTopoIntervalos = ((5 - TonicaLinha) * 15)
                                                    ElseIf TonicaColuna = 5 Then
                                                        AjusteTopoIntervalos = ((13 - TonicaLinha) * 15)
                                                    End If

                                                    ' Make a StringFormat object that centers.
                                                    Dim alinhamento As New StringFormat
                                                    alinhamento.LineAlignment = StringAlignment.Center
                                                    alinhamento.Alignment = StringAlignment.Center

                                                    For ContadorNotasBraçoViolãoLinha = 0 To 22
                                                        For ContadorNotasBraçoViolãoColuna = 0 To 6
                                                            CorNotasBraçoViolão = New SolidBrush(Color.Transparent)
                                                            If IntervalosAcorde2(ContadorNotasBraçoViolãoLinha, ContadorNotasBraçoViolãoColuna) = "T" Then
                                                                CorNotasBraçoViolão = New SolidBrush(Color.FromArgb(150, Color.Black))
                                                            ElseIf IntervalosAcorde2(ContadorNotasBraçoViolãoLinha, ContadorNotasBraçoViolãoColuna) = "9" Then
                                                                CorNotasBraçoViolão = New SolidBrush(Color.FromArgb(150, Color.Yellow))
                                                            ElseIf IntervalosAcorde2(ContadorNotasBraçoViolãoLinha, ContadorNotasBraçoViolãoColuna) = "3" Then
                                                                CorNotasBraçoViolão = New SolidBrush(Color.FromArgb(150, Color.Red))
                                                            ElseIf IntervalosAcorde2(ContadorNotasBraçoViolãoLinha, ContadorNotasBraçoViolãoColuna) = "4" Then
                                                                CorNotasBraçoViolão = New SolidBrush(Color.FromArgb(150, Color.SkyBlue))
                                                            ElseIf IntervalosAcorde2(ContadorNotasBraçoViolãoLinha, ContadorNotasBraçoViolãoColuna) = "5" Then
                                                                CorNotasBraçoViolão = New SolidBrush(Color.FromArgb(150, Color.Orange))
                                                            ElseIf IntervalosAcorde2(ContadorNotasBraçoViolãoLinha, ContadorNotasBraçoViolãoColuna) = "6" Then
                                                                CorNotasBraçoViolão = New SolidBrush(Color.FromArgb(150, Color.Green))
                                                            ElseIf IntervalosAcorde2(ContadorNotasBraçoViolãoLinha, ContadorNotasBraçoViolãoColuna) = "7M" Then
                                                                CorNotasBraçoViolão = New SolidBrush(Color.FromArgb(150, Color.Pink))
                                                            End If

                                                            'If CorNotasBraçoViolão.Color <> Color.Transparent Then
                                                            If ((ContadorNotasBraçoViolãoLinha * 15) + 166 + AjusteTopo + CInt(ArrayAcordePadrão(a, 8, 2)) - AjusteTopoIntervalos) > (160 + AjusteTopo) AndAlso ((ContadorNotasBraçoViolãoLinha * 15) + 166 + AjusteTopo + CInt(ArrayAcordePadrão(a, 8, 2)) - AjusteTopoIntervalos) < (260 + AjusteTopo) Then
                                                                If ((ContadorNotasBraçoViolãoLinha * 15) + 166 + AjusteTopo + CInt(ArrayAcordePadrão(a, 8, 2)) - AjusteTopoIntervalos) < (180 + AjusteTopo) AndAlso NumeraçãoTrastes = 1 Then
                                                                    gr.FillEllipse(CorNotasBraçoViolão, (ContadorNotasBraçoViolãoColuna * 15) + 32 + AjusteLeft(0), (ContadorNotasBraçoViolãoLinha * 15) + 174 + AjusteTopo + CInt(ArrayAcordePadrão(a, 8, 2)) - AjusteTopoIntervalos, 11, 11)
                                                                    If CorNotasBraçoViolão.Color <> Color.Transparent Then
                                                                        gr.DrawEllipse(Pens.Black, (ContadorNotasBraçoViolãoColuna * 15) + 32 + AjusteLeft(0), (ContadorNotasBraçoViolãoLinha * 15) + 174 + AjusteTopo + CInt(ArrayAcordePadrão(a, 8, 2)) - AjusteTopoIntervalos, 11, 11)
                                                                        If CorNotasBraçoViolão.Color = Color.FromArgb(150, Color.Black) OrElse CorNotasBraçoViolão.Color = Color.FromArgb(150, Color.Green) Then
                                                                            gr.DrawString(IntervalosAcorde2(ContadorNotasBraçoViolãoLinha, ContadorNotasBraçoViolãoColuna).Replace("T", "F"), Fonte2, Brushes.White, (ContadorNotasBraçoViolãoColuna * 15) + 38 + AjusteLeft(0), (ContadorNotasBraçoViolãoLinha * 15) + 180 + AjusteTopo + CInt(ArrayAcordePadrão(a, 8, 2)) - AjusteTopoIntervalos, alinhamento)
                                                                        Else
                                                                            gr.DrawString(IntervalosAcorde2(ContadorNotasBraçoViolãoLinha, ContadorNotasBraçoViolãoColuna), Fonte2, Brushes.Black, (ContadorNotasBraçoViolãoColuna * 15) + 38 + AjusteLeft(0), (ContadorNotasBraçoViolãoLinha * 15) + 180 + AjusteTopo + CInt(ArrayAcordePadrão(a, 8, 2)) - AjusteTopoIntervalos, alinhamento)
                                                                        End If
                                                                    Else
                                                                        gr.DrawString(IntervalosAcorde2(ContadorNotasBraçoViolãoLinha, ContadorNotasBraçoViolãoColuna), Fonte2, Brushes.Green, (ContadorNotasBraçoViolãoColuna * 15) + 38 + AjusteLeft(0), (ContadorNotasBraçoViolãoLinha * 15) + 180 + AjusteTopo + CInt(ArrayAcordePadrão(a, 8, 2)) - AjusteTopoIntervalos, alinhamento)
                                                                    End If
                                                                ElseIf ((ContadorNotasBraçoViolãoLinha * 15) + 166 + AjusteTopo + CInt(ArrayAcordePadrão(a, 8, 2)) - AjusteTopoIntervalos) > (180 + AjusteTopo) Then
                                                                    gr.FillEllipse(CorNotasBraçoViolão, (ContadorNotasBraçoViolãoColuna * 15) + 32 + AjusteLeft(0), (ContadorNotasBraçoViolãoLinha * 15) + 167 + AjusteTopo + CInt(ArrayAcordePadrão(a, 8, 2)) - AjusteTopoIntervalos, 11, 11)
                                                                    If CorNotasBraçoViolão.Color <> Color.Transparent Then
                                                                        gr.DrawEllipse(Pens.Black, (ContadorNotasBraçoViolãoColuna * 15) + 32 + AjusteLeft(0), (ContadorNotasBraçoViolãoLinha * 15) + 167 + AjusteTopo + CInt(ArrayAcordePadrão(a, 8, 2)) - AjusteTopoIntervalos, 11, 11)
                                                                        If CorNotasBraçoViolão.Color = Color.FromArgb(150, Color.Black) OrElse CorNotasBraçoViolão.Color = Color.FromArgb(150, Color.Green) Then
                                                                            gr.DrawString(IntervalosAcorde2(ContadorNotasBraçoViolãoLinha, ContadorNotasBraçoViolãoColuna).Replace("T", "F"), Fonte2, Brushes.White, (ContadorNotasBraçoViolãoColuna * 15) + 38 + AjusteLeft(0), (ContadorNotasBraçoViolãoLinha * 15) + 173 + AjusteTopo + CInt(ArrayAcordePadrão(a, 8, 2)) - AjusteTopoIntervalos, alinhamento)
                                                                        Else
                                                                            gr.DrawString(IntervalosAcorde2(ContadorNotasBraçoViolãoLinha, ContadorNotasBraçoViolãoColuna), Fonte2, Brushes.Black, (ContadorNotasBraçoViolãoColuna * 15) + 38 + AjusteLeft(0), (ContadorNotasBraçoViolãoLinha * 15) + 173 + AjusteTopo + CInt(ArrayAcordePadrão(a, 8, 2)) - AjusteTopoIntervalos, alinhamento)
                                                                        End If
                                                                    Else
                                                                        gr.DrawString(IntervalosAcorde2(ContadorNotasBraçoViolãoLinha, ContadorNotasBraçoViolãoColuna), Fonte2, Brushes.Green, (ContadorNotasBraçoViolãoColuna * 15) + 38 + AjusteLeft(0), (ContadorNotasBraçoViolãoLinha * 15) + 173 + AjusteTopo + CInt(ArrayAcordePadrão(a, 8, 2)) - AjusteTopoIntervalos, alinhamento)
                                                                    End If
                                                                End If
                                                            End If
                                                            'End If
                                                        Next
                                                    Next
                                                End If


                                            End If
                                        ElseIf c = 7 AndAlso ImagemAcorde(a) IsNot Nothing Then
                                            gr.FillEllipse(Brushes.Black, (b * 15) + 34 + AjusteLeft(0), (c * 15) + 161 + AjusteTopo, 8, 8)
                                            If ArrayAcordePadrão(a, 7, b) <> "" Then d += 1
                                            If d = 1 Then
                                                gr.FillEllipse(Brushes.Black, (b * 15) + 35 + AjusteLeft(0), (c * 15) + 162 + AjusteTopo, 6, 6)
                                            Else
                                                gr.FillEllipse(Brushes.White, (b * 15) + 35 + AjusteLeft(0), (c * 15) + 162 + AjusteTopo, 6, 6)
                                            End If
                                        End If
                                        If c <= 6 AndAlso ArrayAcordePadrão(a, c, b) <> "" Then

                                            If IntervalosAcorde(c + AjustePosição(0), b) = "b3" OrElse IntervalosAcorde(c + AjustePosição(0), b) = "11" OrElse _
                                                IntervalosAcorde(c + AjustePosição(0), b) = "b5" OrElse IntervalosAcorde(c + AjustePosição(0), b) = "13" OrElse _
                                                IntervalosAcorde(c + AjustePosição(0), b) = "b9" OrElse IntervalosAcorde(c + AjustePosição(0), b) = "3" Then
                                                AjusteLeftIntervalos = 1
                                            Else
                                                AjusteLeftIntervalos = 0
                                            End If

                                            If IntervalosAcorde(c + AjustePosição(0), b) = "T" OrElse _
                                               IntervalosAcorde(c + AjustePosição(0), b) = "3" OrElse _
                                               IntervalosAcorde(c + AjustePosição(0), b) = "5" OrElse _
                                               IntervalosAcorde(c + AjustePosição(0), b) = "b3" Then
                                                Cor = CorFonte2
                                            Else
                                                Cor = New SolidBrush(My.Settings.NovaCorIntervalos) 'CorFonte3
                                            End If


                                            If My.Settings.NovoValorExibirNomeIntervalos = True Then
                                                If (ArrayAcordePadrão(a, c, b) = "P" OrElse ArrayAcordePadrão(a, c, b) = "PT") AndAlso (ArrayAcordePadrão(a, c + 1, b) = "1" OrElse ArrayAcordePadrão(a, c + 1, b) = "2" OrElse ArrayAcordePadrão(a, c + 1, b) = "3" OrElse ArrayAcordePadrão(a, c + 1, b) = "4" OrElse _
                                                                                           ArrayAcordePadrão(a, c + 2, b) = "1" OrElse ArrayAcordePadrão(a, c + 2, b) = "2" OrElse ArrayAcordePadrão(a, c + 2, b) = "3" OrElse ArrayAcordePadrão(a, c + 2, b) = "4" OrElse _
                                                                                           ArrayAcordePadrão(a, c + 3, b) = "1" OrElse ArrayAcordePadrão(a, c + 3, b) = "2" OrElse ArrayAcordePadrão(a, c + 3, b) = "3" OrElse ArrayAcordePadrão(a, c + 3, b) = "4" OrElse _
                                                                                           ArrayAcordePadrão(a, c + 4, b) = "1" OrElse ArrayAcordePadrão(a, c + 4, b) = "2" OrElse ArrayAcordePadrão(a, c + 4, b) = "3" OrElse ArrayAcordePadrão(a, c + 4, b) = "4" OrElse _
                                                                                           ArrayAcordePadrão(a, c + 5, b) = "1" OrElse ArrayAcordePadrão(a, c + 5, b) = "2" OrElse ArrayAcordePadrão(a, c + 5, b) = "3" OrElse ArrayAcordePadrão(a, c + 5, b) = "4" OrElse _
                                                                                           ArrayAcordePadrão(a, c + 1, b) = "1T" OrElse ArrayAcordePadrão(a, c + 1, b) = "2T" OrElse ArrayAcordePadrão(a, c + 1, b) = "3T" OrElse ArrayAcordePadrão(a, c + 1, b) = "4T" OrElse _
                                                                                           ArrayAcordePadrão(a, c + 2, b) = "1T" OrElse ArrayAcordePadrão(a, c + 2, b) = "2T" OrElse ArrayAcordePadrão(a, c + 2, b) = "3T" OrElse ArrayAcordePadrão(a, c + 2, b) = "4T" OrElse _
                                                                                           ArrayAcordePadrão(a, c + 3, b) = "1T" OrElse ArrayAcordePadrão(a, c + 3, b) = "2T" OrElse ArrayAcordePadrão(a, c + 3, b) = "3T" OrElse ArrayAcordePadrão(a, c + 3, b) = "4T" OrElse _
                                                                                           ArrayAcordePadrão(a, c + 4, b) = "1T" OrElse ArrayAcordePadrão(a, c + 4, b) = "2T" OrElse ArrayAcordePadrão(a, c + 4, b) = "3T" OrElse ArrayAcordePadrão(a, c + 4, b) = "4T" OrElse _
                                                                                           ArrayAcordePadrão(a, c + 5, b) = "1T" OrElse ArrayAcordePadrão(a, c + 5, b) = "2T" OrElse ArrayAcordePadrão(a, c + 5, b) = "3T" OrElse ArrayAcordePadrão(a, c + 5, b) = "4T" OrElse _
                                                                                           ArrayAcordePadrão(a, c + 1, b) = "PS" OrElse ArrayAcordePadrão(a, c + 1, b) = "1S" OrElse ArrayAcordePadrão(a, c + 1, b) = "2S" OrElse ArrayAcordePadrão(a, c + 1, b) = "3S" OrElse ArrayAcordePadrão(a, c + 1, b) = "4S" OrElse _
                                                                                           ArrayAcordePadrão(a, c + 2, b) = "1S" OrElse ArrayAcordePadrão(a, c + 2, b) = "2S" OrElse ArrayAcordePadrão(a, c + 2, b) = "3S" OrElse ArrayAcordePadrão(a, c + 2, b) = "4S" OrElse _
                                                                                           ArrayAcordePadrão(a, c + 3, b) = "1S" OrElse ArrayAcordePadrão(a, c + 3, b) = "2S" OrElse ArrayAcordePadrão(a, c + 3, b) = "3S" OrElse ArrayAcordePadrão(a, c + 3, b) = "4S" OrElse _
                                                                                           ArrayAcordePadrão(a, c + 4, b) = "1S" OrElse ArrayAcordePadrão(a, c + 4, b) = "2S" OrElse ArrayAcordePadrão(a, c + 4, b) = "3S" OrElse ArrayAcordePadrão(a, c + 4, b) = "4S" OrElse _
                                                                                           ArrayAcordePadrão(a, c + 5, b) = "1S" OrElse ArrayAcordePadrão(a, c + 5, b) = "2S" OrElse ArrayAcordePadrão(a, c + 5, b) = "3S" OrElse ArrayAcordePadrão(a, c + 5, b) = "4S") Then
                                                    'não executa nada
                                                Else
                                                    If modoJogo = False Then gr.DrawString(IntervalosAcorde(c + AjustePosição(0), b), Fonte3, Cor, (b * 15) + 38 + AjusteLeft(0) + AjusteLeftIntervalos, 279 + AjusteTopo, sf)
                                                End If
                                            End If

                                            If My.Settings.NovoValorExibirNomeDasNotasAcordes = True Then
                                                If (ArrayAcordePadrão(a, c, b) = "P" OrElse ArrayAcordePadrão(a, c, b) = "PT") AndAlso (ArrayAcordePadrão(a, c + 1, b) = "1" OrElse ArrayAcordePadrão(a, c + 1, b) = "2" OrElse ArrayAcordePadrão(a, c + 1, b) = "3" OrElse ArrayAcordePadrão(a, c + 1, b) = "4" OrElse _
                                                                                               ArrayAcordePadrão(a, c + 2, b) = "1" OrElse ArrayAcordePadrão(a, c + 2, b) = "2" OrElse ArrayAcordePadrão(a, c + 2, b) = "3" OrElse ArrayAcordePadrão(a, c + 2, b) = "4" OrElse _
                                                                                               ArrayAcordePadrão(a, c + 3, b) = "1" OrElse ArrayAcordePadrão(a, c + 3, b) = "2" OrElse ArrayAcordePadrão(a, c + 3, b) = "3" OrElse ArrayAcordePadrão(a, c + 3, b) = "4" OrElse _
                                                                                               ArrayAcordePadrão(a, c + 4, b) = "1" OrElse ArrayAcordePadrão(a, c + 4, b) = "2" OrElse ArrayAcordePadrão(a, c + 4, b) = "3" OrElse ArrayAcordePadrão(a, c + 4, b) = "4" OrElse _
                                                                                               ArrayAcordePadrão(a, c + 5, b) = "1" OrElse ArrayAcordePadrão(a, c + 5, b) = "2" OrElse ArrayAcordePadrão(a, c + 5, b) = "3" OrElse ArrayAcordePadrão(a, c + 5, b) = "4" OrElse _
                                                                                               ArrayAcordePadrão(a, c + 1, b) = "1T" OrElse ArrayAcordePadrão(a, c + 1, b) = "2T" OrElse ArrayAcordePadrão(a, c + 1, b) = "3T" OrElse ArrayAcordePadrão(a, c + 1, b) = "4T" OrElse _
                                                                                               ArrayAcordePadrão(a, c + 2, b) = "1T" OrElse ArrayAcordePadrão(a, c + 2, b) = "2T" OrElse ArrayAcordePadrão(a, c + 2, b) = "3T" OrElse ArrayAcordePadrão(a, c + 2, b) = "4T" OrElse _
                                                                                               ArrayAcordePadrão(a, c + 3, b) = "1T" OrElse ArrayAcordePadrão(a, c + 3, b) = "2T" OrElse ArrayAcordePadrão(a, c + 3, b) = "3T" OrElse ArrayAcordePadrão(a, c + 3, b) = "4T" OrElse _
                                                                                               ArrayAcordePadrão(a, c + 4, b) = "1T" OrElse ArrayAcordePadrão(a, c + 4, b) = "2T" OrElse ArrayAcordePadrão(a, c + 4, b) = "3T" OrElse ArrayAcordePadrão(a, c + 4, b) = "4T" OrElse _
                                                                                               ArrayAcordePadrão(a, c + 5, b) = "1T" OrElse ArrayAcordePadrão(a, c + 5, b) = "2T" OrElse ArrayAcordePadrão(a, c + 5, b) = "3T" OrElse ArrayAcordePadrão(a, c + 5, b) = "4T" OrElse _
                                                                                               ArrayAcordePadrão(a, c + 1, b) = "PS" OrElse ArrayAcordePadrão(a, c + 1, b) = "1S" OrElse ArrayAcordePadrão(a, c + 1, b) = "2S" OrElse ArrayAcordePadrão(a, c + 1, b) = "3S" OrElse ArrayAcordePadrão(a, c + 1, b) = "4S" OrElse _
                                                                                               ArrayAcordePadrão(a, c + 2, b) = "1S" OrElse ArrayAcordePadrão(a, c + 2, b) = "2S" OrElse ArrayAcordePadrão(a, c + 2, b) = "3S" OrElse ArrayAcordePadrão(a, c + 2, b) = "4S" OrElse _
                                                                                               ArrayAcordePadrão(a, c + 3, b) = "1S" OrElse ArrayAcordePadrão(a, c + 3, b) = "2S" OrElse ArrayAcordePadrão(a, c + 3, b) = "3S" OrElse ArrayAcordePadrão(a, c + 3, b) = "4S" OrElse _
                                                                                               ArrayAcordePadrão(a, c + 4, b) = "1S" OrElse ArrayAcordePadrão(a, c + 4, b) = "2S" OrElse ArrayAcordePadrão(a, c + 4, b) = "3S" OrElse ArrayAcordePadrão(a, c + 4, b) = "4S" OrElse _
                                                                                               ArrayAcordePadrão(a, c + 5, b) = "1S" OrElse ArrayAcordePadrão(a, c + 5, b) = "2S" OrElse ArrayAcordePadrão(a, c + 5, b) = "3S" OrElse ArrayAcordePadrão(a, c + 5, b) = "4S") Then
                                                    'não executa nada
                                                Else
                                                    NotasAcordeIndiceColuna = 0
                                                    If IntervalosAcorde(c + AjustePosição(0), b) = "T" Then
                                                        NotasAcordeIndiceColuna = 1
                                                    ElseIf IntervalosAcorde(c + AjustePosição(0), b) = "b9" Then
                                                        NotasAcordeIndiceColuna = 2
                                                    ElseIf IntervalosAcorde(c + AjustePosição(0), b) = "9" Then
                                                        NotasAcordeIndiceColuna = 3
                                                    ElseIf IntervalosAcorde(c + AjustePosição(0), b) = "#9" Then
                                                        NotasAcordeIndiceColuna = 4
                                                    ElseIf IntervalosAcorde(c + AjustePosição(0), b) = "b3" Then
                                                        NotasAcordeIndiceColuna = 5
                                                    ElseIf IntervalosAcorde(c + AjustePosição(0), b) = "3" Then
                                                        NotasAcordeIndiceColuna = 6
                                                    ElseIf IntervalosAcorde(c + AjustePosição(0), b) = "4" OrElse IntervalosAcorde(c + AjustePosição(0), b) = "11" Then
                                                        NotasAcordeIndiceColuna = 7
                                                    ElseIf IntervalosAcorde(c + AjustePosição(0), b) = "#11" Then
                                                        NotasAcordeIndiceColuna = 8
                                                    ElseIf IntervalosAcorde(c + AjustePosição(0), b) = "b5" Then
                                                        NotasAcordeIndiceColuna = 9
                                                    ElseIf IntervalosAcorde(c + AjustePosição(0), b) = "5" Then
                                                        NotasAcordeIndiceColuna = 10
                                                    ElseIf IntervalosAcorde(c + AjustePosição(0), b) = "#5" Then
                                                        NotasAcordeIndiceColuna = 11
                                                    ElseIf IntervalosAcorde(c + AjustePosição(0), b) = "b6" OrElse IntervalosAcorde(c + AjustePosição(0), b) = "b13" Then
                                                        NotasAcordeIndiceColuna = 12
                                                    ElseIf IntervalosAcorde(c + AjustePosição(0), b) = "6" OrElse IntervalosAcorde(c + AjustePosição(0), b) = "13" OrElse IntervalosAcorde(c + AjustePosição(0), b) = "º" Then
                                                        NotasAcordeIndiceColuna = 13
                                                    ElseIf IntervalosAcorde(c + AjustePosição(0), b) = "7" Then
                                                        NotasAcordeIndiceColuna = 14
                                                    ElseIf IntervalosAcorde(c + AjustePosição(0), b) = "7M" Then
                                                        NotasAcordeIndiceColuna = 15
                                                    End If
                                                    If NotasAcordeIndiceColuna <> 0 AndAlso modoJogo = False Then gr.DrawString(NotasAcorde(NotasAcordeIndiceLinha, NotasAcordeIndiceColuna), Fonte3, Cor, (b * 15) + 38 + AjusteLeft(0) + AjusteLeftIntervalos, 279 + AjusteTopo, sf)
                                                End If
                                            End If


                                        End If
                                    End If
                                Next

                                AjusteLeftIntervalos = 0
                                If IntervalosAcorde(Pestana + AjustePosição(0), b) = "b3" OrElse IntervalosAcorde(Pestana + AjustePosição(0), b) = "11" OrElse _
                                                IntervalosAcorde(Pestana + AjustePosição(0), b) = "b5" OrElse IntervalosAcorde(Pestana + AjustePosição(0), b) = "13" OrElse _
                                                IntervalosAcorde(Pestana + AjustePosição(0), b) = "b9" OrElse IntervalosAcorde(Pestana + AjustePosição(0), b) = "3" OrElse _
                                    IntervalosAcorde(AjustePosição(0) + 1, b) = "b3" OrElse IntervalosAcorde(AjustePosição(0) + 1, b) = "11" OrElse _
                                                IntervalosAcorde(AjustePosição(0) + 1, b) = "b5" OrElse IntervalosAcorde(AjustePosição(0) + 1, b) = "13" OrElse _
                                                IntervalosAcorde(AjustePosição(0) + 1, b) = "b9" OrElse IntervalosAcorde(AjustePosição(0) + 1, b) = "3" Then
                                    AjusteLeftIntervalos = 1
                                Else
                                    AjusteLeftIntervalos = 0
                                End If

                                If IntervalosAcorde(Pestana + AjustePosição(0), b) = "T" OrElse _
                                           IntervalosAcorde(Pestana + AjustePosição(0), b) = "3" OrElse _
                                           IntervalosAcorde(Pestana + AjustePosição(0), b) = "5" OrElse _
                                           IntervalosAcorde(Pestana + AjustePosição(0), b) = "b3" Then
                                    Cor = CorFonte2
                                Else
                                    Cor = New SolidBrush(My.Settings.NovaCorIntervalos) 'CorFonte3
                                End If

                                If My.Settings.NovoValorExibirNomeIntervalos = True Then
                                    If ArrayAcordePadrão(a, 0, b) = "" AndAlso ArrayAcordePadrão(a, 1, b) = "" AndAlso ArrayAcordePadrão(a, 2, b) = "" AndAlso ArrayAcordePadrão(a, 3, b) = "" AndAlso ArrayAcordePadrão(a, 4, b) = "" AndAlso ArrayAcordePadrão(a, 5, b) = "" AndAlso ArrayAcordePadrão(a, 6, b) = "" Then
                                        If ArrayAcordePadrão(a, 7, b) = "B" AndAlso Pestana = 0 Then
                                            If modoJogo = False Then gr.DrawString(IntervalosAcorde(0 + AjustePosição(1), b), Fonte3, Cor, (b * 15) + 38 + AjusteLeft(0), 279 + AjusteTopo, sf)
                                        ElseIf ArrayAcordePadrão(a, 7, b) = "B" AndAlso Pestana > 0 Then
                                            If modoJogo = False Then gr.DrawString(IntervalosAcorde(Pestana + AjustePosição(0), b), Fonte3, Cor, (b * 15) + 38 + AjusteLeft(0), 279 + AjusteTopo, sf)
                                        End If
                                    End If
                                End If

                                If My.Settings.NovoValorExibirNomeDasNotasAcordes = True Then
                                    If ArrayAcordePadrão(a, 0, b) = "" AndAlso ArrayAcordePadrão(a, 1, b) = "" AndAlso ArrayAcordePadrão(a, 2, b) = "" AndAlso ArrayAcordePadrão(a, 3, b) = "" AndAlso ArrayAcordePadrão(a, 4, b) = "" AndAlso ArrayAcordePadrão(a, 5, b) = "" AndAlso ArrayAcordePadrão(a, 6, b) = "" Then
                                        NotasAcordeIndiceColuna = 0
                                        If IntervalosAcorde(Pestana + AjustePosição(0), b) = "T" Then
                                            NotasAcordeIndiceColuna = 1
                                        ElseIf IntervalosAcorde(Pestana + AjustePosição(0), b) = "b9" Then
                                            NotasAcordeIndiceColuna = 2
                                        ElseIf IntervalosAcorde(Pestana + AjustePosição(0), b) = "9" Then
                                            NotasAcordeIndiceColuna = 3
                                        ElseIf IntervalosAcorde(Pestana + AjustePosição(0), b) = "#9" Then
                                            NotasAcordeIndiceColuna = 4
                                        ElseIf IntervalosAcorde(Pestana + AjustePosição(0), b) = "b3" Then
                                            NotasAcordeIndiceColuna = 5
                                        ElseIf IntervalosAcorde(Pestana + AjustePosição(0), b) = "3" Then
                                            NotasAcordeIndiceColuna = 6
                                        ElseIf IntervalosAcorde(Pestana + AjustePosição(0), b) = "4" OrElse IntervalosAcorde(Pestana + AjustePosição(0), b) = "11" Then
                                            NotasAcordeIndiceColuna = 7
                                        ElseIf IntervalosAcorde(Pestana + AjustePosição(0), b) = "#11" Then
                                            NotasAcordeIndiceColuna = 8
                                        ElseIf IntervalosAcorde(Pestana + AjustePosição(0), b) = "b5" Then
                                            NotasAcordeIndiceColuna = 9
                                        ElseIf IntervalosAcorde(Pestana + AjustePosição(0), b) = "5" Then
                                            NotasAcordeIndiceColuna = 10
                                        ElseIf IntervalosAcorde(Pestana + AjustePosição(0), b) = "#5" Then
                                            NotasAcordeIndiceColuna = 11
                                        ElseIf IntervalosAcorde(Pestana + AjustePosição(0), b) = "b6" OrElse IntervalosAcorde(Pestana + AjustePosição(0), b) = "b13" Then
                                            NotasAcordeIndiceColuna = 12
                                        ElseIf IntervalosAcorde(Pestana + AjustePosição(0), b) = "6" OrElse IntervalosAcorde(Pestana + AjustePosição(0), b) = "13" OrElse IntervalosAcorde(Pestana + AjustePosição(0), b) = "º" Then
                                            NotasAcordeIndiceColuna = 13
                                        ElseIf IntervalosAcorde(Pestana + AjustePosição(0), b) = "7" Then
                                            NotasAcordeIndiceColuna = 14
                                        ElseIf IntervalosAcorde(Pestana + AjustePosição(0), b) = "7M" Then
                                            NotasAcordeIndiceColuna = 15
                                        End If

                                        If ArrayAcordePadrão(a, 7, b) = "B" AndAlso NotasAcordeIndiceColuna <> 0 Then
                                            If modoJogo = False Then gr.DrawString(NotasAcorde(NotasAcordeIndiceLinha, NotasAcordeIndiceColuna), Fonte3, Cor, (b * 15) + 38 + AjusteLeft(0) + AjusteLeftIntervalos, 279 + AjusteTopo, sf)
                                        End If

                                    End If
                                End If


                            Next


                        End If
                    End If
                Next

                If FamiliaAcorde = "Sétima e 4ª com 9ª e 13ª" OrElse NumeraçãoFamiliaAcorde = 52 Then
                    If ImagemAcorde(i) <> "G7(4/9/13)" Then
                        gr.DrawArc(Pens.Black, 239, 183, 25, 25, 180, 80)
                    End If
                End If

                gr.SmoothingMode = SmoothingMode.None


                If i = 1 Then

                    Me.SetBitmap(FaceBit, TransAmount)
                    ImagemCopiada(1) = FaceBit
                Else

                    'copia acordes para a área de transferência, com isso é possível colá-los em softwares como Word, Excel, Fireworks etc
                    If LarguraImagemCopiada > 0 Then
                        For ii = 1 To 8
                            If ImagemAcorde(ii) = "" Then
                                LarguraImagemCopiada = ((ii - 1) * 128) - 32
                                ii = 100
                            End If
                        Next

                        ImagemCopiada(0) = FaceBit.Clone(New Rectangle(42 * FatorEscala, 165 * FatorEscala, LarguraImagemCopiada * FatorEscala, (AjusteTopo + 121) * FatorEscala), FaceBit.PixelFormat)
                        ImagemCopiada(0).SetResolution(96 * FatorEscala, 96 * FatorEscala)
                        Clipboard.SetDataObject(ImagemCopiada(0))
                        ' ImagemCopiada.Save("C:\Teste" & FatorEscala & ".bmp")

                        Dim gr3 As Graphics
                        Dim FaceBit3 As Bitmap
                        FaceBit3 = New Bitmap(ImagemCopiada(1))
                        gr3 = Graphics.FromImage(FaceBit3)
                        gr3.TextRenderingHint = TextRenderingHint.ClearTypeGridFit
                        CenterTextAt(gr3, "Imagem Copiada.    Fator: " & FatorEscala & ".    " & ImagemCopiada(0).HorizontalResolution & " dpi.    Tamanho: " & ImagemCopiada(0).Width & "x" & ImagemCopiada(0).Height & " px.    Reduza para " & FormatNumber(100 / FatorEscala, 2) & "%", CInt(Me.Width / 2), 750)
                        Me.SetBitmap(FaceBit3, TransAmount)
                    End If
                End If


                If My.Settings.NovoValorCopiarImagemDosAcordes = False Then i = 100 'sai da rotina caso "copiar" não esteja ativado

            Next


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub Acordes_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.DoubleClick

        Try

            CopiarAcorde()


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Acordes_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'MsgBox("Nesta versão a maioria dos acordes é exibida corretamente quando se muda de tom." & Chr(13) & _
        '"Erros em alguns acordes:" & Chr(13) & _
        '"    - Posição dos dedos deveria mudar quando o acorde utiliza cordas soltas" & Chr(13) & Chr(13) & _
        '"Ir usando e aproveitando o que der.... ", MsgBoxStyle.Exclamation, "Atenção")

        CorIntervalos.BackColor = My.Settings.NovaCorIntervalos

        AjustePosiçãoTonalidadeDoAcorde = 0
        TonalidadeAcorde = "C"
        NumeraçãoFamiliaAcorde = 1
        LocalizaNomeFamiliaAcordes()
        GerarAcordes()
    End Sub

    Private Sub ab_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseDown

        Try
            VGa = Me.MousePosition.X - Me.Location.X
            VGb = Me.MousePosition.Y - Me.Location.Y

            If e.Button = MouseButtons.Right OrElse e.Button = MouseButtons.Middle Then CopiarAcorde()


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub ab_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove

        If e.Button = MouseButtons.Left Then
            VGnewPoint = Me.MousePosition
            VGnewPoint.X -= VGa
            VGnewPoint.Y -= VGb
            Me.Location = VGnewPoint
        End If

        ValorEsquerda = -1
        If e.X >= 42 AndAlso e.X <= 138 AndAlso e.Y >= 165 AndAlso e.Y <= 286 Then
            ValorEsquerda = 0
            ValorTopo = 0
        ElseIf e.X >= 170 AndAlso e.X <= 266 AndAlso e.Y >= 165 AndAlso e.Y <= 286 Then
            ValorEsquerda = 128
            ValorTopo = 0
        ElseIf e.X >= 298 AndAlso e.X <= 394 AndAlso e.Y >= 165 AndAlso e.Y <= 286 Then
            ValorEsquerda = 256
            ValorTopo = 0
        ElseIf e.X >= 426 AndAlso e.X <= 522 AndAlso e.Y >= 165 AndAlso e.Y <= 286 Then
            ValorEsquerda = 384
            ValorTopo = 0
        ElseIf e.X >= 554 AndAlso e.X <= 650 AndAlso e.Y >= 165 AndAlso e.Y <= 286 Then
            ValorEsquerda = 512
            ValorTopo = 0
        ElseIf e.X >= 682 AndAlso e.X <= 778 AndAlso e.Y >= 165 AndAlso e.Y <= 286 Then
            ValorEsquerda = 640
            ValorTopo = 0
        ElseIf e.X >= 810 AndAlso e.X <= 906 AndAlso e.Y >= 165 AndAlso e.Y <= 286 Then
            ValorEsquerda = 768
            ValorTopo = 0
        ElseIf e.X >= 938 AndAlso e.X <= 1034 AndAlso e.Y >= 165 AndAlso e.Y <= 286 Then
            ValorEsquerda = 896
            ValorTopo = 0



        ElseIf e.X >= 42 AndAlso e.X <= 138 AndAlso e.Y >= 283 AndAlso e.Y <= 404 Then
            ValorEsquerda = 0
            ValorTopo = 118
        ElseIf e.X >= 170 AndAlso e.X <= 266 AndAlso e.Y >= 283 AndAlso e.Y <= 404 Then
            ValorEsquerda = 128
            ValorTopo = 118
        ElseIf e.X >= 298 AndAlso e.X <= 394 AndAlso e.Y >= 283 AndAlso e.Y <= 404 Then
            ValorEsquerda = 256
            ValorTopo = 118
        ElseIf e.X >= 426 AndAlso e.X <= 522 AndAlso e.Y >= 283 AndAlso e.Y <= 404 Then
            ValorEsquerda = 384
            ValorTopo = 118
        ElseIf e.X >= 554 AndAlso e.X <= 650 AndAlso e.Y >= 283 AndAlso e.Y <= 404 Then
            ValorEsquerda = 512
            ValorTopo = 118
        ElseIf e.X >= 682 AndAlso e.X <= 778 AndAlso e.Y >= 283 AndAlso e.Y <= 404 Then
            ValorEsquerda = 640
            ValorTopo = 118
        ElseIf e.X >= 810 AndAlso e.X <= 906 AndAlso e.Y >= 283 AndAlso e.Y <= 404 Then
            ValorEsquerda = 768
            ValorTopo = 118
        ElseIf e.X >= 938 AndAlso e.X <= 1034 AndAlso e.Y >= 283 AndAlso e.Y <= 404 Then
            ValorEsquerda = 896
            ValorTopo = 118



        ElseIf e.X >= 42 AndAlso e.X <= 138 AndAlso e.Y >= 401 AndAlso e.Y <= 522 Then
            ValorEsquerda = 0
            ValorTopo = 236
        ElseIf e.X >= 170 AndAlso e.X <= 266 AndAlso e.Y >= 401 AndAlso e.Y <= 522 Then
            ValorEsquerda = 128
            ValorTopo = 236
        ElseIf e.X >= 298 AndAlso e.X <= 394 AndAlso e.Y >= 401 AndAlso e.Y <= 522 Then
            ValorEsquerda = 256
            ValorTopo = 236
        ElseIf e.X >= 426 AndAlso e.X <= 522 AndAlso e.Y >= 401 AndAlso e.Y <= 522 Then
            ValorEsquerda = 384
            ValorTopo = 236
        ElseIf e.X >= 554 AndAlso e.X <= 650 AndAlso e.Y >= 401 AndAlso e.Y <= 522 Then
            ValorEsquerda = 512
            ValorTopo = 236
        ElseIf e.X >= 682 AndAlso e.X <= 778 AndAlso e.Y >= 401 AndAlso e.Y <= 522 Then
            ValorEsquerda = 640
            ValorTopo = 236
        ElseIf e.X >= 810 AndAlso e.X <= 906 AndAlso e.Y >= 401 AndAlso e.Y <= 522 Then
            ValorEsquerda = 768
            ValorTopo = 236
        ElseIf e.X >= 938 AndAlso e.X <= 1034 AndAlso e.Y >= 401 AndAlso e.Y <= 522 Then
            ValorEsquerda = 896
            ValorTopo = 236



        ElseIf e.X >= 42 AndAlso e.X <= 138 AndAlso e.Y >= 519 AndAlso e.Y <= 640 Then
            ValorEsquerda = 0
            ValorTopo = 354
        ElseIf e.X >= 170 AndAlso e.X <= 266 AndAlso e.Y >= 519 AndAlso e.Y <= 640 Then
            ValorEsquerda = 128
            ValorTopo = 354
        ElseIf e.X >= 298 AndAlso e.X <= 394 AndAlso e.Y >= 519 AndAlso e.Y <= 640 Then
            ValorEsquerda = 256
            ValorTopo = 354
        ElseIf e.X >= 426 AndAlso e.X <= 522 AndAlso e.Y >= 519 AndAlso e.Y <= 640 Then
            ValorEsquerda = 384
            ValorTopo = 354
        ElseIf e.X >= 554 AndAlso e.X <= 650 AndAlso e.Y >= 519 AndAlso e.Y <= 640 Then
            ValorEsquerda = 512
            ValorTopo = 354
        ElseIf e.X >= 682 AndAlso e.X <= 778 AndAlso e.Y >= 519 AndAlso e.Y <= 640 Then
            ValorEsquerda = 640
            ValorTopo = 354
        ElseIf e.X >= 810 AndAlso e.X <= 906 AndAlso e.Y >= 519 AndAlso e.Y <= 640 Then
            ValorEsquerda = 768
            ValorTopo = 354
        ElseIf e.X >= 938 AndAlso e.X <= 1034 AndAlso e.Y >= 519 AndAlso e.Y <= 640 Then
            ValorEsquerda = 896
            ValorTopo = 354



        ElseIf e.X >= 42 AndAlso e.X <= 138 AndAlso e.Y >= 637 AndAlso e.Y <= 758 Then
            ValorEsquerda = 0
            ValorTopo = 472
        ElseIf e.X >= 170 AndAlso e.X <= 266 AndAlso e.Y >= 637 AndAlso e.Y <= 758 Then
            ValorEsquerda = 128
            ValorTopo = 472
        ElseIf e.X >= 298 AndAlso e.X <= 394 AndAlso e.Y >= 637 AndAlso e.Y <= 758 Then
            ValorEsquerda = 256
            ValorTopo = 472
        ElseIf e.X >= 426 AndAlso e.X <= 522 AndAlso e.Y >= 637 AndAlso e.Y <= 758 Then
            ValorEsquerda = 384
            ValorTopo = 472
        ElseIf e.X >= 554 AndAlso e.X <= 650 AndAlso e.Y >= 637 AndAlso e.Y <= 758 Then
            ValorEsquerda = 512
            ValorTopo = 472
        ElseIf e.X >= 682 AndAlso e.X <= 778 AndAlso e.Y >= 637 AndAlso e.Y <= 758 Then
            ValorEsquerda = 640
            ValorTopo = 472
        ElseIf e.X >= 810 AndAlso e.X <= 906 AndAlso e.Y >= 637 AndAlso e.Y <= 758 Then
            ValorEsquerda = 768
            ValorTopo = 472
        ElseIf e.X >= 938 AndAlso e.X <= 1034 AndAlso e.Y >= 637 AndAlso e.Y <= 758 Then
            ValorEsquerda = 896
            ValorTopo = 472
        End If


    End Sub

    Private Sub IdentificaNomeMenu(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles X6ToolStripMenuItem.Click, X7MToolStripMenuItem.Click, X7M9ToolStripMenuItem.Click, X7M911ToolStripMenuItem.Click, X7M6ToolStripMenuItem.Click, X7M5ToolStripMenuItem.Click, X7M11ToolStripMenuItem.Click, X69ToolStripMenuItem.Click, XToolStripMenuItem.Click, Xadd9ToolStripMenuItem.Click, X5ToolStripMenuItem.Click, ToolStripMenuItem58.Click, ToolStripMenuItem119.Click, ToolStripMenuItem118.Click, ToolStripMenuItem117.Click, ToolStripMenuItem116.Click, ToolStripMenuItem115.Click, ToolStripMenuItem114.Click, ToolStripMenuItem113.Click, ToolStripMenuItem112.Click, ToolStripMenuItem111.Click, ToolStripMenuItem110.Click, ToolStripMenuItem109.Click, ToolStripMenuItem108.Click, ToolStripMenuItem31.Click, ToolStripMenuItem30.Click, ToolStripMenuItem29.Click, ToolStripMenuItem28.Click, MenorCom7ªNoBaixoToolStripMenuItem.Click, TríadeMenorCom5ªNoBaixoToolStripMenuItem.Click, TríadeMenorCom3ªNoBaixoToolStripMenuItem.Click, TríadeMaiorCom5ªNoBaixoToolStripMenuItem.Click, TríadeMaiorCom3ªNoBaixoToolStripMenuItem.Click, SétimaE4ªToolStripMenuItem.Click, SétimaE4ªCom9ªToolStripMenuItem.Click, SétimaE4ªCom9ªMenorToolStripMenuItem.Click, SétimaE4ªCom9ªE13ªToolStripMenuItem.Click, SétimaE4ªCom13ªToolStripMenuItem.Click, SétimaCom9ªToolStripMenuItem.Click, SétimaCom9ªMenorToolStripMenuItem.Click, SétimaCom9ªMenorE13ªToolStripMenuItem.Click, SétimaCom9ªE13ªToolStripMenuItem.Click, SétimaCom9ªE11ªAumentadaToolStripMenuItem.Click, SétimaCom9ªAumentadaToolStripMenuItem.Click, SétimaCom9ªAumentadaE11ªAumentadaOuSétimaCom5ªDiminutaE9ªAumentadaToolStripMenuItem.Click, SétimaCom5ªDiminutaOu11ªAumentadaToolStripMenuItem.Click, SétimaCom5ªDiminutaE9ªMenorOuSétimaCom9ªMenorE11ªAumentadaToolStripMenuItem.Click, SétimaCom5ªAumentadaOu13ªMenorToolStripMenuItem.Click, SétimaCom5ªAumentadaE9ªToolStripMenuItem.Click, SétimaCom5ªAumentadaE9ªMenorOuSétimaCom9ªMenorOu13ªMenorToolStripMenuItem.Click, SétimaCom5ªAumentadaE9ªAumentadaToolStripMenuItem.Click, SétimaCom13ªToolStripMenuItem.Click, SétimaCom11ªAumentadaE13ªToolStripMenuItem.Click, ªDiminutaToolStripMenuItem.Click, TríadesMenoresTônicaNoBaixoToolStripMenuItem.Click, TríadesMenores4CordasComInversõesToolStripMenuItem.Click, TríadesMenores3CordasComInversõesToolStripMenuItem.Click, TríadesMaioresTônicaNoBaixoToolStripMenuItem.Click, TríadesMaiores4CordasComInversõesToolStripMenuItem.Click, TríadesMaiores3CordasComInversõesToolStripMenuItem.Click, TríadesComQuartaTônicaNoBaixoToolStripMenuItem.Click, TríadesComQuarta4CordasComInversõesToolStripMenuItem.Click, TríadesComQuarta3CordasComInversõesToolStripMenuItem.Click, TríadesAumentadasTônicaNoBaixoToolStripMenuItem.Click, TríadesAumentadas4CordasComInversõesToolStripMenuItem.Click, TétradesMenoresComSextaCEbGAToolStripMenuItem.Click, TétradesMenoresComSétimaMaiorACEGToolStripMenuItem.Click, TétradesMenoresComSétimaEQuintaDiminutaACEbGToolStripMenuItem.Click, TétradesMenoresComSétimaACEGToolStripMenuItem.Click, TétradesComSextaCEGAToolStripMenuItem.Click, TétradesComSétimaMaiorEQuintaDiminutaCEGbBToolStripMenuItem.Click, TétradesComSétimaMaiorEQuintaAumentadaCEGBToolStripMenuItem.Click, TétradesComSétimaMaiorCEGBToolStripMenuItem.Click, TétradesComSétimaEQuintaDiminutaCEGbBbToolStripMenuItem.Click, TétradesComSétimaEQuintaAumentadaCEGBbToolStripMenuItem.Click, TétradesComSétimaEQuartaCFGBbToolStripMenuItem.Click, TétradesComSétimaCEGBbToolStripMenuItem.Click
        Dim NomeMenu As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
        FamiliaAcorde = NomeMenu.Text
        NumeraçãoFamiliaAcorde = CInt(NomeMenu.Tag)
        RedefiniçãoTonalidade()
    End Sub

    Public Sub GerarAcordes()

        Try

            Array.Clear(ArrayAcordePadrão, 0, ArrayAcordePadrão.Length)
            Array.Clear(ImagemAcorde, 0, ImagemAcorde.Length)
            Array.Clear(LinhaFinalDaSeta, 0, LinhaFinalDaSeta.Length)

            'melhor definir isso para todos os campos da linha 6 do array logo no começo, e daí nos acordes, limpa-se os campos que não terão valor. Isso gera menos código
            For i = 1 To 48
                ArrayAcordePadrão(i, 7, 1) = "B" : ArrayAcordePadrão(i, 7, 2) = "B" : ArrayAcordePadrão(i, 7, 3) = "B" : ArrayAcordePadrão(i, 7, 4) = "B" : ArrayAcordePadrão(i, 7, 5) = "B" : ArrayAcordePadrão(i, 7, 6) = "B"
                ArrayAcordePadrão(i, 1, 0) = "Traste1"
            Next


            Array.Copy(IntervalosAcorde2, IntervalosAcorde, IntervalosAcorde.Length)

            Index = 0
            Array.Clear(AjustePosição, 0, AjustePosição.Length)

            'Nos acordes montados abaixo, se existir algum que não seja da tonalidade de C, efetuar os seguintes ajustes:
            'AjustePosição(1) = -1 ===> Tonalidade C#/Db
            'AjustePosição(1) = 10 ===> Tonalidade D
            'AjustePosição(1) = 9 ===> Tonalidade D#/Eb
            'AjustePosição(1) = 8 ===> Tonalidade E
            'AjustePosição(1) = 7 ===> Tonalidade F
            'AjustePosição(1) = 6 ===> Tonalidade F#/Gb
            'AjustePosição(1) = 5 ===> Tonalidade G
            'AjustePosição(1) = 4 ===> Tonalidade G#/Ab
            'AjustePosição(1) = 3 ===> Tonalidade A
            'AjustePosição(1) = 2 ===> Tonalidade A#/Bb
            'AjustePosição(1) = 1 ===> Tonalidade B

            BaixoDoAcordeTerça = NotasAcorde(NotasAcordeIndiceLinha, 6)
            BaixoDoAcordeTerçaMenor = NotasAcorde(NotasAcordeIndiceLinha, 5)
            BaixoDoAcordeQuarta = NotasAcorde(NotasAcordeIndiceLinha, 7)
            BaixoDoAcordeQuinta = NotasAcorde(NotasAcordeIndiceLinha, 10)
            BaixoDoAcordeQuintaDiminuida = NotasAcorde(NotasAcordeIndiceLinha, 9)
            BaixoDoAcordeQuintaAumentada = NotasAcorde(NotasAcordeIndiceLinha, 11)
            BaixoDoAcordeSexta = NotasAcorde(NotasAcordeIndiceLinha, 13)
            BaixoDoAcordeSétima = NotasAcorde(NotasAcordeIndiceLinha, 14)
            BaixoDoAcordeSétimaMaior = NotasAcorde(NotasAcordeIndiceLinha, 15)

            TétradesComSextaCEGAToolStripMenuItem.Text = "Tétrades com sexta – " & TonalidadeAcorde & " " & BaixoDoAcordeTerça & " " & BaixoDoAcordeQuinta & " " & BaixoDoAcordeSexta
            TétradesMenoresComSextaCEbGAToolStripMenuItem.Text = "Tétrades menores com sexta – " & TonalidadeAcorde & " " & BaixoDoAcordeTerçaMenor & " " & BaixoDoAcordeQuinta & " " & BaixoDoAcordeSexta
            TétradesComSétimaCEGBbToolStripMenuItem.Text = "Tétrades com sétima – " & TonalidadeAcorde & " " & BaixoDoAcordeTerça & " " & BaixoDoAcordeQuinta & " " & BaixoDoAcordeSétima
            TétradesComSétimaEQuartaCFGBbToolStripMenuItem.Text = "Tétrades com sétima e quarta – " & TonalidadeAcorde & " " & BaixoDoAcordeQuarta & " " & BaixoDoAcordeQuinta & " " & BaixoDoAcordeSétima
            TétradesComSétimaEQuintaDiminutaCEGbBbToolStripMenuItem.Text = "Tétrades com sétima e quinta diminuta – " & TonalidadeAcorde & " " & BaixoDoAcordeTerça & " " & BaixoDoAcordeQuintaDiminuida & " " & BaixoDoAcordeSétima
            TétradesComSétimaEQuintaAumentadaCEGBbToolStripMenuItem.Text = "Tétrades com sétima e quinta aumentada – " & TonalidadeAcorde & " " & BaixoDoAcordeTerça & " " & BaixoDoAcordeQuintaAumentada & " " & BaixoDoAcordeSétima
            TétradesComSétimaMaiorCEGBToolStripMenuItem.Text = "Tétrades com sétima maior – " & TonalidadeAcorde & " " & BaixoDoAcordeTerça & " " & BaixoDoAcordeQuinta & " " & BaixoDoAcordeSétimaMaior
            TétradesComSétimaMaiorEQuintaDiminutaCEGbBToolStripMenuItem.Text = "Tétrades com sétima maior e quinta diminuta – " & TonalidadeAcorde & " " & BaixoDoAcordeTerça & " " & BaixoDoAcordeQuintaDiminuida & " " & BaixoDoAcordeSétimaMaior
            TétradesComSétimaMaiorEQuintaAumentadaCEGBToolStripMenuItem.Text = "Tétrades com sétima maior e quinta aumentada – " & TonalidadeAcorde & " " & BaixoDoAcordeTerça & " " & BaixoDoAcordeQuintaAumentada & " " & BaixoDoAcordeSétimaMaior
            TétradesMenoresComSétimaMaiorACEGToolStripMenuItem.Text = "Tétrades menores com sétima maior – " & TonalidadeAcorde & " " & BaixoDoAcordeTerçaMenor & " " & BaixoDoAcordeQuinta & " " & BaixoDoAcordeSétimaMaior
            TétradesMenoresComSétimaACEGToolStripMenuItem.Text = "Tétrades menores com sétima – " & TonalidadeAcorde & " " & BaixoDoAcordeTerçaMenor & " " & BaixoDoAcordeQuinta & " " & BaixoDoAcordeSétima
            TétradesMenoresComSétimaEQuintaDiminutaACEbGToolStripMenuItem.Text = "Tétrades menores com sétima e quinta diminuta – " & TonalidadeAcorde & " " & BaixoDoAcordeTerçaMenor & " " & BaixoDoAcordeQuintaDiminuida & " " & BaixoDoAcordeSétima

            If NumeraçãoFamiliaAcorde = 66 Then
                FamiliaAcorde = TétradesComSextaCEGAToolStripMenuItem.Text
            ElseIf NumeraçãoFamiliaAcorde = 67 Then
                FamiliaAcorde = TétradesMenoresComSextaCEbGAToolStripMenuItem.Text
            ElseIf NumeraçãoFamiliaAcorde = 68 Then
                FamiliaAcorde = TétradesComSétimaCEGBbToolStripMenuItem.Text
            ElseIf NumeraçãoFamiliaAcorde = 69 Then
                FamiliaAcorde = TétradesComSétimaEQuartaCFGBbToolStripMenuItem.Text
            ElseIf NumeraçãoFamiliaAcorde = 70 Then
                FamiliaAcorde = TétradesComSétimaEQuintaDiminutaCEGbBbToolStripMenuItem.Text
            ElseIf NumeraçãoFamiliaAcorde = 71 Then
                FamiliaAcorde = TétradesComSétimaEQuintaAumentadaCEGBbToolStripMenuItem.Text
            ElseIf NumeraçãoFamiliaAcorde = 72 Then
                FamiliaAcorde = TétradesComSétimaMaiorCEGBToolStripMenuItem.Text
            ElseIf NumeraçãoFamiliaAcorde = 73 Then
                FamiliaAcorde = TétradesComSétimaMaiorEQuintaDiminutaCEGbBToolStripMenuItem.Text
            ElseIf NumeraçãoFamiliaAcorde = 74 Then
                FamiliaAcorde = TétradesComSétimaMaiorEQuintaAumentadaCEGBToolStripMenuItem.Text
            ElseIf NumeraçãoFamiliaAcorde = 75 Then
                FamiliaAcorde = TétradesMenoresComSétimaMaiorACEGToolStripMenuItem.Text
            ElseIf NumeraçãoFamiliaAcorde = 76 Then
                FamiliaAcorde = TétradesMenoresComSétimaACEGToolStripMenuItem.Text
            ElseIf NumeraçãoFamiliaAcorde = 77 Then
                FamiliaAcorde = TétradesMenoresComSétimaEQuintaDiminutaACEbGToolStripMenuItem.Text
            End If


            'a ordem dos acordes está aqui colocada um pouco diferente do que está no livro do Chediak
            'o objetivo é que esteja num ordenamento mais claro, que facilite a memorização 
            If FamiliaAcorde = "Tríade Maior" OrElse NumeraçãoFamiliaAcorde = 1 Then 'página 53 do Dicionário de Acordes Cifrados 

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 0, 4) = "P" : LinhaFinalDaSeta(Index) = 6 'a pestana foi colocada na linha 0, pois como o diagrama será deslocado 15 pixeis para baixo, o resultado é que a pestana será exibida na linha 1
                ArrayAcordePadrão(Index, 1, 5) = "1"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*" 'indica se este acorde está entre os mais usados 
                'ArrayAcordePadrão(Index,8, 2) = "T" : ArrayAcordePadrão(Index,8, 3) = "3" : ArrayAcordePadrão(Index,8, 4) = "5" : ArrayAcordePadrão(Index,8, 5) = "T" : ArrayAcordePadrão(Index,8, 6) = "3"
                ImagemAcorde(Index) = TonalidadeAcorde
                'If TonalidadeAcorde = "B" Then ArrayAcordePadrão(Index, 8, 6) = "11" 'específica o traste. Ajuste manual necessário em alguns casos

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 5, 3) = "2"
                ArrayAcordePadrão(Index, 5, 4) = "3"
                ArrayAcordePadrão(Index, 5, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 3, 2) = "1"
                    ArrayAcordePadrão(Index, 4, 1) = "2T"
                    ArrayAcordePadrão(Index, 4, 5) = "3"
                    ArrayAcordePadrão(Index, 4, 6) = "4"
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                Else
                    ArrayAcordePadrão(Index, 3, 2) = "2"
                    ArrayAcordePadrão(Index, 4, 1) = "3T"
                    ArrayAcordePadrão(Index, 4, 6) = "4"
                End If
                ArrayAcordePadrão(Index, 1, 1) = "P" : LinhaFinalDaSeta(Index) = 6
                ImagemAcorde(Index) = TonalidadeAcorde
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-3" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 5
                ArrayAcordePadrão(Index, 4, 1) = "4T"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 0, 4) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 1, 5) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ImagemAcorde(Index) = TonalidadeAcorde

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 3, 3) = "4"
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde
                ArrayAcordePadrão(Index, 1, 0) = "Traste8" 'indica a casa inicial do acorde 

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 3, 4) = "2"
                ArrayAcordePadrão(Index, 3, 6) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Tríade Maior com 3ª no baixo" OrElse NumeraçãoFamiliaAcorde = 2 Then 'página 54 do Dicionário de Acordes Cifrados 

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 3, 6) = "3"
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                Else
                    ArrayAcordePadrão(Index, 3, 6) = "4"
                End If
                ArrayAcordePadrão(Index, 0, 4) = "0"
                ArrayAcordePadrão(Index, 1, 5) = "1T"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeTerça

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 3, 6) = "3"
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                Else
                    ArrayAcordePadrão(Index, 3, 6) = "4"
                End If
                ArrayAcordePadrão(Index, 0, 1) = "P" : LinhaFinalDaSeta(Index) = 6 'a pestana foi colocada na linha 0, pois como o diagrama será deslocado 15 pixeis para baixo, o resultado é que a pestana será exibida na linha 1
                ArrayAcordePadrão(Index, 1, 5) = "1T"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeTerça

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 4, 5) = "4"
                    ArrayAcordePadrão(Index, 4, 6) = "5T"
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                Else
                    ArrayAcordePadrão(Index, 4, 5) = "3"
                    ArrayAcordePadrão(Index, 4, 6) = "4T"
                End If
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "2"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 4) = "T"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 3, 1) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4T"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "1"
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 3) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Tríade Maior com 5ª no baixo" OrElse NumeraçãoFamiliaAcorde = 3 Then 'página 55 do Dicionário de Acordes Cifrados 

                Index += 1
                If TonalidadeAcorde = "C" Then
                    'só para C, pois nos demais tons este acorde é impossível de ser feito
                    ArrayAcordePadrão(Index, 0, 4) = "P" : LinhaFinalDaSeta(Index) = 6 'a pestana foi colocada na linha 0, pois como o diagrama será deslocado 15 pixeis para baixo, o resultado é que a pestana será exibida na linha 1
                    ArrayAcordePadrão(Index, 1, 5) = "1"
                    ArrayAcordePadrão(Index, 2, 3) = "2"
                    ArrayAcordePadrão(Index, 3, 1) = "3"
                    ArrayAcordePadrão(Index, 3, 2) = "4T"
                    ArrayAcordePadrão(Index, 8, 1) = "*"
                    ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                    ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeQuinta
                Else
                    ImagemAcorde(Index) = " "
                End If

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 1) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 3, 2) = "T"
                ArrayAcordePadrão(Index, 5, 3) = "2"
               ArrayAcordePadrão(Index, 5, 4) = "3"
                ArrayAcordePadrão(Index, 5, 5) = "4"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeQuinta

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 6) = "1"
                ArrayAcordePadrão(Index, 5, 3) = "2"
                ArrayAcordePadrão(Index, 5, 4) = "3T"
                ArrayAcordePadrão(Index, 5, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeQuinta

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 6) = "T"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 3, 3) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "1"
                ArrayAcordePadrão(Index, 3, 4) = "2"
                ArrayAcordePadrão(Index, 3, 6) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 2, 2) = "2"
                ArrayAcordePadrão(Index, 2, 3) = "3"
                ArrayAcordePadrão(Index, 5, 5) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste9"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "9ª adicional" OrElse NumeraçãoFamiliaAcorde = 4 Then 'página 56 do Dicionário de Acordes Cifrados

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 0, 2) = "P" : LinhaFinalDaSeta(Index) = 6 'a pestana foi colocada na linha 0, pois como o diagrama será deslocado 15 pixeis para baixo, o resultado é que a pestana será exibida na linha 1
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 3, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "(add9)"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 3) = "2"
                ArrayAcordePadrão(Index, 3, 5) = "3"
               ArrayAcordePadrão(Index, 5, 4) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "(add9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste3"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 2) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 1) = "4T"
                ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "(add9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 5, 3) = "4"
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "(add9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 5) = "1"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3T"
                ArrayAcordePadrão(Index, 3, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "(add9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 0, 1) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 1, 5) = "1T"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "(add9)/" & BaixoDoAcordeTerça

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Tríade Aumentada" OrElse NumeraçãoFamiliaAcorde = 5 Then 'página 57 do Dicionário de Acordes Cifrados

                StringSubstituição(0) = "b6" : StringSubstituição(1) = "#5" : SubstituirString()

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    If TonalidadeAcorde <> "B" Then
                        ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    Else
                        ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                        ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                    End If
                End If
                ArrayAcordePadrão(Index, 1, 4) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "(#5)"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 4) = "2"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 4, 3) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "(#5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste3"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 4, 1) = "4T"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "(#5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 0, 6) = "0"
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 1, 5) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "(#5)"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 2, 5) = "3"
                ArrayAcordePadrão(Index, 3, 3) = "4"
                ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "(#5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3"
                ArrayAcordePadrão(Index, 4, 2) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "(#5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 6) = "2"
                ArrayAcordePadrão(Index, 4, 4) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "(#5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "com 6ª" OrElse NumeraçãoFamiliaAcorde = 6 Then 'página 57 do Dicionário de Acordes Cifrados

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 4) = "1S"
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 3, 6) = "3"
                ArrayAcordePadrão(Index, 5, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "6"

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 4) = "1S"
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 5, 3) = "3"
                ArrayAcordePadrão(Index, 5, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "6"

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 1, 5) = "1"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 2, 4) = "3S"
                ArrayAcordePadrão(Index, 3, 2) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "6 (omit5)"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-3" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 5, 3) = "4"
                ArrayAcordePadrão(Index, 5, 4) = "4"
                ArrayAcordePadrão(Index, 5, 5) = "4"
                ArrayAcordePadrão(Index, 5, 6) = "4S"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "6"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-3" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 4) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 5, 5) = "4"
                ArrayAcordePadrão(Index, 5, 6) = "4S"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "6 (omit5)"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 4) = "2"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 5, 3) = "4S"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "6"
                ArrayAcordePadrão(Index, 1, 0) = "Traste3"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "1T"
                ArrayAcordePadrão(Index, 3, 4) = "2"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 5, 3) = "4S"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "6 (omit5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste3"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1S"
                ArrayAcordePadrão(Index, 2, 1) = "2T"
                ArrayAcordePadrão(Index, 2, 5) = "3"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "6"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1S"
                ArrayAcordePadrão(Index, 2, 1) = "2T"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 2) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "6"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 5) = "4S"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "6 (omit5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

               Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1S"
                ArrayAcordePadrão(Index, 2, 1) = "2T"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "6 (omit5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                   ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 6) = "1S"
                ArrayAcordePadrão(Index, 4, 1) = "2T"
                ArrayAcordePadrão(Index, 4, 5) = "3"
                ArrayAcordePadrão(Index, 5, 4) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "6"
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "PT" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 5) = "S"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 3, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "6"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 6) = "1"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3T"
                ArrayAcordePadrão(Index, 3, 5) = "4S"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "6 (omit5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 2, 3) = "2T"
                ArrayAcordePadrão(Index, 2, 5) = "3S"
                ArrayAcordePadrão(Index, 4, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "6 (omit5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste9"

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 3) = "P" : LinhaFinalDaSeta(Index) = 4 : ArrayAcordePadrão(Index, 2, 4) = "S"
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 5, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "6 (omit5)"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 3 : ArrayAcordePadrão(Index, 1, 3) = "S"
                ArrayAcordePadrão(Index, 2, 1) = "2T"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "6 (omit5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "1T"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3"
                ArrayAcordePadrão(Index, 5, 2) = "4S"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "6 (omit5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 3) = "1"
                ArrayAcordePadrão(Index, 2, 4) = "2S"
                ArrayAcordePadrão(Index, 3, 1) = "3"
                ArrayAcordePadrão(Index, 3, 2) = "4T"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7/" & BaixoDoAcordeQuinta 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "6/" & BaixoDoAcordeQuinta

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 4) = "T" : ArrayAcordePadrão(Index, 1, 6) = "S"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7/" & BaixoDoAcordeQuinta 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "6/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 2, 2) = "2"
                ArrayAcordePadrão(Index, 2, 3) = "3T"
                ArrayAcordePadrão(Index, 2, 5) = "4S"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7/" & BaixoDoAcordeQuinta 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "6/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste9"

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 5) = "1T"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 2, 4) = "3S"
                ArrayAcordePadrão(Index, 3, 1) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7/" & BaixoDoAcordeQuinta 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "6/" & BaixoDoAcordeQuinta

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "1"
                ArrayAcordePadrão(Index, 3, 4) = "2T"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 3, 6) = "4S"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7/" & BaixoDoAcordeQuinta 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "6/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste3"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 6) = "1T"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 3, 5) = "4S"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7/" & BaixoDoAcordeQuinta 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "6/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "com 6ª e 9ª" OrElse NumeraçãoFamiliaAcorde = 7 Then 'página 59 do Dicionário de Acordes Cifrados

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 3) = "1"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 3, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "6(9)"

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 3) = "P" : LinhaFinalDaSeta(Index) = 5
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 3, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "6(9)"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "P" : LinhaFinalDaSeta(Index) = 4
                ArrayAcordePadrão(Index, 2, 1) = "2T"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "6(9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 1) = "2T"
                ArrayAcordePadrão(Index, 2, 5) = "3"
                ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "6(9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 1) = "2T"
                ArrayAcordePadrão(Index, 2, 5) = "3"
                ArrayAcordePadrão(Index, 2, 6) = "4"
                ImagemAcorde(Index) = TonalidadeAcorde & "6(9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 2, 3) = "2T"
                ArrayAcordePadrão(Index, 2, 5) = "3"
                ArrayAcordePadrão(Index, 2, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "6(9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste9"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 2, 6) = "3T"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "6(9)/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 3) = "P" : LinhaFinalDaSeta(Index) = 5
                ArrayAcordePadrão(Index, 3, 1) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 3, 5) = "4"
                ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "6(9)/" & BaixoDoAcordeQuinta

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 2, 2) = "2"
                ArrayAcordePadrão(Index, 2, 3) = "2T"
                ArrayAcordePadrão(Index, 2, 5) = "3"
                ArrayAcordePadrão(Index, 2, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "6(9)/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste9"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "com 7ª Maior" OrElse NumeraçãoFamiliaAcorde = 8 Then 'página 57 do Dicionário de Acordes Cifrados

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 0, 4) = "P" : LinhaFinalDaSeta(Index) = 5 'a pestana foi colocada na linha 0, pois como o diagrama será deslocado 15 pixeis para baixo, o resultado é que a pestana será exibida na linha 1
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7M"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 4, 4) = "2"
                ArrayAcordePadrão(Index, 5, 3) = "3"
                ArrayAcordePadrão(Index, 5, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7M"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "1T"
                ArrayAcordePadrão(Index, 3, 4) = "2"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 5, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7M"
                ArrayAcordePadrão(Index, 1, 0) = "Traste3"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 2, 4) = "3"
                ArrayAcordePadrão(Index, 3, 2) = "4"
                ImagemAcorde(Index) = TonalidadeAcorde & "7M"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "1T"
                ArrayAcordePadrão(Index, 1, 5) = "2"
                ArrayAcordePadrão(Index, 2, 3) = "3"
                ArrayAcordePadrão(Index, 2, 4) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7M"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 2) = "1T"
                ArrayAcordePadrão(Index, 3, 6) = "2"
                ArrayAcordePadrão(Index, 4, 4) = "3"
                ArrayAcordePadrão(Index, 5, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7M"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 6) = "1"
                ArrayAcordePadrão(Index, 2, 1) = "2T"
                ArrayAcordePadrão(Index, 2, 5) = "3"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7M"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3"
                ArrayAcordePadrão(Index, 5, 5) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7M"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 3, 4) = "2"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 3, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7M"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 6) = "1"
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 3) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7M"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 5) = "1T"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 3, 6) = "3"
                ArrayAcordePadrão(Index, 4, 4) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M/" & BaixoDoAcordeTerça

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "1T"
                ArrayAcordePadrão(Index, 3, 2) = "2"
                ArrayAcordePadrão(Index, 3, 6) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "1T"
                ArrayAcordePadrão(Index, 3, 2) = "2"
                ArrayAcordePadrão(Index, 4, 5) = "3"
                ArrayAcordePadrão(Index, 5, 3) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 3, 1) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 3, 5) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 3, 1) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 5, 2) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 0, 4) = "P" : LinhaFinalDaSeta(Index) = 6 'a pestana foi colocada na linha 0, pois como o diagrama será deslocado 15 pixeis para baixo, o resultado é que a pestana será exibida na linha 1
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 3, 1) = "3"
                ArrayAcordePadrão(Index, 3, 2) = "4T"
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M/" & BaixoDoAcordeQuinta

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
               End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 4) = "T"
                ArrayAcordePadrão(Index, 3, 6) = "3"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M/" & BaixoDoAcordeQuinta
               ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "com 7ª Maior e 5ª Aumentada" OrElse NumeraçãoFamiliaAcorde = 9 Then 'página 62 do Dicionário de Acordes Cifrados

                StringSubstituição(0) = "b6" : StringSubstituição(1) = "#5" : SubstituirString()

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 0, 5) = "P" : LinhaFinalDaSeta(Index) = 6 'a pestana foi colocada na linha 0, pois como o diagrama será deslocado 15 pixeis para baixo, o resultado é que a pestana será exibida na linha 1
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(#5)"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "1T"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 4, 3) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(#5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste3"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 2) = "1T"
                ArrayAcordePadrão(Index, 4, 4) = "2"
                ArrayAcordePadrão(Index, 4, 6) = "3"
                ArrayAcordePadrão(Index, 5, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(#5)"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "1T"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 2, 4) = "3"
                ArrayAcordePadrão(Index, 2, 5) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(#5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 3, 5) = "2"
                ArrayAcordePadrão(Index, 3, 6) = "3"
                ArrayAcordePadrão(Index, 4, 4) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(#5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 6) = "1"
                ArrayAcordePadrão(Index, 3, 4) = "2"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 4, 3) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(#5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 3, 1) = "2"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 4, 4) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(#5)/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "1T"
                ArrayAcordePadrão(Index, 3, 2) = "2"
                ArrayAcordePadrão(Index, 3, 6) = "3"
                ArrayAcordePadrão(Index, 5, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(#5)/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "com 7ª Maior e 6ª" OrElse NumeraçãoFamiliaAcorde = 10 Then 'página 63 do Dicionário de Acordes Cifrados

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 0, 5) = "P" : LinhaFinalDaSeta(Index) = 6 'a pestana foi colocada na linha 0, pois como o diagrama será deslocado 15 pixeis para baixo, o resultado é que a pestana será exibida na linha 1
                ArrayAcordePadrão(Index, 2, 3) = "1"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(6)"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "1T"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 5, 3) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(6)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste3"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 2) = "1T"
                ArrayAcordePadrão(Index, 4, 4) = "2"
                ArrayAcordePadrão(Index, 5, 5) = "3"
                ArrayAcordePadrão(Index, 5, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(6)"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "1T"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 2, 4) = "3"
                ArrayAcordePadrão(Index, 3, 5) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(6)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
               ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 3, 5) = "2"
                ArrayAcordePadrão(Index, 3, 6) = "3"
                ArrayAcordePadrão(Index, 5, 4) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(6)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 6) = "1"
                ArrayAcordePadrão(Index, 3, 4) = "2"
                ArrayAcordePadrão(Index, 4, 3) = "3T"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(6)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "com 7ª Maior e 9ª" OrElse NumeraçãoFamiliaAcorde = 11 Then 'página 63 do Dicionário de Acordes Cifrados

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 3) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 4, 4) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(9)"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "P" : LinhaFinalDaSeta(Index) = 4
                ArrayAcordePadrão(Index, 2, 1) = "2T"
                ArrayAcordePadrão(Index, 3, 3) = "3"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 6) = "3"
                ArrayAcordePadrão(Index, 5, 5) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 1) = "2T"
                ArrayAcordePadrão(Index, 2, 5) = "3"
                ArrayAcordePadrão(Index, 4, 3) = "4"
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 2, 4) = "3"
                ArrayAcordePadrão(Index, 3, 6) = "4"
                ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 2, 3) = "2T"
                ArrayAcordePadrão(Index, 2, 6) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste9"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 4, 3) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(9)/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 3) = "1"
                ArrayAcordePadrão(Index, 3, 1) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 4, 4) = "4"
                ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(9)/" & BaixoDoAcordeQuinta

                Index += 1
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 2, 2) = "2"
                ArrayAcordePadrão(Index, 2, 3) = "2T"
                ArrayAcordePadrão(Index, 2, 6) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(9)/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste9"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 4, 3) = "4T"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(9)/" & BaixoDoAcordeSétimaMaior
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 1) = "2T"
                ArrayAcordePadrão(Index, 2, 5) = "3"
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(6/9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "com 7ª Maior e 11ª Aumentada" OrElse NumeraçãoFamiliaAcorde = 12 Then 'página 65 do Dicionário de Acordes Cifrados

                StringSubstituição(0) = "b5" : StringSubstituição(1) = "#11" : SubstituirString()

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 0, 4) = "P" : LinhaFinalDaSeta(Index) = 5 'a pestana foi colocada na linha 0, pois como o diagrama será deslocado 15 pixeis para baixo, o resultado é que a pestana será exibida na linha 1
                ArrayAcordePadrão(Index, 2, 3) = "1"
                ArrayAcordePadrão(Index, 2, 6) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(#11)"

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 6) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 4, 4) = "3"
                ArrayAcordePadrão(Index, 5, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(#11)"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 2) = "1T"
                ArrayAcordePadrão(Index, 4, 3) = "2"
                ArrayAcordePadrão(Index, 4, 4) = "3"
                ArrayAcordePadrão(Index, 5, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(#11)"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 5) = "1"
                ArrayAcordePadrão(Index, 2, 1) = "2T"
                ArrayAcordePadrão(Index, 3, 3) = "3"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(#11)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "1T"
                ArrayAcordePadrão(Index, 2, 2) = "2"
                ArrayAcordePadrão(Index, 2, 3) = "3"
                ArrayAcordePadrão(Index, 2, 4) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(#11)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 1) = "2T"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(#11)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 1) = "2T"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 3) = "4"
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(#11)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 3) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(#11)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 4, 3) = "3"
                ArrayAcordePadrão(Index, 5, 2) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(#11)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste11"

                Index += 1
                If TonalidadeAcorde = "C" Then
                    ArrayAcordePadrão(Index, 0, 4) = "P" : LinhaFinalDaSeta(Index) = 5 'a pestana foi colocada na linha 0, pois como o diagrama será deslocado 15 pixeis para baixo, o resultado é que a pestana será exibida na linha 1
                    ArrayAcordePadrão(Index, 2, 3) = "1"
                    ArrayAcordePadrão(Index, 2, 6) = "2"
                    ArrayAcordePadrão(Index, 3, 1) = "3"
                    ArrayAcordePadrão(Index, 3, 2) = "4T"
                    ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                    ImagemAcorde(Index) = TonalidadeAcorde & "7M(#11)/" & BaixoDoAcordeQuinta
                Else
                    ImagemAcorde(Index) = " "
                End If

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "com 7ª Maior, 9ª e 11ª Aumentada" OrElse NumeraçãoFamiliaAcorde = 13 Then 'página 66 do Dicionário de Acordes Cifrados

                StringSubstituição(0) = "b5" : StringSubstituição(1) = "#11" : SubstituirString()

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 2) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 4, 4) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(9/#11)"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "P" : LinhaFinalDaSeta(Index) = 5
                ArrayAcordePadrão(Index, 2, 1) = "2T"
                ArrayAcordePadrão(Index, 3, 3) = "3"
                ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(9/#11)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 3) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 1) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 4, 4) = "4"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(9/#11)/" & BaixoDoAcordeQuinta

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Tríade Menor" OrElse NumeraçãoFamiliaAcorde = 14 Then 'página 67 do Dicionário de Acordes Cifrados

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                    ArrayAcordePadrão(Index, 3, 2) = "3T"
                Else
                    ArrayAcordePadrão(Index, 3, 2) = "4T"
                End If
                ArrayAcordePadrão(Index, 0, 4) = "0"
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 1, 5) = "2"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "m"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
               ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 4, 5) = "2"
                ArrayAcordePadrão(Index, 5, 3) = "3"
                ArrayAcordePadrão(Index, 5, 4) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "m"

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 3, 2) = "3T"
                Else
                    ArrayAcordePadrão(Index, 3, 2) = "4T"
                End If
                ArrayAcordePadrão(Index, 0, 4) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 3, 6) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "m"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 4, 1) = "2T"
                ArrayAcordePadrão(Index, 4, 4) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "m"
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 3, 3) = "4"
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "m"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 2, 6) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "m"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Tríade Menor com 3ª no baixo" OrElse NumeraçãoFamiliaAcorde = 15 Then 'página 68 do Dicionário de Acordes Cifrados

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                    ArrayAcordePadrão(Index, 3, 6) = "3"
                Else
                    ArrayAcordePadrão(Index, 3, 6) = "4"
                End If
                ArrayAcordePadrão(Index, 0, 4) = "0"
                ArrayAcordePadrão(Index, 1, 3) = "1S"
                ArrayAcordePadrão(Index, 1, 5) = "2T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6 (omit5)"  'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeTerçaMenor

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 2, 2) = "2S"
                ArrayAcordePadrão(Index, 4, 5) = "3"
                ArrayAcordePadrão(Index, 4, 6) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6 (omit5)"  'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeTerçaMenor
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 5 : ArrayAcordePadrão(Index, 1, 4) = "T"
                ArrayAcordePadrão(Index, 2, 2) = "2S"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6 (omit5)"  'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeTerçaMenor
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 2, 1) = "2S"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4T"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6 (omit5)"  'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeTerçaMenor
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                ArrayAcordePadrão(Index, 1, 6) = "1"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3S"
                ArrayAcordePadrão(Index, 3, 5) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6 (omit5)"  'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeTerçaMenor
                ArrayAcordePadrão(Index, 1, 0) = "Traste11"

                Index += 1
                ArrayAcordePadrão(Index, 1, 1) = "PS" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 5) = "3T"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6 (omit5)"  'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeTerçaMenor
                ArrayAcordePadrão(Index, 1, 0) = "Traste11"

                Index += 1
                ArrayAcordePadrão(Index, 1, 2) = "1S"
                ArrayAcordePadrão(Index, 3, 4) = "2"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 5, 3) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6 (omit5)"  'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeTerçaMenor
                ArrayAcordePadrão(Index, 1, 0) = "Traste6"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Tríade Menor com 5ª no baixo" OrElse NumeraçãoFamiliaAcorde = 16 Then 'página 69 do Dicionário de Acordes Cifrados

                Index += 1
                If TonalidadeAcorde = "C" Then
                    ArrayAcordePadrão(Index, 1, 3) = "1"
                    ArrayAcordePadrão(Index, 1, 5) = "2T"
                    ArrayAcordePadrão(Index, 3, 1) = "3"
                    ArrayAcordePadrão(Index, 3, 2) = "4"
                    ArrayAcordePadrão(Index, 7, 6) = ""
                    ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                    ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeQuinta
                Else
                    ImagemAcorde(Index) = " "
                End If

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 1) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 3, 2) = "T"
                ArrayAcordePadrão(Index, 4, 5) = "2"
                ArrayAcordePadrão(Index, 5, 3) = "3"
                ArrayAcordePadrão(Index, 5, 4) = "4"
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeQuinta

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou 
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 6) = "1"
                ArrayAcordePadrão(Index, 4, 5) = "2"
                ArrayAcordePadrão(Index, 5, 3) = "3"
                ArrayAcordePadrão(Index, 5, 4) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeQuinta

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 6) = "T"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 3, 3) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 3) = "T"
                ArrayAcordePadrão(Index, 2, 6) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Menor com 9ª Adicional" OrElse NumeraçãoFamiliaAcorde = 17 Then 'página 70 do Dicionário de Acordes Cifrados

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                    ArrayAcordePadrão(Index, 3, 2) = "2T"
                    ArrayAcordePadrão(Index, 3, 5) = "3"
                Else
                    ArrayAcordePadrão(Index, 3, 2) = "3T"
                    ArrayAcordePadrão(Index, 3, 5) = "4"
                End If
                ArrayAcordePadrão(Index, 0, 4) = "0"
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "m(add9)"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3"
                ArrayAcordePadrão(Index, 5, 4) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "m(add9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste3"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 2, 2) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 1) = "4T"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "m(add9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 5, 3) = "4"
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "m(add9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 3) = "3T"
                ArrayAcordePadrão(Index, 3, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "m(add9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Menor com 6ª" OrElse NumeraçãoFamiliaAcorde = 18 Then 'página 70 do Dicionário de Acordes Cifrados

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 4) = "1S"
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 3, 6) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7(b5)/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m6"

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "m6"

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 4) = "1S"
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 4, 5) = "3"
                ArrayAcordePadrão(Index, 5, 3) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7(b5)/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m6"

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 4) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 4, 5) = "3"
                ArrayAcordePadrão(Index, 5, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "m6"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 5, 3) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "m6"
                ArrayAcordePadrão(Index, 1, 0) = "Traste3"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1S"
                ArrayAcordePadrão(Index, 2, 1) = "2T"
                ArrayAcordePadrão(Index, 2, 4) = "3"
                ArrayAcordePadrão(Index, 2, 5) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7(b5)/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m6"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1S"
                ArrayAcordePadrão(Index, 2, 1) = "2T"
                ArrayAcordePadrão(Index, 2, 4) = "3"
                ArrayAcordePadrão(Index, 4, 2) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7(b5)/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m6"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 2, 1) = "2T"
                ArrayAcordePadrão(Index, 2, 4) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "m6"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "m6"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 6) = "1"
                ArrayAcordePadrão(Index, 4, 1) = "2T"
                ArrayAcordePadrão(Index, 4, 4) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "m6"
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "PT" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 5) = "S"
                ArrayAcordePadrão(Index, 2, 6) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7(b5)/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m6"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 3) = "3T"
                ArrayAcordePadrão(Index, 3, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "m6"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 5) = "T"
                ArrayAcordePadrão(Index, 2, 4) = "2S"
                ArrayAcordePadrão(Index, 3, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7(b5)/" & BaixoDoAcordeTerçaMenor 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m6/" & BaixoDoAcordeTerçaMenor

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 4) = "T"
                ArrayAcordePadrão(Index, 2, 2) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3S"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7(b5)/" & BaixoDoAcordeTerçaMenor 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m6/" & BaixoDoAcordeTerçaMenor
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 1) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3S"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7(b5)/" & BaixoDoAcordeTerçaMenor 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m6/" & BaixoDoAcordeTerçaMenor
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 4) = "1"
                ArrayAcordePadrão(Index, 3, 1) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 3, 6) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "m6/" & BaixoDoAcordeQuinta

                Index += 1
                ArrayAcordePadrão(Index, 1, 5) = "1"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 2, 4) = "3T"
                ArrayAcordePadrão(Index, 2, 6) = "4S"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7(b5)/" & BaixoDoAcordeQuinta 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m6/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste4"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3T"
                ArrayAcordePadrão(Index, 3, 5) = "4S"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7(b5)/" & BaixoDoAcordeQuinta 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m6/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 2, 4) = "2S"
                ArrayAcordePadrão(Index, 3, 1) = "3"
                ArrayAcordePadrão(Index, 3, 2) = "4T"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7(b5)/" & BaixoDoAcordeQuinta 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m6/" & BaixoDoAcordeQuinta

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1S"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 2, 5) = "3"
                ArrayAcordePadrão(Index, 2, 6) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7(b5)" 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m6/" & BaixoDoAcordeSexta
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 0, 2) = "1S"
                ArrayAcordePadrão(Index, 0, 4) = "2"
                ArrayAcordePadrão(Index, 1, 3) = "3"
                ArrayAcordePadrão(Index, 1, 5) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7(b5)" 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m6/" & BaixoDoAcordeSexta

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PS" : LinhaFinalDaSeta(Index) = 4 : ArrayAcordePadrão(Index, 1, 4) = "T"
                ArrayAcordePadrão(Index, 2, 2) = "2"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7(b5)" 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m6/" & BaixoDoAcordeSexta
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 6) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "2S"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7(b5)" 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m6/" & BaixoDoAcordeSexta
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                ArrayAcordePadrão(Index, 1, 5) = "1"
                ArrayAcordePadrão(Index, 2, 1) = "2S"
                ArrayAcordePadrão(Index, 2, 3) = "3"
                ArrayAcordePadrão(Index, 2, 4) = "4T"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7(b5)" 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m6/" & BaixoDoAcordeSexta
                ArrayAcordePadrão(Index, 1, 0) = "Traste4"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PS" : ArrayAcordePadrão(Index, 1, 4) = "T" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 2) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7(b5)" 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m6/" & BaixoDoAcordeSexta
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Menor com 6ª e 9ª" OrElse NumeraçãoFamiliaAcorde = 19 Then 'página 72 do Dicionário de Acordes Cifrados

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 3, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "m6(9)"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 2, 1) = "2T"
                ArrayAcordePadrão(Index, 2, 4) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 4, 6) = "4"
                ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "m6(9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 3, 3) = "2T"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 3, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "m6(9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 3) = "T"
                ArrayAcordePadrão(Index, 2, 1) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "m6(9)/" & BaixoDoAcordeTerçaMenor
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Menor com 7ª" OrElse NumeraçãoFamiliaAcorde = 20 Then 'página 73 do Dicionário de Acordes Cifrados

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "m7"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 4, 5) = "2S"
                ArrayAcordePadrão(Index, 5, 3) = "3"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 2) = "1T"
                ArrayAcordePadrão(Index, 3, 4) = "2"
                ArrayAcordePadrão(Index, 3, 6) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4S"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "m7"
                ArrayAcordePadrão(Index, 1, 0) = "Traste3"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 4) = "S"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "1T"
                ArrayAcordePadrão(Index, 1, 3) = "2"
                ArrayAcordePadrão(Index, 1, 4) = "3S"
                ArrayAcordePadrão(Index, 1, 5) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 4) = "S"
                ArrayAcordePadrão(Index, 3, 2) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 2, 6) = "3S"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                ArrayAcordePadrão(Index, 1, 5) = "1"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3"
                ArrayAcordePadrão(Index, 5, 2) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7"
                ArrayAcordePadrão(Index, 1, 0) = "Traste11"

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 3) = "S" : ArrayAcordePadrão(Index, 1, 5) = "T"
                ArrayAcordePadrão(Index, 3, 1) = "3"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6/" & BaixoDoAcordeQuinta 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7/" & BaixoDoAcordeQuinta

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 4) = "S" : ArrayAcordePadrão(Index, 1, 6) = "T"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6/" & BaixoDoAcordeQuinta 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Menor com 7ª no baixo" OrElse NumeraçãoFamiliaAcorde = 21 Then 'página 74 do Dicionário de Acordes Cifrados

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 5) = "2S"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 1) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6/" & BaixoDoAcordeSétima 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeSétima
                ArrayAcordePadrão(Index, 1, 0) = "Traste3"

                Index += 1
                ArrayAcordePadrão(Index, 1, 5) = "1S"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 2, 4) = "3T"
                ArrayAcordePadrão(Index, 3, 1) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6/" & BaixoDoAcordeSétima 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeSétima
                ArrayAcordePadrão(Index, 1, 0) = "Traste4"

                Index += 1
                ArrayAcordePadrão(Index, 1, 1) = "1"
                ArrayAcordePadrão(Index, 3, 4) = "2S"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 3, 6) = "4T"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6/" & BaixoDoAcordeSétima 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeSétima
                ArrayAcordePadrão(Index, 1, 0) = "Traste6"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 4) = "S" : ArrayAcordePadrão(Index, 1, 6) = "T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6/" & BaixoDoAcordeSétima 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeSétima
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                ArrayAcordePadrão(Index, 1, 6) = "1S"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 3, 5) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6/" & BaixoDoAcordeSétima 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeSétima
                ArrayAcordePadrão(Index, 1, 0) = "Traste11"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Menor com 7ª e 9ª" OrElse NumeraçãoFamiliaAcorde = 22 Then 'página 75 do Dicionário de Acordes Cifrados

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 3, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "m7(9)"

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 3, 5) = "4"
                ArrayAcordePadrão(Index, 3, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "m7(9)"

                Index += 1
                ArrayAcordePadrão(Index, 1, 2) = "1"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 1) = "3T"
                ArrayAcordePadrão(Index, 3, 3) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "m7(9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste6"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 2) = "2"
                ArrayAcordePadrão(Index, 3, 6) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "m7(9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 3, 6) = "4"
                ImagemAcorde(Index) = TonalidadeAcorde & "m7(9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 3, 3) = "2T"
                ArrayAcordePadrão(Index, 3, 6) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "m7(9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 3, 1) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 3, 5) = "4"
                ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "m7(9)/" & BaixoDoAcordeQuinta

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Menor com 7ª e 5ª Diminuta" OrElse NumeraçãoFamiliaAcorde = 23 Then 'página 76 do Dicionário de Acordes Cifrados

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 2) = "1T"
                ArrayAcordePadrão(Index, 3, 4) = "2"
                ArrayAcordePadrão(Index, 4, 3) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4S"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "m6/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7(b5)"

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 6) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4S"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "m6/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7(b5)"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 2, 5) = "2S"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "m6/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7(b5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste3"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 5) = "1"
                ArrayAcordePadrão(Index, 2, 1) = "2T"
                ArrayAcordePadrão(Index, 2, 3) = "3"
                ArrayAcordePadrão(Index, 2, 4) = "4S"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "m6/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7(b5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 4) = "S"
                ArrayAcordePadrão(Index, 2, 2) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "m6/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7(b5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "1T"
                ArrayAcordePadrão(Index, 1, 3) = "2"
                ArrayAcordePadrão(Index, 1, 4) = "3S"
                ArrayAcordePadrão(Index, 2, 2) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "m6/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7(b5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 2, 5) = "3"
                ArrayAcordePadrão(Index, 2, 6) = "4S"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "m6/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7(b5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Menor com 7ª e 11ª" OrElse NumeraçãoFamiliaAcorde = 24 Then 'página 77 do Dicionário de Acordes Cifrados

                StringSubstituição(0) = "4" : StringSubstituição(1) = "11" : SubstituirString()

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "m7(11)"

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 6) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "m7(11)"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 4, 5) = "2"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "m7(11)"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "m7(11)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste3"

                Index += 1
                ArrayAcordePadrão(Index, 1, 5) = "1"
                ArrayAcordePadrão(Index, 3, 1) = "2T"
                ArrayAcordePadrão(Index, 3, 3) = "3"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "m7(11)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste6"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 3) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "m7(11)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "1T"
                ArrayAcordePadrão(Index, 1, 2) = "2"
                ArrayAcordePadrão(Index, 1, 3) = "3"
                ArrayAcordePadrão(Index, 1, 4) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "m7(11)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 1, 4) = "2"
                ArrayAcordePadrão(Index, 2, 5) = "3"
                ArrayAcordePadrão(Index, 2, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "m7(11)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Menor com 7ª, 9ª e 11ª" OrElse NumeraçãoFamiliaAcorde = 25 Then 'página 77 do Dicionário de Acordes Cifrados

                StringSubstituição(0) = "4" : StringSubstituição(1) = "11" : SubstituirString()

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 3, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "m7(9/11)"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 1) = "3T"
                ArrayAcordePadrão(Index, 3, 3) = "4"
                ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "m7(9/11)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste6"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Menor com 7ª Maior" OrElse NumeraçãoFamiliaAcorde = 26 Then 'página 78 do Dicionário de Acordes Cifrados

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 0, 4) = "P" : LinhaFinalDaSeta(Index) = 6 'a pestana foi colocada na linha 0, pois como o diagrama será deslocado 15 pixeis para baixo, o resultado é que a pestana será exibida na linha 1
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "m(7M)"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 4, 4) = "2"
                ArrayAcordePadrão(Index, 4, 5) = "3"
                ArrayAcordePadrão(Index, 5, 3) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "m(7M)"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "1T"
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 5, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "m(7M)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste3"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 2) = "1T"
                ArrayAcordePadrão(Index, 3, 6) = "2"
                ArrayAcordePadrão(Index, 4, 4) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "m(7M)"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "m(7M)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 2, 6) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 3, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "m(7M)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 6) = "1"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 2, 5) = "3"
                ArrayAcordePadrão(Index, 4, 3) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "m(7M)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 6) = "1"
                ArrayAcordePadrão(Index, 2, 1) = "2T"
                ArrayAcordePadrão(Index, 2, 4) = "3"
                ArrayAcordePadrão(Index, 2, 5) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "m(7M)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "1T"
                ArrayAcordePadrão(Index, 2, 2) = "2"
                ArrayAcordePadrão(Index, 3, 6) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "m(7M)/" & BaixoDoAcordeTerçaMenor
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Menor com 7ª Maior e 6ª" OrElse NumeraçãoFamiliaAcorde = 27 Then 'página 79 do Dicionário de Acordes Cifrados

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 0, 5) = "0"
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "m(7M/6)"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 4, 4) = "2"
                ArrayAcordePadrão(Index, 4, 5) = "3"
                ArrayAcordePadrão(Index, 5, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "m(7M/6)"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 1) = "2T"
                ArrayAcordePadrão(Index, 2, 4) = "3"
                ArrayAcordePadrão(Index, 2, 5) = "4"
                ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "m(7M/6)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Menor com 7ª Maior e 9ª" OrElse NumeraçãoFamiliaAcorde = 28 Then 'página 79 do Dicionário de Acordes Cifrados

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 4, 4) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "m(7M/9)"

                Index += 1
                ArrayAcordePadrão(Index, 1, 2) = "1"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 1) = "3T"
                ArrayAcordePadrão(Index, 4, 3) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "m(7M/9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste6"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 3, 3) = "2T"
                ArrayAcordePadrão(Index, 3, 6) = "3"
                ArrayAcordePadrão(Index, 5, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "m(7M/9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 3, 6) = "4"
                ImagemAcorde(Index) = TonalidadeAcorde & "m(7M/9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 6) = "3"
                ArrayAcordePadrão(Index, 5, 5) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "m(7M/9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Tríade com 4ª" OrElse NumeraçãoFamiliaAcorde = 29 Then 'página 80 do Dicionário de Acordes Cifrados

                Index += 1
                If TonalidadeAcorde = "C" Then
                    ArrayAcordePadrão(Index, 1, 5) = "P" : LinhaFinalDaSeta(Index) = 6
                    ArrayAcordePadrão(Index, 3, 2) = "3T"
                    ArrayAcordePadrão(Index, 3, 3) = "4"
                    ArrayAcordePadrão(Index, 7, 1) = ""
                    ImagemAcorde(Index) = TonalidadeAcorde & "4"
                Else
                    ImagemAcorde(Index) = " "
                End If

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 3) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "4"
                ArrayAcordePadrão(Index, 1, 0) = "Traste3"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 2) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "4"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 3, 4) = "2"
                ArrayAcordePadrão(Index, 4, 5) = "3"
                ArrayAcordePadrão(Index, 4, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "4"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 5
                ArrayAcordePadrão(Index, 3, 3) = "2T"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "4/" & BaixoDoAcordeQuarta
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 5
                ArrayAcordePadrão(Index, 4, 5) = "3T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "4/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Sétima" OrElse NumeraçãoFamiliaAcorde = 30 Then 'página 81 do Dicionário de Acordes Cifrados

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 5) = "1"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 5, 3) = "3"
                ArrayAcordePadrão(Index, 5, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 4) = "2"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 4, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7"
                ArrayAcordePadrão(Index, 1, 0) = "Traste3"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 3) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "2"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 4, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7"
                ArrayAcordePadrão(Index, 1, 0) = "Traste3"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 6) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 4, 1) = "4T"
                ImagemAcorde(Index) = TonalidadeAcorde & "7"
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-3" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "1T"
                ArrayAcordePadrão(Index, 1, 3) = "2"
                ArrayAcordePadrão(Index, 1, 5) = "3"
                ArrayAcordePadrão(Index, 2, 4) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 3, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                ArrayAcordePadrão(Index, 1, 6) = "1"
                ArrayAcordePadrão(Index, 3, 5) = "2"
                ArrayAcordePadrão(Index, 4, 4) = "3"
                ArrayAcordePadrão(Index, 5, 3) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7"
                ArrayAcordePadrão(Index, 1, 0) = "Traste6"

                Index += 1
                ArrayAcordePadrão(Index, 1, 5) = "1"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 4, 3) = "3"
                ArrayAcordePadrão(Index, 5, 2) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7"
                ArrayAcordePadrão(Index, 1, 0) = "Traste11"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Sétima com 3ª no baixo" OrElse NumeraçãoFamiliaAcorde = 31 Then 'página 82 do Dicionário de Acordes Cifrados

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 5) = "1T"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 3, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7/" & BaixoDoAcordeTerça

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "1T"
                ArrayAcordePadrão(Index, 2, 6) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "1T"
                ArrayAcordePadrão(Index, 3, 2) = "2"
                ArrayAcordePadrão(Index, 4, 3) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 4) = "T"
                ArrayAcordePadrão(Index, 2, 6) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 5) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 3, 1) = "3"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Sétima com 5ª no baixo" OrElse NumeraçãoFamiliaAcorde = 32 Then 'página 83 do Dicionário de Acordes Cifrados

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 5) = "1T"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 3, 1) = "3"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7/" & BaixoDoAcordeQuinta

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "1"
                ArrayAcordePadrão(Index, 3, 4) = "2T"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 4, 6) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste3"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 4) = "T"
                ArrayAcordePadrão(Index, 2, 6) = "2"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 2, 2) = "2"
                ArrayAcordePadrão(Index, 2, 3) = "3T"
                ArrayAcordePadrão(Index, 3, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste9"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 6) = "1T"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Sétima com 7ª no baixo" OrElse NumeraçãoFamiliaAcorde = 33 Then 'página 84 do Dicionário de Acordes Cifrados

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 0, 4) = "1"
                ArrayAcordePadrão(Index, 1, 2) = "2"
                ArrayAcordePadrão(Index, 1, 5) = "3T"
                ArrayAcordePadrão(Index, 2, 3) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeSétima

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "2" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 0, 4) = "-1"
                ArrayAcordePadrão(Index, 0, 6) = "0"
                ArrayAcordePadrão(Index, 1, 2) = "1"
                ArrayAcordePadrão(Index, 1, 5) = "2T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeSétima

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 4) = "T"
                ArrayAcordePadrão(Index, 2, 1) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeSétima
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                ArrayAcordePadrão(Index, 1, 1) = "1"
                ArrayAcordePadrão(Index, 3, 5) = "2"
                ArrayAcordePadrão(Index, 3, 6) = "3T"
                ArrayAcordePadrão(Index, 4, 4) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeSétima
                ArrayAcordePadrão(Index, 1, 0) = "Traste6"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 4) = "T"
                ArrayAcordePadrão(Index, 2, 1) = "2"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeSétima
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 0, 4) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 1, 2) = "2"
                ArrayAcordePadrão(Index, 1, 5) = "3T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeSétima
          
                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 6) = "T"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeSétima
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Sétima com 5ª diminuta ou 11ª aumentada" OrElse NumeraçãoFamiliaAcorde = 34 Then 'página 85 do Dicionário de Acordes Cifrados

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 6) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 5, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(b5)"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "1T"
                ArrayAcordePadrão(Index, 1, 4) = "2"
                ArrayAcordePadrão(Index, 2, 3) = "3"
                ArrayAcordePadrão(Index, 3, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7(b5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste3"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 5) = "1"
                ArrayAcordePadrão(Index, 2, 1) = "2T"
                ArrayAcordePadrão(Index, 2, 3) = "3"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(b5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "1T"
                ArrayAcordePadrão(Index, 1, 3) = "2"
                ArrayAcordePadrão(Index, 2, 2) = "3"
                ArrayAcordePadrão(Index, 2, 4) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7(b5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 2, 5) = "3"
                ArrayAcordePadrão(Index, 3, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7(b5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 5) = "1T"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 2, 6) = "3"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7(b5)/" & BaixoDoAcordeTerça

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Sétima com 9ª e 11ª aumentada" OrElse NumeraçãoFamiliaAcorde = 35 Then 'página 86 do Dicionário de Acordes Cifrados

                StringSubstituição(0) = "b5" : StringSubstituição(1) = "#11" : SubstituirString()

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 3) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 3, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(9/#11)"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 1) = "2T"
                ArrayAcordePadrão(Index, 2, 3) = "3"
                ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7(9/#11)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Sétima com 11ª aumentada e 13ª" OrElse NumeraçãoFamiliaAcorde = 36 Then 'página 87 do Dicionário de Acordes Cifrados

                StringSubstituição(0) = "b5" : StringSubstituição(1) = "#11" : SubstituirString()
                StringSubstituição(0) = "6" : StringSubstituição(1) = "13" : SubstituirString()

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 6) = "1"
                ArrayAcordePadrão(Index, 3, 5) = "2"
                ArrayAcordePadrão(Index, 4, 1) = "3T"
                ArrayAcordePadrão(Index, 4, 3) = "3"
                ArrayAcordePadrão(Index, 5, 4) = "4"
                ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7(#11/13)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Sétima com 5ª aumentada ou 13ª menor" OrElse NumeraçãoFamiliaAcorde = 37 Then 'página 88 do Dicionário de Acordes Cifrados

                StringSubstituição(0) = "b6" : StringSubstituição(1) = "#5" : SubstituirString()

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 4, 6) = "2"
                ArrayAcordePadrão(Index, 5, 5) = "3"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(#5)"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "1T"
                ArrayAcordePadrão(Index, 1, 3) = "2"
                ArrayAcordePadrão(Index, 2, 4) = "3"
                ArrayAcordePadrão(Index, 2, 5) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(#5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 2, 5) = "3"
                ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7(#5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 5) = "1T"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7(#5)/" & BaixoDoAcordeTerça

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 0, 6) = "0"
                ArrayAcordePadrão(Index, 1, 2) = "1"
                ArrayAcordePadrão(Index, 1, 4) = "2"
                ArrayAcordePadrão(Index, 1, 5) = "3T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7(#5)/" & BaixoDoAcordeSétima

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "1"
                ArrayAcordePadrão(Index, 1, 4) = "2"
                ArrayAcordePadrão(Index, 1, 5) = "3T"
                ArrayAcordePadrão(Index, 2, 3) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7(#5)/" & BaixoDoAcordeSétima

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 4) = "T"
                ArrayAcordePadrão(Index, 2, 1) = "2"
                ArrayAcordePadrão(Index, 2, 3) = "3"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7(#5)/" & BaixoDoAcordeSétima
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Sétima com 5ª aumentada e 9ª" OrElse NumeraçãoFamiliaAcorde = 38 Then 'página 89 do Dicionário de Acordes Cifrados

                StringSubstituição(0) = "b6" : StringSubstituição(1) = "#5" : SubstituirString()

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 3) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 4, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7(#5/9)"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 2, 5) = "3"
                ArrayAcordePadrão(Index, 3, 6) = "4"
                ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(#5/9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 1) = "2T"
                ArrayAcordePadrão(Index, 2, 3) = "3"
                ArrayAcordePadrão(Index, 3, 5) = "4"
                ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7(#5/9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Sétima com 13ª" OrElse NumeraçãoFamiliaAcorde = 39 Then 'página 89 do Dicionário de Acordes Cifrados

                StringSubstituição(0) = "6" : StringSubstituição(1) = "13" : SubstituirString()

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 5, 5) = "3"
                ArrayAcordePadrão(Index, 5, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(13)"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 5, 3) = "2"
                ArrayAcordePadrão(Index, 5, 5) = "3"
                ArrayAcordePadrão(Index, 5, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7(13)"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "1T"
                ArrayAcordePadrão(Index, 1, 3) = "2"
                ArrayAcordePadrão(Index, 2, 4) = "3"
                ArrayAcordePadrão(Index, 3, 5) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(13)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 3, 5) = "4"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(13)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Sétima com 9ª" OrElse NumeraçãoFamiliaAcorde = 40 Then 'página 90 do Dicionário de Acordes Cifrados

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 3) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 3, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(9)"

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 3) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 3, 6) = "3"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7(9)"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 5) = "1"
                ArrayAcordePadrão(Index, 3, 4) = "2"
                ArrayAcordePadrão(Index, 4, 1) = "3T"
                ArrayAcordePadrão(Index, 4, 3) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 0, 6) = "0"
                ArrayAcordePadrão(Index, 3, 2) = "1T"
                ArrayAcordePadrão(Index, 3, 4) = "2"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(9)"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "1"
                ArrayAcordePadrão(Index, 1, 4) = "2"
                ArrayAcordePadrão(Index, 2, 1) = "3T"
                ArrayAcordePadrão(Index, 2, 3) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7(9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 3, 6) = "4"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 6) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7(9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 2, 3) = "2T"
                ArrayAcordePadrão(Index, 2, 6) = "3"
                ArrayAcordePadrão(Index, 3, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste9"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Sétima com 9ª e 13ª" OrElse NumeraçãoFamiliaAcorde = 41 Then 'página 91 do Dicionário de Acordes Cifrados

                StringSubstituição(0) = "6" : StringSubstituição(1) = "13" : SubstituirString()

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 3) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 5, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(9/13)"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 5) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 4) = "2"
                ArrayAcordePadrão(Index, 4, 1) = "3T"
                ArrayAcordePadrão(Index, 4, 3) = "4"
                ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(9/13)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 3, 6) = "4"
                ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(9/13)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Sétima com 9ª menor" OrElse NumeraçãoFamiliaAcorde = 42 Then 'página 91 do Dicionário de Acordes Cifrados

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 3) = "1"
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(b9)"

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 3) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 3, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7(b9)"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 5) = "1"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 4, 1) = "3T"
                ArrayAcordePadrão(Index, 4, 3) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(b9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 0, 6) = "0"
                ArrayAcordePadrão(Index, 2, 5) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(b9)"

                Index += 1
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 2, 2) = "2"
                ArrayAcordePadrão(Index, 3, 1) = "3T"
                ArrayAcordePadrão(Index, 3, 3) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7(b9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste6"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 2, 6) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(b9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 2, 6) = "3"
                ArrayAcordePadrão(Index, 3, 2) = "4"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(b9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 1, 6) = "2"
                ArrayAcordePadrão(Index, 2, 3) = "3T"
                ArrayAcordePadrão(Index, 3, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(b9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste9"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Sétima com 5ª diminuta e 9ª menor ou Sétima com 9ª menor e 11ª aumentada" OrElse NumeraçãoFamiliaAcorde = 43 Then 'página 92 do Dicionário de Acordes Cifrados

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 3) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7(b5/b9)"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Sétima com 5ª aumentada e 9ª menor ou  Sétima com 9ª menor ou 13ª menor" OrElse NumeraçãoFamiliaAcorde = 44 Then 'página 93 do Dicionário de Acordes Cifrados 

                StringSubstituição(0) = "b6" : StringSubstituição(1) = "#5" : SubstituirString()

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 3) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(#5/b9)"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 2, 5) = "3"
                ArrayAcordePadrão(Index, 2, 6) = "4"
                ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(#5/b9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Sétima com 9ª menor e 13ª" OrElse NumeraçãoFamiliaAcorde = 45 Then 'página 93 do Dicionário de Acordes Cifrados 

                StringSubstituição(0) = "6" : StringSubstituição(1) = "13" : SubstituirString()

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 3) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 5, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7(b9/13)"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 5) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 4, 1) = "3T"
                ArrayAcordePadrão(Index, 4, 3) = "4"
                ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(b9/13)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 2, 6) = "3"
                ArrayAcordePadrão(Index, 3, 5) = "4"
                ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(b9/13)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Sétima com 9ª aumentada" OrElse NumeraçãoFamiliaAcorde = 46 Then 'página 94 do Dicionário de Acordes Cifrados

                StringSubstituição(0) = "b3" : StringSubstituição(1) = "#9" : SubstituirString()

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 3) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(#9)"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "1"
                ArrayAcordePadrão(Index, 2, 1) = "2T"
                ArrayAcordePadrão(Index, 2, 3) = "3"
                ArrayAcordePadrão(Index, 2, 4) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7(#9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 4, 6) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(#9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 4, 6) = "4"
                ArrayAcordePadrão(Index, 7, 3) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7(#9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 2, 3) = "2T"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 3, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(#9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste9"

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 3) = "1"
                ArrayAcordePadrão(Index, 3, 1) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7(#9)/" & BaixoDoAcordeQuinta

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Sétima com 5ª aumentada e 9ª aumentada" OrElse NumeraçãoFamiliaAcorde = 47 Then 'página 95 do Dicionário de Acordes Cifrados

                StringSubstituição(0) = "b3" : StringSubstituição(1) = "#9" : SubstituirString()
                StringSubstituição(0) = "b6" : StringSubstituição(1) = "#5" : SubstituirString()

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 3) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 4, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(#5/#9)"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "1"
                ArrayAcordePadrão(Index, 2, 1) = "2T"
                ArrayAcordePadrão(Index, 2, 3) = "3"
                ArrayAcordePadrão(Index, 2, 4) = "3"
                ArrayAcordePadrão(Index, 3, 5) = "4"
                ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(#5/#9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 2, 5) = "3"
                ArrayAcordePadrão(Index, 4, 6) = "4"
                ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(#5/#9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Sétima com 9ª aumentada e 11ª aumentada ou  Sétima com 5ª diminuta e 9ª aumentada" OrElse NumeraçãoFamiliaAcorde = 48 Then 'página 95 do Dicionário de Acordes Cifrados

                StringSubstituição(0) = "b3" : StringSubstituição(1) = "#9" : SubstituirString()
                StringSubstituição(0) = "b5" : StringSubstituição(1) = "#11" : SubstituirString()

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 3) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(#9/#11)"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 1) = "2T"
                ArrayAcordePadrão(Index, 2, 3) = "3"
                ArrayAcordePadrão(Index, 2, 4) = "4"
                ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7(#9/#11)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Sétima e 4ª" OrElse NumeraçãoFamiliaAcorde = 49 Then 'página 96 do Dicionário de Acordes Cifrados

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 5) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 3, 3) = "3"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7(4)"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 3) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(4)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste3"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(4)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(4)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Sétima e 4ª com 9ª" OrElse NumeraçãoFamiliaAcorde = 50 Then 'página 96 do Dicionário de Acordes Cifrados

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 2) = "1T"
                ArrayAcordePadrão(Index, 3, 3) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 3, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(4/9)"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 7, 1) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(4/9)"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(4/9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste3"

                Index += 1
                ArrayAcordePadrão(Index, 1, 5) = "1"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 1) = "3T"
                ArrayAcordePadrão(Index, 3, 3) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(4/9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste6"

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 6) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 3, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(4/9)"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 2, 1) = "2T"
                ArrayAcordePadrão(Index, 2, 2) = "3"
                ArrayAcordePadrão(Index, 2, 3) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(4/9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(4/9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Sétima e 4ª com 13ª" OrElse NumeraçãoFamiliaAcorde = 51 Then 'página 97 do Dicionário de Acordes Cifrados

                StringSubstituição(0) = "6" : StringSubstituição(1) = "13" : SubstituirString()

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 3) = "2"
                ArrayAcordePadrão(Index, 3, 6) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7(4/13)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste3"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 2) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 3, 5) = "4"
                ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7(4/13)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Sétima e 4ª com 9ª e 13ª" OrElse NumeraçãoFamiliaAcorde = 52 Then 'página 97 do Dicionário de Acordes Cifrados

                StringSubstituição(0) = "6" : StringSubstituição(1) = "13" : SubstituirString()

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 5, 6) = "3"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(4/9/13)"

                'neste precisa desenhar o arco
                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 1, 6) = "1"
                ArrayAcordePadrão(Index, 2, 5) = "1"
                ArrayAcordePadrão(Index, 3, 4) = "2"
                ArrayAcordePadrão(Index, 4, 1) = "3T"
                ArrayAcordePadrão(Index, 4, 3) = "4"
                ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7(4/9/13)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Sétima e 4ª com 9ª menor" OrElse NumeraçãoFamiliaAcorde = 53 Then 'página 98 do Dicionário de Acordes Cifrados

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 5) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 3, 3) = "3"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(4/b9)"

                Index += 1
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 3, 1) = "2T"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 3, 3) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7(4/b9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste6"

                Index += 1
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 1) = "3T"
                ArrayAcordePadrão(Index, 3, 3) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(4/b9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste6"

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 6) = "1"
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "7(4/b9)"

                Index += 1
                ArrayAcordePadrão(Index, 1, 6) = "1"
                ArrayAcordePadrão(Index, 2, 3) = "2T"
                ArrayAcordePadrão(Index, 2, 4) = "3"
                ArrayAcordePadrão(Index, 3, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7(4/b9)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste9"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "7ª diminuta" OrElse NumeraçãoFamiliaAcorde = 54 Then 'página 99 do Dicionário de Acordes Cifrados

                StringSubstituição(0) = "6" : StringSubstituição(1) = "º" : SubstituirString()
                StringSubstituição(0) = "b6" : StringSubstituição(1) = "b13" : SubstituirString()

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 4) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 4, 3) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "º"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 2, 1) = "2T"
                ArrayAcordePadrão(Index, 2, 4) = "3"
                ArrayAcordePadrão(Index, 3, 2) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "º"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 1, 5) = "2"
                ArrayAcordePadrão(Index, 2, 1) = "3T"
                ArrayAcordePadrão(Index, 2, 4) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "º"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 2, 4) = "1"
                ArrayAcordePadrão(Index, 2, 6) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "º"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 1, 5) = "2"
                ArrayAcordePadrão(Index, 2, 4) = "3"
                ArrayAcordePadrão(Index, 2, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "º"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 4) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 4, 5) = "3"
                ArrayAcordePadrão(Index, 4, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 1) = "*"
                ImagemAcorde(Index) = TonalidadeAcorde & "º(b13)"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 2, 1) = "2T"
                ArrayAcordePadrão(Index, 2, 4) = "3"
                ArrayAcordePadrão(Index, 3, 5) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "º(b13)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 2) = "1T"
                ArrayAcordePadrão(Index, 4, 3) = "2"
                ArrayAcordePadrão(Index, 4, 4) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "º(7M)"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 1) = "2T"
                ArrayAcordePadrão(Index, 2, 4) = "3"
                ArrayAcordePadrão(Index, 3, 2) = "4"
                ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "º(7M)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Tríades maiores – Tônica no baixo" OrElse NumeraçãoFamiliaAcorde = 55 Then

                AjustePosiçãoTonalidadeDoAcorde = 0
                NotasAcordeIndiceLinha = 0

                Index += 1
                ArrayAcordePadrão(Index, 1, 5) = "1"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ImagemAcorde(Index) = "C (Forma de C)"

                Index += 1
                ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 5, 3) = "2"
                ArrayAcordePadrão(Index, 5, 4) = "3"
                ArrayAcordePadrão(Index, 5, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ImagemAcorde(Index) = "C (Forma de A)"

                Index += 1
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 5
                ArrayAcordePadrão(Index, 4, 1) = "4T"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "C (Forma de G)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 3, 3) = "4"
                ImagemAcorde(Index) = "C (Forma de E)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 3, 4) = "2"
                ArrayAcordePadrão(Index, 3, 6) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = "C (Forma de D)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 4 : AjustePosição(Index) = 3
                ArrayAcordePadrão(Index, 0, 2) = "T"
                ArrayAcordePadrão(Index, 0, 6) = " "
                ArrayAcordePadrão(Index, 2, 3) = "1"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 2, 5) = "3"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ImagemAcorde(Index) = "A (Forma de A)"

                Index += 1 : AjustePosição(Index) = 3
                ArrayAcordePadrão(Index, 2, 3) = "P" : LinhaFinalDaSeta(Index) = 5
                ArrayAcordePadrão(Index, 5, 1) = "4T"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "A (Forma de G)"

                Index += 1 : AjustePosição(Index) = 3
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 3, 3) = "4"
                ImagemAcorde(Index) = "A (Forma de E)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1 : AjustePosição(Index) = 3
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 3, 4) = "2"
                ArrayAcordePadrão(Index, 3, 6) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = "A (Forma de D)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1 : AjustePosição(Index) = 3
                ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3"
                ArrayAcordePadrão(Index, 4, 2) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ImagemAcorde(Index) = "A (Forma de C)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste9"

                Index += 4 : AjustePosição(Index) = 5
                ArrayAcordePadrão(Index, 0, 3) = " "
                ArrayAcordePadrão(Index, 0, 4) = " "
                ArrayAcordePadrão(Index, 0, 5) = " "
                ArrayAcordePadrão(Index, 3, 1) = "3T"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "G (Forma de G)"

                Index += 1 : AjustePosição(Index) = 5
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 3, 3) = "4"
                ImagemAcorde(Index) = "G (Forma de E)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste3"

                Index += 1 : AjustePosição(Index) = 5
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 3, 4) = "2"
                ArrayAcordePadrão(Index, 3, 6) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = "G (Forma de D)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1 : AjustePosição(Index) = 5
                ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3"
                ArrayAcordePadrão(Index, 4, 2) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ImagemAcorde(Index) = "G (Forma de C)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1 : AjustePosição(Index) = 5
                ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 3) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 3, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ImagemAcorde(Index) = "G (Forma de A)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 4 : AjustePosição(Index) = 8
                ArrayAcordePadrão(Index, 0, 1) = "T"
                ArrayAcordePadrão(Index, 0, 5) = " "
                ArrayAcordePadrão(Index, 0, 6) = " "
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 2, 2) = "2"
                ArrayAcordePadrão(Index, 2, 3) = "3"
                ImagemAcorde(Index) = "E (Forma de E)"

                Index += 1 : AjustePosição(Index) = 8
                ArrayAcordePadrão(Index, 2, 3) = "1T"
                ArrayAcordePadrão(Index, 4, 4) = "2"
                ArrayAcordePadrão(Index, 4, 6) = "3"
                ArrayAcordePadrão(Index, 5, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = "E (Forma de D)"

                Index += 1 : AjustePosição(Index) = 8
                ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3"
                ArrayAcordePadrão(Index, 4, 2) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ImagemAcorde(Index) = "E (Forma de C)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste4"

                Index += 1 : AjustePosição(Index) = 8
                ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 3) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 3, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ImagemAcorde(Index) = "E (Forma de A)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1 : AjustePosição(Index) = 8
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 5
                ArrayAcordePadrão(Index, 4, 1) = "4T"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "E (Forma de G)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste9"

                Index += 4 : AjustePosição(Index) = 10
                ArrayAcordePadrão(Index, 0, 3) = "T"
                ArrayAcordePadrão(Index, 2, 4) = "1"
                ArrayAcordePadrão(Index, 2, 6) = "2"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = "D (Forma de D)"

                Index += 1 : AjustePosição(Index) = 10
                ArrayAcordePadrão(Index, 2, 2) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 5) = "2"
                ArrayAcordePadrão(Index, 4, 3) = "3"
                ArrayAcordePadrão(Index, 5, 2) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ImagemAcorde(Index) = "D (Forma de C)"

                Index += 1 : AjustePosição(Index) = 10
                ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 3) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 3, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ImagemAcorde(Index) = "D (Forma de A)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1 : AjustePosição(Index) = 10
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 5
                ArrayAcordePadrão(Index, 4, 1) = "4T"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "D (Forma de G)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1 : AjustePosição(Index) = 10
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 3, 3) = "4"
                ImagemAcorde(Index) = "D (Forma de E)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Tríades menores – Tônica no baixo" OrElse NumeraçãoFamiliaAcorde = 56 Then

                AjustePosiçãoTonalidadeDoAcorde = 0
                NotasAcordeIndiceLinha = 0

                Index += 1
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 1, 5) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "Cm (Forma de C)"

                Index += 1
                ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 4, 5) = "2"
                ArrayAcordePadrão(Index, 5, 3) = "3"
                ArrayAcordePadrão(Index, 5, 4) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ImagemAcorde(Index) = "Cm (Forma de A)"

                Index += 1
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 4
                ArrayAcordePadrão(Index, 2, 2) = "2"
                ArrayAcordePadrão(Index, 4, 1) = "4T"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "Cm (Forma de G)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 3, 3) = "4"
                ImagemAcorde(Index) = "Cm (Forma de E)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 2, 6) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = "Cm (Forma de D)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 4 : AjustePosição(Index) = 3
                ArrayAcordePadrão(Index, 0, 2) = "T"
                ArrayAcordePadrão(Index, 0, 6) = " "
                ArrayAcordePadrão(Index, 1, 5) = "1"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 2, 4) = "3"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ImagemAcorde(Index) = "Am (Forma de A)"

                Index += 1 : AjustePosição(Index) = 3
                ArrayAcordePadrão(Index, 2, 3) = "P" : LinhaFinalDaSeta(Index) = 4
                ArrayAcordePadrão(Index, 3, 2) = "2"
                ArrayAcordePadrão(Index, 5, 1) = "4T"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "Am (Forma de G)"

                Index += 1 : AjustePosição(Index) = 3
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 3, 3) = "4"
                ImagemAcorde(Index) = "Am (Forma de E)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1 : AjustePosição(Index) = 3
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 2, 6) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = "Am (Forma de D)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1 : AjustePosição(Index) = 3
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 2, 5) = "3"
                ArrayAcordePadrão(Index, 4, 2) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "Am (Forma de C)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste9"

                Index += 4 : AjustePosição(Index) = 5
                ArrayAcordePadrão(Index, 0, 3) = " "
                ArrayAcordePadrão(Index, 0, 4) = " "
                ArrayAcordePadrão(Index, 1, 2) = "1"
                ArrayAcordePadrão(Index, 3, 1) = "3T"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "Gm (Forma de G)"

                Index += 1 : AjustePosição(Index) = 5
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 3, 3) = "4"
                ImagemAcorde(Index) = "Gm (Forma de E)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste3"

                Index += 1 : AjustePosição(Index) = 5
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 2, 6) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = "Gm (Forma de D)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1 : AjustePosição(Index) = 5
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 2, 5) = "3"
                ArrayAcordePadrão(Index, 4, 2) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "Gm (Forma de C)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1 : AjustePosição(Index) = 5
                ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ImagemAcorde(Index) = "Gm (Forma de A)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 4 : AjustePosição(Index) = 8
                ArrayAcordePadrão(Index, 0, 1) = "T"
                ArrayAcordePadrão(Index, 0, 4) = " "
                ArrayAcordePadrão(Index, 0, 5) = " "
                ArrayAcordePadrão(Index, 0, 6) = " "
                ArrayAcordePadrão(Index, 2, 2) = "2"
                ArrayAcordePadrão(Index, 2, 3) = "3"
                ImagemAcorde(Index) = "Em (Forma de E)"

                Index += 1 : AjustePosição(Index) = 8
                ArrayAcordePadrão(Index, 2, 3) = "1T"
                ArrayAcordePadrão(Index, 3, 6) = "2"
                ArrayAcordePadrão(Index, 4, 4) = "3"
                ArrayAcordePadrão(Index, 5, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = "Em (Forma de D)"

                Index += 1 : AjustePosição(Index) = 8
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 2, 5) = "3"
                ArrayAcordePadrão(Index, 4, 2) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "Em (Forma de C)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste4"

                Index += 1 : AjustePosição(Index) = 8
                ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ImagemAcorde(Index) = "Em (Forma de A)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1 : AjustePosição(Index) = 8
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 4
                ArrayAcordePadrão(Index, 2, 2) = "2"
                ArrayAcordePadrão(Index, 4, 1) = "4T"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "Em (Forma de G)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste9"

                Index += 4 : AjustePosição(Index) = 10
                ArrayAcordePadrão(Index, 0, 3) = "T"
                ArrayAcordePadrão(Index, 1, 6) = "1"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = "Dm (Forma de D)"

                Index += 1 : AjustePosição(Index) = 10
                ArrayAcordePadrão(Index, 2, 4) = "1"
                ArrayAcordePadrão(Index, 3, 3) = "2"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 5, 2) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "Dm (Forma de C)"

                Index += 1 : AjustePosição(Index) = 10
                ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ImagemAcorde(Index) = "Dm (Forma de A)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1 : AjustePosição(Index) = 10
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 4
                ArrayAcordePadrão(Index, 2, 2) = "2"
                ArrayAcordePadrão(Index, 4, 1) = "4T"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "Dm (Forma de G)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1 : AjustePosição(Index) = 10
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 3, 3) = "4"
                ImagemAcorde(Index) = "Dm (Forma de E)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Tríades aumentadas – Tônica no baixo" OrElse NumeraçãoFamiliaAcorde = 57 Then

                StringSubstituição(0) = "b6" : StringSubstituição(1) = "#5" : SubstituirString()

                AjustePosiçãoTonalidadeDoAcorde = 0
                NotasAcordeIndiceLinha = 0

                Index += 1
                ArrayAcordePadrão(Index, 1, 4) = "P" : LinhaFinalDaSeta(Index) = 5
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "C(#5) (Forma de C)"

                Index += 1
                ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 5, 4) = "2"
                ArrayAcordePadrão(Index, 5, 5) = "3"
                ArrayAcordePadrão(Index, 6, 3) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "C(#5) (Forma de A)"

                Index += 1
                ArrayAcordePadrão(Index, 1, 4) = "P" : LinhaFinalDaSeta(Index) = 5
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 4, 1) = "4T"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "C(#5) (Forma de G)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 2, 5) = "3"
                ArrayAcordePadrão(Index, 3, 3) = "4"
                ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = "C(#5) (Forma de E)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                ArrayAcordePadrão(Index, 1, 3) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 6) = "2"
                ArrayAcordePadrão(Index, 4, 4) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = "C(#5) (Forma de D)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 4 : AjustePosição(Index) = 3
                ArrayAcordePadrão(Index, 0, 2) = "T"
                ArrayAcordePadrão(Index, 2, 4) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 3) = "2"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "A(#5) (Forma de A)"

                Index += 1 : AjustePosição(Index) = 3
                ArrayAcordePadrão(Index, 2, 4) = "P" : LinhaFinalDaSeta(Index) = 5
                ArrayAcordePadrão(Index, 3, 3) = "2"
                ArrayAcordePadrão(Index, 5, 1) = "4T"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "A(#5) (Forma de G)"

                Index += 1 : AjustePosição(Index) = 3
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 2, 5) = "3"
                ArrayAcordePadrão(Index, 3, 3) = "4"
                ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = "A(#5) (Forma de E)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1 : AjustePosição(Index) = 3
                ArrayAcordePadrão(Index, 1, 3) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 6) = "2"
                ArrayAcordePadrão(Index, 4, 4) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = "A(#5) (Forma de D)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1 : AjustePosição(Index) = 3
                ArrayAcordePadrão(Index, 1, 4) = "P" : LinhaFinalDaSeta(Index) = 5
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "A(#5) (Forma de C)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 4 : AjustePosição(Index) = 5
                ArrayAcordePadrão(Index, 0, 4) = " "
                ArrayAcordePadrão(Index, 0, 5) = " "
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 3, 1) = "3T"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "G(#5) (Forma de G)"

                Index += 1 : AjustePosição(Index) = 5
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 2, 5) = "3"
                ArrayAcordePadrão(Index, 3, 3) = "4"
                ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = "G(#5) (Forma de E)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste3"

                Index += 1 : AjustePosição(Index) = 5
                ArrayAcordePadrão(Index, 1, 3) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 6) = "2"
                ArrayAcordePadrão(Index, 4, 4) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = "G(#5) (Forma de D)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1 : AjustePosição(Index) = 5
                ArrayAcordePadrão(Index, 1, 4) = "P" : LinhaFinalDaSeta(Index) = 5
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "G(#5) (Forma de C)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1 : AjustePosição(Index) = 5
                ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 4) = "2"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 4, 3) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "G(#5) (Forma de A)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 4 : AjustePosição(Index) = 8
                ArrayAcordePadrão(Index, 0, 1) = "T"
                ArrayAcordePadrão(Index, 0, 6) = " "
                ArrayAcordePadrão(Index, 1, 4) = "P" : LinhaFinalDaSeta(Index) = 5
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = "E(#5) (Forma de E)"

                Index += 1 : AjustePosição(Index) = 8
                ArrayAcordePadrão(Index, 2, 3) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 4, 6) = "2"
                ArrayAcordePadrão(Index, 5, 4) = "3"
                ArrayAcordePadrão(Index, 5, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = "E(#5) (Forma de D)"

                Index += 1 : AjustePosição(Index) = 8
                ArrayAcordePadrão(Index, 1, 4) = "P" : LinhaFinalDaSeta(Index) = 5
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "E(#5) (Forma de C)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1 : AjustePosição(Index) = 8
                ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 4) = "2"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 4, 3) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "E(#5) (Forma de A)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1 : AjustePosição(Index) = 8
                ArrayAcordePadrão(Index, 1, 4) = "P" : LinhaFinalDaSeta(Index) = 5
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 4, 1) = "4T"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "E(#5) (Forma de G)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste9"

                Index += 4 : AjustePosição(Index) = 10
                ArrayAcordePadrão(Index, 0, 3) = "T"
                ArrayAcordePadrão(Index, 2, 6) = "1"
                ArrayAcordePadrão(Index, 3, 4) = "2"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = "D(#5) (Forma de D)"

                Index += 1 : AjustePosição(Index) = 10
                ArrayAcordePadrão(Index, 3, 4) = "P" : LinhaFinalDaSeta(Index) = 5
                ArrayAcordePadrão(Index, 4, 3) = "2"
                ArrayAcordePadrão(Index, 5, 2) = "3T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "D(#5) (Forma de C)"

                Index += 1 : AjustePosição(Index) = 10
                ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 4) = "2"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 4, 3) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "D(#5) (Forma de A)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1 : AjustePosição(Index) = 10
                ArrayAcordePadrão(Index, 1, 4) = "P" : LinhaFinalDaSeta(Index) = 5
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 4, 1) = "4T"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "D(#5) (Forma de G)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1 : AjustePosição(Index) = 10
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 2, 5) = "3"
                ArrayAcordePadrão(Index, 3, 3) = "4"
                ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = "D(#5) (Forma de E)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Tríades com quarta – Tônica no baixo" OrElse NumeraçãoFamiliaAcorde = 58 Then

                AjustePosiçãoTonalidadeDoAcorde = 0
                NotasAcordeIndiceLinha = 0

                Index += 1
                ArrayAcordePadrão(Index, 1, 5) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 3, 3) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "C4 (Forma de C)"

                Index += 1
                ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 5, 3) = "2"
                ArrayAcordePadrão(Index, 5, 4) = "3"
                ArrayAcordePadrão(Index, 6, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ImagemAcorde(Index) = "C4 (Forma de A)"

                Index += 1
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 5
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 4, 1) = "4T"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "C4 (Forma de G)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 2) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ImagemAcorde(Index) = "C4 (Forma de E)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 3, 4) = "2"
                ArrayAcordePadrão(Index, 4, 5) = "3"
                ArrayAcordePadrão(Index, 4, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = "C4 (Forma de D)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 4 : AjustePosição(Index) = 3
                ArrayAcordePadrão(Index, 0, 2) = "T"
                ArrayAcordePadrão(Index, 0, 6) = " "
                ArrayAcordePadrão(Index, 2, 3) = "P" : LinhaFinalDaSeta(Index) = 5
                ArrayAcordePadrão(Index, 3, 5) = "2"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ImagemAcorde(Index) = "A4 (Forma de A)"

                Index += 1 : AjustePosição(Index) = 3
                ArrayAcordePadrão(Index, 2, 3) = "P" : LinhaFinalDaSeta(Index) = 5
                ArrayAcordePadrão(Index, 3, 5) = "2"
                ArrayAcordePadrão(Index, 5, 1) = "4T"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "A4 (Forma de G)"

                Index += 1 : AjustePosição(Index) = 3
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 2) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ImagemAcorde(Index) = "A4 (Forma de E)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1 : AjustePosição(Index) = 3
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 3, 4) = "2"
                ArrayAcordePadrão(Index, 4, 5) = "3"
                ArrayAcordePadrão(Index, 4, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = "A4 (Forma de D)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1 : AjustePosição(Index) = 3
                ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 4, 2) = "3T"
                ArrayAcordePadrão(Index, 4, 3) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ImagemAcorde(Index) = "A4 (Forma de C)" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 1, 0) = "Traste9"

                Index += 4 : AjustePosição(Index) = 5
                ArrayAcordePadrão(Index, 0, 3) = " "
                ArrayAcordePadrão(Index, 0, 4) = " "
                ArrayAcordePadrão(Index, 1, 5) = "1"
                ArrayAcordePadrão(Index, 3, 1) = "3T"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "G4 (Forma de G)"

                Index += 1 : AjustePosição(Index) = 5
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 2) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ImagemAcorde(Index) = "G4 (Forma de E)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste3"

                Index += 1 : AjustePosição(Index) = 5
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 3, 4) = "2"
                ArrayAcordePadrão(Index, 4, 5) = "3"
                ArrayAcordePadrão(Index, 4, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = "G4 (Forma de D)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1 : AjustePosição(Index) = 5
                ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 4, 2) = "3T"
                ArrayAcordePadrão(Index, 4, 3) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ImagemAcorde(Index) = "A4 (Forma de C)" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "G4 (Forma de C)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1 : AjustePosição(Index) = 5
                ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 3) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ImagemAcorde(Index) = "G4 (Forma de A)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 4 : AjustePosição(Index) = 8
                ArrayAcordePadrão(Index, 0, 1) = "T"
                ArrayAcordePadrão(Index, 0, 5) = " "
                ArrayAcordePadrão(Index, 0, 6) = " "
                ArrayAcordePadrão(Index, 2, 2) = "2"
                ArrayAcordePadrão(Index, 2, 3) = "3"
                ArrayAcordePadrão(Index, 2, 4) = "4"
                ImagemAcorde(Index) = "E4 (Forma de E)"

                Index += 1 : AjustePosição(Index) = 8
                ArrayAcordePadrão(Index, 2, 3) = "1T"
                ArrayAcordePadrão(Index, 4, 4) = "2"
                ArrayAcordePadrão(Index, 5, 5) = "3"
                ArrayAcordePadrão(Index, 5, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = "E4 (Forma de D)"

                Index += 1 : AjustePosição(Index) = 8
                ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 4, 2) = "3T"
                ArrayAcordePadrão(Index, 4, 3) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ImagemAcorde(Index) = "A4 (Forma de C)" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "E4 (Forma de C)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste4"

                Index += 1 : AjustePosição(Index) = 8
                ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 3) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ImagemAcorde(Index) = "E4 (Forma de A)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1 : AjustePosição(Index) = 8
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 5
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 4, 1) = "4T"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "E4 (Forma de G)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste9"

                Index += 4 : AjustePosição(Index) = 10
                ArrayAcordePadrão(Index, 0, 3) = "T"
                ArrayAcordePadrão(Index, 2, 4) = "1"
                ArrayAcordePadrão(Index, 3, 5) = "2"
                ArrayAcordePadrão(Index, 3, 6) = "3"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = "D4 (Forma de D)"

                Index += 1 : AjustePosição(Index) = 10
                ArrayAcordePadrão(Index, 2, 2) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 5) = "2"
                ArrayAcordePadrão(Index, 5, 2) = "3T"
                ArrayAcordePadrão(Index, 5, 3) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ImagemAcorde(Index) = "A4 (Forma de C)" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "D4 (Forma de C)"

                Index += 1 : AjustePosição(Index) = 10
                ArrayAcordePadrão(Index, 1, 2) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 3) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = ""
                ImagemAcorde(Index) = "D4 (Forma de A)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1 : AjustePosição(Index) = 10
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 5
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 4, 1) = "4T"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = "D4 (Forma de G)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1 : AjustePosição(Index) = 10
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 2) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ImagemAcorde(Index) = "D4 (Forma de E)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Tríades maiores – 3 cordas com inversões" OrElse NumeraçãoFamiliaAcorde = 59 Then

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "2" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 0, 4) = "-1"
                ArrayAcordePadrão(Index, 0, 6) = "0"
                ArrayAcordePadrão(Index, 1, 5) = "1T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeQuinta

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 0, 4) = "0"
                ArrayAcordePadrão(Index, 1, 5) = "1T"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeTerça

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 0, 4) = "0"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 3, 1) = "3"
                ArrayAcordePadrão(Index, 3, 2) = "4T"
                ArrayAcordePadrão(Index, 7, 4) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeQuinta

                Index += 5
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 3, 6) = "1"
                ArrayAcordePadrão(Index, 5, 4) = "3T"
                ArrayAcordePadrão(Index, 5, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ImagemAcorde(Index) = TonalidadeAcorde

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 5 : ArrayAcordePadrão(Index, 1, 4) = "T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 4 : ArrayAcordePadrão(Index, 1, 4) = "T"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 4, 1) = "4T"
                ArrayAcordePadrão(Index, 7, 4) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 5
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 5) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 6) = "T"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 5) = "1"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                ArrayAcordePadrão(Index, 1, 4) = "2"
                ArrayAcordePadrão(Index, 2, 2) = "3"
                ArrayAcordePadrão(Index, 2, 3) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste9"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 3 : ArrayAcordePadrão(Index, 1, 3) = "T"
                ArrayAcordePadrão(Index, 3, 1) = "3"
                ArrayAcordePadrão(Index, 7, 4) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                'Index += 5
                'ArrayAcordePadrão(Index, 1, 4) = "1"
                'ArrayAcordePadrão(Index, 1, 6) = "2"
                'ArrayAcordePadrão(Index, 2, 5) = "3T"
                'ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                'ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeQuinta
                'ArrayAcordePadrão(Index, 1, 0) = "Traste12"

                'Index += 1
                'ArrayAcordePadrão(Index, 1, 4) = "1"
                'ArrayAcordePadrão(Index, 2, 5) = "2T"
                'ArrayAcordePadrão(Index, 3, 3) = "3"
                'ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                'ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeTerça
                'ArrayAcordePadrão(Index, 1, 0) = "Traste12"

                'Index += 1
                'ArrayAcordePadrão(Index, 1, 4) = "1"
                'ArrayAcordePadrão(Index, 3, 3) = "3"
                'ArrayAcordePadrão(Index, 4, 2) = "4T"
                'ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                'ImagemAcorde(Index) = TonalidadeAcorde
                'ArrayAcordePadrão(Index, 1, 0) = "Traste12"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Tríades menores – 3 cordas com inversões" OrElse NumeraçãoFamiliaAcorde = 60 Then

                Index += 1
                ArrayAcordePadrão(Index, 1, 6) = "1"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 5) = "3T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste11"

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 0, 4) = "0"
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 1, 5) = "2T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeTerçaMenor

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 0, 4) = "0"
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "m"

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 3, 1) = "3"
                ArrayAcordePadrão(Index, 3, 2) = "4T"
                ArrayAcordePadrão(Index, 7, 4) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeQuinta

                Index += 5
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 6) = "1"
                ArrayAcordePadrão(Index, 4, 5) = "2"
                ArrayAcordePadrão(Index, 5, 4) = "3T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "m"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 4, 5) = "1"
                ArrayAcordePadrão(Index, 5, 3) = "2"
                ArrayAcordePadrão(Index, 5, 4) = "3T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeQuinta

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 4 : ArrayAcordePadrão(Index, 1, 4) = "T"
                ArrayAcordePadrão(Index, 2, 2) = "2"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeTerçaMenor
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 2, 2) = "2"
                ArrayAcordePadrão(Index, 4, 1) = "4T"
                ArrayAcordePadrão(Index, 7, 4) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "m"
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 5
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 6) = "T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeTerçaMenor
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "P" : LinhaFinalDaSeta(Index) = 5
                ArrayAcordePadrão(Index, 3, 3) = "3T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "m"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "4"
                ArrayAcordePadrão(Index, 3, 3) = "3T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 3 : ArrayAcordePadrão(Index, 1, 3) = "T"
                ArrayAcordePadrão(Index, 2, 1) = "2"
                ArrayAcordePadrão(Index, 7, 4) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeTerçaMenor
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                'Index += 5
                'ArrayAcordePadrão(Index, 1, 6) = "1"
                'ArrayAcordePadrão(Index, 2, 4) = "2"
                'ArrayAcordePadrão(Index, 3, 5) = "3T"
                'ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                'ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeQuinta
                'ArrayAcordePadrão(Index, 1, 0) = "Traste11"

                'Index += 1
                'ArrayAcordePadrão(Index, 1, 4) = "1"
                'ArrayAcordePadrão(Index, 2, 3) = "2"
                'ArrayAcordePadrão(Index, 2, 5) = "3T"
                'ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                'ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeTerçaMenor
                'ArrayAcordePadrão(Index, 1, 0) = "Traste12"

                'Index += 1
                'ArrayAcordePadrão(Index, 1, 4) = "1"
                'ArrayAcordePadrão(Index, 2, 3) = "2"
                'ArrayAcordePadrão(Index, 4, 2) = "4T"
                'ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                'ImagemAcorde(Index) = TonalidadeAcorde & "m"
                'ArrayAcordePadrão(Index, 1, 0) = "Traste12"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Tríades com quarta – 3 cordas com inversões" OrElse NumeraçãoFamiliaAcorde = 61 Then

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 0, 4) = "0"
                ArrayAcordePadrão(Index, 1, 5) = "1T"
                ArrayAcordePadrão(Index, 1, 6) = "2"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "4/" & BaixoDoAcordeQuinta

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 0, 4) = "0"
                ArrayAcordePadrão(Index, 1, 5) = "1T"
                ArrayAcordePadrão(Index, 3, 3) = "3"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "4/" & BaixoDoAcordeQuarta

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 0, 4) = "0"
                ArrayAcordePadrão(Index, 3, 2) = "1T"
                ArrayAcordePadrão(Index, 3, 3) = "2"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "4"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 1) = "P" : LinhaFinalDaSeta(Index) = 3 : ArrayAcordePadrão(Index, 3, 2) = "T"
                ArrayAcordePadrão(Index, 7, 4) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "4/" & BaixoDoAcordeQuinta

                Index += 5
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 6) = "1"
                ArrayAcordePadrão(Index, 5, 4) = "3T"
                ArrayAcordePadrão(Index, 6, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "4"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 4 : ArrayAcordePadrão(Index, 1, 4) = "T"
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "4/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 4 : ArrayAcordePadrão(Index, 1, 4) = "T"
                ArrayAcordePadrão(Index, 4, 2) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "4/" & BaixoDoAcordeQuarta
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 4, 1) = "3T"
                ArrayAcordePadrão(Index, 4, 2) = "4"
                ArrayAcordePadrão(Index, 7, 4) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "4"
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 5
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 5) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 6) = "T"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "4/" & BaixoDoAcordeQuarta
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 5) = "1"
                ArrayAcordePadrão(Index, 3, 3) = "3T"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "4"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 4 : ArrayAcordePadrão(Index, 1, 3) = "T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "4/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 3 : ArrayAcordePadrão(Index, 1, 3) = "T"
                ArrayAcordePadrão(Index, 4, 1) = "4"
                ArrayAcordePadrão(Index, 7, 4) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "4/" & BaixoDoAcordeQuarta
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                'Index += 5
                'ArrayAcordePadrão(Index, 1, 4) = "1"
                'ArrayAcordePadrão(Index, 2, 5) = "2T"
                'ArrayAcordePadrão(Index, 2, 6) = "3"
                'ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                'ImagemAcorde(Index) = TonalidadeAcorde & "4/" & BaixoDoAcordeQuinta
                'ArrayAcordePadrão(Index, 1, 0) = "Traste12"

                'Index += 1
                'ArrayAcordePadrão(Index, 1, 4) = "1"
                'ArrayAcordePadrão(Index, 2, 5) = "2T"
                'ArrayAcordePadrão(Index, 4, 3) = "4"
                'ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                'ImagemAcorde(Index) = TonalidadeAcorde & "4/" & BaixoDoAcordeQuarta
                'ArrayAcordePadrão(Index, 1, 0) = "Traste12"

                'Index += 1
                'ArrayAcordePadrão(Index, 1, 4) = "1"
                'ArrayAcordePadrão(Index, 4, 2) = "3T"
                'ArrayAcordePadrão(Index, 4, 3) = "4"
                'ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                'ImagemAcorde(Index) = TonalidadeAcorde & "4"
                'ArrayAcordePadrão(Index, 1, 0) = "Traste12"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Tríades maiores – 4 cordas com inversões" OrElse NumeraçãoFamiliaAcorde = 62 Then

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "2" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 0, 4) = "-1"
                ArrayAcordePadrão(Index, 0, 6) = "0"
                ArrayAcordePadrão(Index, 1, 5) = "1T"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeTerça

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                    ArrayAcordePadrão(Index, 3, 6) = "3"
                Else
                    ArrayAcordePadrão(Index, 3, 6) = "4"
                End If
                ArrayAcordePadrão(Index, 0, 4) = "0"
                ArrayAcordePadrão(Index, 1, 5) = "1T"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeTerça

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 6) = "1"
                ArrayAcordePadrão(Index, 5, 3) = "2"
                ArrayAcordePadrão(Index, 5, 4) = "3T"
                ArrayAcordePadrão(Index, 5, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeQuinta

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 4, 6) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 5) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = TonalidadeAcorde
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 3, 4) = "2"
                ArrayAcordePadrão(Index, 3, 6) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = TonalidadeAcorde
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 3
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 0, 4) = "0"
                ArrayAcordePadrão(Index, 1, 5) = "1"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 5
                ArrayAcordePadrão(Index, 5, 3) = "2"
                ArrayAcordePadrão(Index, 5, 4) = "3"
                ArrayAcordePadrão(Index, 5, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 5 : ArrayAcordePadrão(Index, 1, 4) = "T"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 5 : ArrayAcordePadrão(Index, 1, 4) = "T"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 5) = "1"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 3, 3) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "1"
                ArrayAcordePadrão(Index, 3, 4) = "2"
                ArrayAcordePadrão(Index, 4, 5) = "3T"
                ArrayAcordePadrão(Index, 5, 3) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 3
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 0, 4) = "1"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 3, 1) = "3"
                ArrayAcordePadrão(Index, 3, 2) = "4T"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeQuinta

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "1"
                ArrayAcordePadrão(Index, 3, 3) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3T"
                ArrayAcordePadrão(Index, 5, 2) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste3"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 4
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 4, 1) = "4T"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "1T"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 3, 3) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 3 : ArrayAcordePadrão(Index, 1, 3) = "T"
                ArrayAcordePadrão(Index, 3, 1) = "3"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 0, 1) = "P" : LinhaFinalDaSeta(Index) = 4
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "/" & BaixoDoAcordeTerça

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Tríades menores – 4 cordas com inversões" OrElse NumeraçãoFamiliaAcorde = 63 Then

                Index += 1
                ArrayAcordePadrão(Index, 1, 6) = "1"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3S"
                ArrayAcordePadrão(Index, 3, 5) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6 (omit5)"  'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeTerçaMenor
                ArrayAcordePadrão(Index, 1, 0) = "Traste11"

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 0, 4) = "0"
                ArrayAcordePadrão(Index, 1, 3) = "1S"
                ArrayAcordePadrão(Index, 1, 5) = "2T"
                ArrayAcordePadrão(Index, 3, 6) = "3"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6 (omit5)"  'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeTerçaMenor

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 6) = "1"
                ArrayAcordePadrão(Index, 4, 5) = "2"
                ArrayAcordePadrão(Index, 5, 3) = "3"
                ArrayAcordePadrão(Index, 5, 4) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeQuinta

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 4, 4) = "2"
                ArrayAcordePadrão(Index, 4, 5) = "3"
                ArrayAcordePadrão(Index, 4, 6) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 3) = "3T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "m"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 2, 6) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "m"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 3
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                    ArrayAcordePadrão(Index, 3, 2) = "3T"
                Else
                    ArrayAcordePadrão(Index, 3, 2) = "4T"
                End If
                ArrayAcordePadrão(Index, 0, 4) = "0"
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 1, 5) = "2"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "m"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 2) = "1T"
                ArrayAcordePadrão(Index, 4, 5) = "2"
                ArrayAcordePadrão(Index, 5, 3) = "3"
                ArrayAcordePadrão(Index, 5, 4) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "m"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 4, 5) = "1"
                ArrayAcordePadrão(Index, 5, 3) = "2"
                ArrayAcordePadrão(Index, 5, 4) = "3T"
                ArrayAcordePadrão(Index, 6, 2) = "4S"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6 (omit5)"  'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeTerçaMenor

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 4 : ArrayAcordePadrão(Index, 1, 4) = "T"
                ArrayAcordePadrão(Index, 2, 2) = "2S"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6 (omit5)"  'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeTerçaMenor
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "P" : LinhaFinalDaSeta(Index) = 5
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 3, 3) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "1"
                ArrayAcordePadrão(Index, 3, 4) = "2"
                ArrayAcordePadrão(Index, 4, 3) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 3
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                    ArrayAcordePadrão(Index, 3, 1) = "2"
                    ArrayAcordePadrão(Index, 3, 2) = "3T"
                Else
                    ArrayAcordePadrão(Index, 3, 1) = "3"
                    ArrayAcordePadrão(Index, 3, 2) = "4T"
                End If
                ArrayAcordePadrão(Index, 0, 4) = "0"
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeQuinta

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 1) = "1"
                ArrayAcordePadrão(Index, 5, 3) = "2"
                ArrayAcordePadrão(Index, 5, 4) = "3T"
                ArrayAcordePadrão(Index, 6, 2) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeQuinta

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 4
                ArrayAcordePadrão(Index, 2, 2) = "2"
                ArrayAcordePadrão(Index, 4, 1) = "4T"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "m"
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 4
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 3, 3) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "m"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 3 : ArrayAcordePadrão(Index, 1, 3) = "T"
                ArrayAcordePadrão(Index, 2, 1) = "2S"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6 (omit5)"  'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeTerçaMenor
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                ArrayAcordePadrão(Index, 1, 1) = "1S"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3"
                ArrayAcordePadrão(Index, 5, 2) = "4T"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6 (omit5)"  'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m/" & BaixoDoAcordeTerçaMenor
                ArrayAcordePadrão(Index, 1, 0) = "Traste11"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Tríades aumentadas – 4 cordas com inversões" OrElse NumeraçãoFamiliaAcorde = 64 Then

                StringSubstituição(0) = "b6" : StringSubstituição(1) = "#5" : SubstituirString()

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 0, 6) = "0"
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 1, 5) = "2T"
                ArrayAcordePadrão(Index, 2, 3) = "3"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "(#5)/" & BaixoDoAcordeTerça

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 5) = "T"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 4, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "(#5)/" & BaixoDoAcordeTerça

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 4, 6) = "1"
                ArrayAcordePadrão(Index, 5, 4) = "2T"
                ArrayAcordePadrão(Index, 5, 5) = "3"
                ArrayAcordePadrão(Index, 6, 3) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "(#5)/" & BaixoDoAcordeQuintaAumentada

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 4, 6) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "(#5)/" & BaixoDoAcordeQuintaAumentada
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 6) = "1"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 2, 5) = "3"
                ArrayAcordePadrão(Index, 3, 3) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "(#5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 3, 6) = "2"
                ArrayAcordePadrão(Index, 4, 4) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "(#5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 3
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "P" : LinhaFinalDaSeta(Index) = 5
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "(#5)"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 2) = "1T"
                ArrayAcordePadrão(Index, 5, 4) = "2"
                ArrayAcordePadrão(Index, 5, 5) = "3"
               ArrayAcordePadrão(Index, 6, 3) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "(#5)"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 5 : ArrayAcordePadrão(Index, 1, 4) = "T"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "(#5)/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "1T"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 5, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "(#5)/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                ArrayAcordePadrão(Index, 1, 4) = "P" : LinhaFinalDaSeta(Index) = 5
                ArrayAcordePadrão(Index, 2, 3) = "2T"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "(#5)/" & BaixoDoAcordeQuintaAumentada
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 1, 0) = "Traste9"

                Index += 1
                ArrayAcordePadrão(Index, 1, 2) = "1"
                ArrayAcordePadrão(Index, 3, 4) = "2"
                ArrayAcordePadrão(Index, 3, 5) = "3T"
                ArrayAcordePadrão(Index, 4, 3) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "(#5)/" & BaixoDoAcordeQuintaAumentada
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 1, 0) = "Traste11"

                Index += 3
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 4, 1) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "(#5)/" & BaixoDoAcordeQuintaAumentada

                Index += 1
                ArrayAcordePadrão(Index, 1, 1) = "1"
                ArrayAcordePadrão(Index, 2, 4) = "2T"
                ArrayAcordePadrão(Index, 3, 3) = "3"
                ArrayAcordePadrão(Index, 4, 2) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "(#5)/" & BaixoDoAcordeQuintaAumentada
                ArrayAcordePadrão(Index, 1, 0) = "Traste4"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 4, 1) = "4T"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "(#5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "1T"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3"
                ArrayAcordePadrão(Index, 4, 2) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "(#5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 2, 2) = "2"
                ArrayAcordePadrão(Index, 3, 1) = "3"
                ArrayAcordePadrão(Index, 4, 4) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "(#5)/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 0, 1) = "0"
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "(#5)/" & BaixoDoAcordeTerça

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Tríades com quarta – 4 cordas com inversões" OrElse NumeraçãoFamiliaAcorde = 65 Then

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                    ArrayAcordePadrão(Index, 1, 5) = "1T"
                    ArrayAcordePadrão(Index, 1, 6) = "2"
                Else
                    ArrayAcordePadrão(Index, 1, 5) = "PT" : LinhaFinalDaSeta(Index) = 6
                End If
                ArrayAcordePadrão(Index, 0, 4) = "0"
                ArrayAcordePadrão(Index, 3, 3) = "3"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "4/" & BaixoDoAcordeQuarta

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                    ArrayAcordePadrão(Index, 3, 3) = "2"
                    ArrayAcordePadrão(Index, 3, 6) = "3"
                Else
                    ArrayAcordePadrão(Index, 3, 3) = "3"
                    ArrayAcordePadrão(Index, 3, 6) = "4"
                End If
                ArrayAcordePadrão(Index, 0, 4) = "0"
                ArrayAcordePadrão(Index, 1, 5) = "1T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "4/" & BaixoDoAcordeQuarta

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 6) = "1"
                ArrayAcordePadrão(Index, 5, 3) = "2"
                ArrayAcordePadrão(Index, 5, 4) = "3T"
                ArrayAcordePadrão(Index, 6, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "4/" & BaixoDoAcordeQuinta

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 4, 6) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "4/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 5) = "P" : LinhaFinalDaSeta(Index) = 6
                ArrayAcordePadrão(Index, 3, 3) = "3T"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "4"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 3, 4) = "2"
                ArrayAcordePadrão(Index, 4, 5) = "3"
                ArrayAcordePadrão(Index, 4, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "4"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 3
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                    ArrayAcordePadrão(Index, 3, 2) = "2T"
                    ArrayAcordePadrão(Index, 3, 3) = "3"
                Else
                    ArrayAcordePadrão(Index, 3, 2) = "3T"
                    ArrayAcordePadrão(Index, 3, 3) = "4"
                End If
                ArrayAcordePadrão(Index, 0, 4) = "0"
                ArrayAcordePadrão(Index, 1, 5) = "1"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "4"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 2) = "1T"
                ArrayAcordePadrão(Index, 5, 3) = "2"
                ArrayAcordePadrão(Index, 5, 4) = "3"
                ArrayAcordePadrão(Index, 6, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "4"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 5 : ArrayAcordePadrão(Index, 1, 4) = "T"
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 4, 2) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "4/" & BaixoDoAcordeQuarta
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 4 : ArrayAcordePadrão(Index, 1, 4) = "T"
                ArrayAcordePadrão(Index, 4, 2) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "4/" & BaixoDoAcordeQuarta
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 5) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3T"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "4/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "1"
                ArrayAcordePadrão(Index, 3, 4) = "2"
                ArrayAcordePadrão(Index, 4, 5) = "3T"
                ArrayAcordePadrão(Index, 6, 3) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "4/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 3
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 0, 4) = "1"
                ArrayAcordePadrão(Index, 3, 1) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 3, 3) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "4/" & BaixoDoAcordeQuinta

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 1) = "P" : LinhaFinalDaSeta(Index) = 4
                ArrayAcordePadrão(Index, 5, 4) = "3T"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "4/" & BaixoDoAcordeQuinta

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 4
                ArrayAcordePadrão(Index, 4, 1) = "3T"
                ArrayAcordePadrão(Index, 4, 2) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "4"
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "1T"
                ArrayAcordePadrão(Index, 3, 2) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "4"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 3 : ArrayAcordePadrão(Index, 1, 3) = "T"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 1) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "4/" & BaixoDoAcordeQuarta
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 0, 4) = "0"
                ArrayAcordePadrão(Index, 1, 1) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 3, 3) = "3"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "4/" & BaixoDoAcordeQuarta

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

           ElseIf FamiliaAcorde = "Tétrades com sexta – " & TonalidadeAcorde & " " & BaixoDoAcordeTerça & " " & BaixoDoAcordeQuinta & " " & BaixoDoAcordeSexta OrElse NumeraçãoFamiliaAcorde = 66 Then

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 5) = "1T"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 2, 4) = "3S"
                ArrayAcordePadrão(Index, 3, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7/" & BaixoDoAcordeTerça 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "6/" & BaixoDoAcordeTerça

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 4) = "T" : ArrayAcordePadrão(Index, 1, 6) = "S"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7/" & BaixoDoAcordeQuinta 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "6/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1S"
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 2, 6) = "3T"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7" 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "6/" & BaixoDoAcordeSexta
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "PT" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 5) = "S"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 3, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "6"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 5
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 0, 2) = "PS" : LinhaFinalDaSeta(Index) = 5
                ArrayAcordePadrão(Index, 1, 5) = "1T"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7" 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "6/" & BaixoDoAcordeSexta

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 4) = "1S"
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 5, 3) = "3"
                ArrayAcordePadrão(Index, 5, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "6"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "1T"
                ArrayAcordePadrão(Index, 3, 2) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3S"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7/" & BaixoDoAcordeTerça 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "6/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 2, 2) = "2"
                ArrayAcordePadrão(Index, 2, 3) = "3T"
                ArrayAcordePadrão(Index, 2, 5) = "4S"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7/" & BaixoDoAcordeQuinta 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "6/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste9"

                Index += 5
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 3) = "P" : LinhaFinalDaSeta(Index) = 4 : ArrayAcordePadrão(Index, 2, 4) = "S"
                ArrayAcordePadrão(Index, 3, 1) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7/" & BaixoDoAcordeQuinta 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "6/" & BaixoDoAcordeQuinta

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PS" : LinhaFinalDaSeta(Index) = 4 : ArrayAcordePadrão(Index, 1, 4) = "T"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7" 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "6/" & BaixoDoAcordeSexta
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1S"
                ArrayAcordePadrão(Index, 2, 1) = "2T"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 2) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "6"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 3, 1) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3S"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7/" & BaixoDoAcordeTerça 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "6/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Tétrades menores com sexta – " & TonalidadeAcorde & " " & BaixoDoAcordeTerçaMenor & " " & BaixoDoAcordeQuinta & " " & BaixoDoAcordeSexta OrElse NumeraçãoFamiliaAcorde = 67 Then

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 5) = "T"
                ArrayAcordePadrão(Index, 2, 4) = "2S"
                ArrayAcordePadrão(Index, 3, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7(b5)/" & BaixoDoAcordeTerçaMenor 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m6/" & BaixoDoAcordeTerçaMenor

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 4, 5) = "1"
                ArrayAcordePadrão(Index, 5, 3) = "2"
                ArrayAcordePadrão(Index, 5, 4) = "3T"
                ArrayAcordePadrão(Index, 5, 6) = "4S"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7(b5)/" & BaixoDoAcordeQuinta 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m6/" & BaixoDoAcordeQuinta

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1S"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 2, 5) = "3"
                ArrayAcordePadrão(Index, 2, 6) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7(b5)" 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m6/" & BaixoDoAcordeSexta
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "PT" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 5) = "S"
                ArrayAcordePadrão(Index, 2, 6) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7(b5)/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m6"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 5
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 0, 2) = "PS" : LinhaFinalDaSeta(Index) = 5
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 1, 5) = "2T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7(b5)" 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m6/" & BaixoDoAcordeSexta

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 4) = "1S"
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 4, 5) = "3"
                ArrayAcordePadrão(Index, 5, 3) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7(b5)/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m6"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "1T"
                ArrayAcordePadrão(Index, 2, 2) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3S"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7(b5)/" & BaixoDoAcordeTerçaMenor 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m6/" & BaixoDoAcordeTerçaMenor
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3T"
                ArrayAcordePadrão(Index, 3, 5) = "4S"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7(b5)/" & BaixoDoAcordeQuinta 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m6/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 5
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 2, 4) = "2S"
                ArrayAcordePadrão(Index, 3, 1) = "3"
                ArrayAcordePadrão(Index, 3, 2) = "4T"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7(b5)/" & BaixoDoAcordeQuinta 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m6/" & BaixoDoAcordeQuinta

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PS" : LinhaFinalDaSeta(Index) = 4 : ArrayAcordePadrão(Index, 1, 4) = "T"
                ArrayAcordePadrão(Index, 2, 2) = "2"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7(b5)" 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m6/" & BaixoDoAcordeSexta
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1S"
                ArrayAcordePadrão(Index, 2, 1) = "2T"
                ArrayAcordePadrão(Index, 2, 4) = "3"
                ArrayAcordePadrão(Index, 4, 2) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7(b5)/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m6"
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 2, 1) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3S"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeSexta & "m7(b5)/" & BaixoDoAcordeTerçaMenor 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m6/" & BaixoDoAcordeTerçaMenor
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Tétrades com sétima – " & TonalidadeAcorde & " " & BaixoDoAcordeTerça & " " & BaixoDoAcordeQuinta & " " & BaixoDoAcordeSétima OrElse NumeraçãoFamiliaAcorde = 68 Then

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 5) = "1T"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 3, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7/" & BaixoDoAcordeTerça

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 4) = "T"
                ArrayAcordePadrão(Index, 2, 6) = "2"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 6) = "T"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7/" & BaixoDoAcordeSétima
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 3, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 5
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 0, 4) = "0"
                ArrayAcordePadrão(Index, 1, 2) = "1"
                ArrayAcordePadrão(Index, 1, 5) = "2T"
                ArrayAcordePadrão(Index, 2, 3) = "3"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7/" & BaixoDoAcordeSétima

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 5
                ArrayAcordePadrão(Index, 5, 3) = "3"
                ArrayAcordePadrão(Index, 5, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "1T"
                ArrayAcordePadrão(Index, 3, 2) = "2"
                ArrayAcordePadrão(Index, 4, 3) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 2, 2) = "2"
                ArrayAcordePadrão(Index, 2, 3) = "3T"
                ArrayAcordePadrão(Index, 3, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste9"

                Index += 5
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 3) = "1"
                ArrayAcordePadrão(Index, 3, 1) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7/" & BaixoDoAcordeQuinta

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 4 : ArrayAcordePadrão(Index, 1, 4) = "T"
                ArrayAcordePadrão(Index, 2, 1) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7/" & BaixoDoAcordeSétima
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 4
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                   ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 3, 1) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 2) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Tétrades com sétima e quarta – " & TonalidadeAcorde & " " & BaixoDoAcordeQuarta & " " & BaixoDoAcordeQuinta & " " & BaixoDoAcordeSétima OrElse NumeraçãoFamiliaAcorde = 69 Then

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 5) = "1T"
                ArrayAcordePadrão(Index, 3, 3) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 3, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7(4)/" & BaixoDoAcordeQuarta

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 4) = "T"
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 2, 6) = "3"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7(4)/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 6) = "T"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7(4)/" & BaixoDoAcordeSétima
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7(4)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 5
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                    ArrayAcordePadrão(Index, 3, 3) = "3"
                Else
                    ArrayAcordePadrão(Index, 3, 3) = "4"
                End If
                ArrayAcordePadrão(Index, 0, 4) = "0"
                ArrayAcordePadrão(Index, 1, 2) = "1"
                ArrayAcordePadrão(Index, 1, 5) = "2T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7(4)/" & BaixoDoAcordeSétima

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 5
                ArrayAcordePadrão(Index, 5, 3) = "3"
                ArrayAcordePadrão(Index, 6, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7(4)"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "1T"
                ArrayAcordePadrão(Index, 4, 2) = "2"
                ArrayAcordePadrão(Index, 4, 3) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7(4)/" & BaixoDoAcordeQuarta
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 5 : ArrayAcordePadrão(Index, 1, 3) = "T"
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7(4)/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 5
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 1) = "P" : LinhaFinalDaSeta(Index) = 4 : ArrayAcordePadrão(Index, 3, 2) = "T"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7(4)/" & BaixoDoAcordeQuinta

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 4 : ArrayAcordePadrão(Index, 1, 4) = "T"
                ArrayAcordePadrão(Index, 2, 1) = "2"
                ArrayAcordePadrão(Index, 4, 2) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7(4)/" & BaixoDoAcordeSétima
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 4
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7(4)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 3, 4) = "2"
                ArrayAcordePadrão(Index, 4, 1) = "3"
                ArrayAcordePadrão(Index, 4, 2) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7(4)/" & BaixoDoAcordeQuarta
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Tétrades com sétima e quinta diminuta – " & TonalidadeAcorde & " " & BaixoDoAcordeTerça & " " & BaixoDoAcordeQuintaDiminuida & " " & BaixoDoAcordeSétima OrElse NumeraçãoFamiliaAcorde = 70 Then

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 5) = "1T"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 2, 6) = "3"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7(b5)/" & BaixoDoAcordeTerça

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 4, 3) = "1"
                ArrayAcordePadrão(Index, 5, 4) = "2T"
                ArrayAcordePadrão(Index, 5, 5) = "3"
               ArrayAcordePadrão(Index, 6, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7(b5)/" & BaixoDoAcordeQuintaDiminuida

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 5) = "1"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 2, 6) = "3T"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7(b5)/" & BaixoDoAcordeSétima
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 2, 5) = "3"
                ArrayAcordePadrão(Index, 3, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7(b5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 5
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "2"
                ArrayAcordePadrão(Index, 3, 5) = "3T"
                ArrayAcordePadrão(Index, 4, 3) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7(b5)/" & BaixoDoAcordeSétima
                ArrayAcordePadrão(Index, 1, 0) = "Traste11"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 5
                ArrayAcordePadrão(Index, 4, 3) = "2"
                ArrayAcordePadrão(Index, 5, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7(b5)"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "1T"
                ArrayAcordePadrão(Index, 3, 2) = "2"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 4, 3) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7(b5)/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                ArrayAcordePadrão(Index, 1, 2) = "1"
                ArrayAcordePadrão(Index, 1, 4) = "2"
                ArrayAcordePadrão(Index, 2, 3) = "3T"
                ArrayAcordePadrão(Index, 3, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7(b5)/" & BaixoDoAcordeQuintaDiminuida
                ArrayAcordePadrão(Index, 1, 0) = "Traste9"

                Index += 5
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 1) = "1"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7(b5)/" & BaixoDoAcordeQuintaDiminuida

                Index += 1
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 2, 4) = "2T"
                ArrayAcordePadrão(Index, 3, 1) = "3"
                ArrayAcordePadrão(Index, 4, 2) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7(b5)/" & BaixoDoAcordeSétima
                ArrayAcordePadrão(Index, 1, 0) = "Traste4"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "1T"
                ArrayAcordePadrão(Index, 1, 3) = "2"
                ArrayAcordePadrão(Index, 2, 2) = "3"
                ArrayAcordePadrão(Index, 2, 4) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7(b5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 1) = "3"
                ArrayAcordePadrão(Index, 4, 2) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7(b5)/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Tétrades com sétima e quinta aumentada – " & TonalidadeAcorde & " " & BaixoDoAcordeTerça & " " & BaixoDoAcordeQuintaAumentada & " " & BaixoDoAcordeSétima OrElse NumeraçãoFamiliaAcorde = 71 Then

                StringSubstituição(0) = "b6" : StringSubstituição(1) = "#5" : SubstituirString()

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 5) = "1T"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7(#5)/" & BaixoDoAcordeTerça

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 5, 4) = "1T"
                ArrayAcordePadrão(Index, 5, 5) = "2"
                ArrayAcordePadrão(Index, 6, 3) = "3"
                ArrayAcordePadrão(Index, 6, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7(#5)/" & BaixoDoAcordeQuintaAumentada

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 1, 6) = "2T"
                ArrayAcordePadrão(Index, 2, 4) = "3"
                ArrayAcordePadrão(Index, 2, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7(#5)/" & BaixoDoAcordeSétima
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
               End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 3, 6) = "3"
                ArrayAcordePadrão(Index, 4, 4) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7(#5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 5
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 5 : ArrayAcordePadrão(Index, 1, 5) = "T"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7(#5)/" & BaixoDoAcordeSétima

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 5
                ArrayAcordePadrão(Index, 5, 5) = "3"
                ArrayAcordePadrão(Index, 6, 3) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7(#5)"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "1T"
                ArrayAcordePadrão(Index, 3, 2) = "2"
                ArrayAcordePadrão(Index, 4, 3) = "3"
                ArrayAcordePadrão(Index, 5, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7(#5)/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 2, 3) = "2T"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 3, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7(#5)/" & BaixoDoAcordeQuintaAumentada
                ArrayAcordePadrão(Index, 1, 0) = "Traste9"

                Index += 5
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 3) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 1) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7(#5)/" & BaixoDoAcordeQuintaAumentada

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "1T"
                ArrayAcordePadrão(Index, 2, 1) = "2"
                ArrayAcordePadrão(Index, 2, 3) = "3"
                ArrayAcordePadrão(Index, 3, 2) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7(#5)/" & BaixoDoAcordeSétima
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 4
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 4, 2) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7(#5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 3, 1) = "2"
                ArrayAcordePadrão(Index, 4, 2) = "3"
                ArrayAcordePadrão(Index, 4, 4) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7(#5)/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Tétrades com sétima maior – " & TonalidadeAcorde & " " & BaixoDoAcordeTerça & " " & BaixoDoAcordeQuinta & " " & BaixoDoAcordeSétimaMaior OrElse NumeraçãoFamiliaAcorde = 72 Then

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 5) = "1T"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 3, 6) = "3"
                ArrayAcordePadrão(Index, 4, 4) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M/" & BaixoDoAcordeTerça

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 4) = "T"
                ArrayAcordePadrão(Index, 3, 6) = "3"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 5) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 6) = "T"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 2, 4) = "3"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M/" & BaixoDoAcordeSétimaMaior
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 3, 4) = "2"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 3, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7M"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 5
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 0, 4) = "0"
                ArrayAcordePadrão(Index, 1, 5) = "1T"
                ArrayAcordePadrão(Index, 2, 2) = "2"
                ArrayAcordePadrão(Index, 2, 3) = "3"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M/" & BaixoDoAcordeSétimaMaior

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 2) = "1T"
                ArrayAcordePadrão(Index, 4, 4) = "2"
                ArrayAcordePadrão(Index, 5, 3) = "3"
                ArrayAcordePadrão(Index, 5, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7M"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "1T"
                ArrayAcordePadrão(Index, 3, 2) = "2"
                ArrayAcordePadrão(Index, 4, 5) = "3"
                ArrayAcordePadrão(Index, 5, 3) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 2, 2) = "2"
                ArrayAcordePadrão(Index, 2, 3) = "3T"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste9"

                Index += 5
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 3) = "1"
                ArrayAcordePadrão(Index, 3, 1) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 4, 4) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M/" & BaixoDoAcordeQuinta

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 4 : ArrayAcordePadrão(Index, 1, 4) = "T"
                ArrayAcordePadrão(Index, 3, 1) = "3"
                ArrayAcordePadrão(Index, 3, 2) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M/" & BaixoDoAcordeSétimaMaior
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "1T"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 2, 4) = "3"
                ArrayAcordePadrão(Index, 3, 2) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7M"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 3, 1) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 5, 2) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Tétrades com sétima maior e quinta diminuta – " & TonalidadeAcorde & " " & BaixoDoAcordeTerça & " " & BaixoDoAcordeQuintaDiminuida & " " & BaixoDoAcordeSétimaMaior OrElse NumeraçãoFamiliaAcorde = 73 Then

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 5) = "1T"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 2, 6) = "3"
                ArrayAcordePadrão(Index, 4, 4) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(b5)/" & BaixoDoAcordeTerça

                Index += 1
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 2, 4) = "2T"
                ArrayAcordePadrão(Index, 2, 5) = "3"
                ArrayAcordePadrão(Index, 4, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(b5)/" & BaixoDoAcordeQuintaDiminuida
                ArrayAcordePadrão(Index, 1, 0) = "Traste4"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 5) = "1"
                ArrayAcordePadrão(Index, 2, 6) = "2T"
                ArrayAcordePadrão(Index, 3, 3) = "3"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(b5)/" & BaixoDoAcordeSétimaMaior
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 3, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(b5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 5
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 3, 5) = "2T"
                ArrayAcordePadrão(Index, 4, 2) = "3"
                ArrayAcordePadrão(Index, 4, 3) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(b5)/" & BaixoDoAcordeSétimaMaior
                ArrayAcordePadrão(Index, 1, 0) = "Traste11"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 2) = "1T"
                ArrayAcordePadrão(Index, 4, 3) = "2"
                ArrayAcordePadrão(Index, 4, 4) = "3"
                ArrayAcordePadrão(Index, 5, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(b5)"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "1T"
                ArrayAcordePadrão(Index, 3, 2) = "2"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 5, 3) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(b5)/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                ArrayAcordePadrão(Index, 1, 2) = "P" : LinhaFinalDaSeta(Index) = 4
                ArrayAcordePadrão(Index, 2, 3) = "2T"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(b5)/" & BaixoDoAcordeQuintaDiminuida
                ArrayAcordePadrão(Index, 1, 0) = "Traste9"

                Index += 5
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 1) = "P" : LinhaFinalDaSeta(Index) = 4
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 4, 4) = "3"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(b5)/" & BaixoDoAcordeQuintaDiminuida

                Index += 1
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 2, 4) = "2T"
                ArrayAcordePadrão(Index, 4, 1) = "3"
               ArrayAcordePadrão(Index, 4, 2) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(b5)/" & BaixoDoAcordeSétimaMaior
                ArrayAcordePadrão(Index, 1, 0) = "Traste4"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "1T"
                ArrayAcordePadrão(Index, 2, 2) = "2"
                ArrayAcordePadrão(Index, 2, 3) = "3"
                ArrayAcordePadrão(Index, 2, 4) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(b5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 3, 1) = "3"
                ArrayAcordePadrão(Index, 5, 2) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(b5)/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Tétrades com sétima maior e quinta aumentada – " & TonalidadeAcorde & " " & BaixoDoAcordeTerça & " " & BaixoDoAcordeQuintaAumentada & " " & BaixoDoAcordeSétimaMaior OrElse NumeraçãoFamiliaAcorde = 74 Then

                StringSubstituição(0) = "b6" : StringSubstituição(1) = "#5" : SubstituirString()

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 5) = "1T"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 4, 4) = "3"
                ArrayAcordePadrão(Index, 4, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(#5)/" & BaixoDoAcordeTerça

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "PT" : LinhaFinalDaSeta(Index) = 5
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 3, 6) = "3"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(#5)/" & BaixoDoAcordeQuintaAumentada
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 6) = "1T"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 2, 4) = "3"
                ArrayAcordePadrão(Index, 2, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(#5)/" & BaixoDoAcordeSétimaMaior
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 3, 5) = "2"
                ArrayAcordePadrão(Index, 3, 6) = "3"
                ArrayAcordePadrão(Index, 4, 4) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(#5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 5
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "P" : LinhaFinalDaSeta(Index) = 5 : ArrayAcordePadrão(Index, 1, 5) = "T"
                ArrayAcordePadrão(Index, 2, 2) = "2"
                ArrayAcordePadrão(Index, 2, 3) = "3"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(#5)/" & BaixoDoAcordeSétimaMaior

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 2) = "1T"
                ArrayAcordePadrão(Index, 4, 4) = "2"
                ArrayAcordePadrão(Index, 5, 5) = "3"
                ArrayAcordePadrão(Index, 6, 3) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(#5)"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "1T"
                ArrayAcordePadrão(Index, 3, 2) = "2"
                ArrayAcordePadrão(Index, 5, 3) = "3"
                ArrayAcordePadrão(Index, 5, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(#5)/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 2, 3) = "2T"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(#5)/" & BaixoDoAcordeQuintaAumentada
                ArrayAcordePadrão(Index, 1, 0) = "Traste9"

                Index += 5
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 2, 3) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "2T"
                ArrayAcordePadrão(Index, 4, 1) = "3"
                ArrayAcordePadrão(Index, 4, 4) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(#5)/" & BaixoDoAcordeQuintaAumentada

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "1T"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 3, 1) = "3"
                ArrayAcordePadrão(Index, 3, 2) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(#5)/" & BaixoDoAcordeSétimaMaior
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "1T"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 2, 4) = "3"
                ArrayAcordePadrão(Index, 4, 2) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(#5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 3, 1) = "2"
                ArrayAcordePadrão(Index, 4, 4) = "3"
                ArrayAcordePadrão(Index, 5, 2) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "7M(#5)/" & BaixoDoAcordeTerça
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Tétrades menores com sétima maior – " & TonalidadeAcorde & " " & BaixoDoAcordeTerçaMenor & " " & BaixoDoAcordeQuinta & " " & BaixoDoAcordeSétimaMaior OrElse NumeraçãoFamiliaAcorde = 75 Then

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 5) = "T"
                ArrayAcordePadrão(Index, 3, 6) = "3"
                ArrayAcordePadrão(Index, 4, 4) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "m(7M)/" & BaixoDoAcordeTerçaMenor

                Index += 1
                ArrayAcordePadrão(Index, 1, 5) = "1"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 2, 4) = "3T"
                ArrayAcordePadrão(Index, 4, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "m(7M)/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste4"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 6) = "T"
                ArrayAcordePadrão(Index, 2, 3) = "3"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "m(7M)/" & BaixoDoAcordeSétimaMaior
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 2, 6) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 3, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "m(7M)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 5
                If TonalidadeAcorde <> "C" Then
                   ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 0, 4) = "0"
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 1, 5) = "2T"
                ArrayAcordePadrão(Index, 2, 2) = "3"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "m(7M)/" & BaixoDoAcordeSétimaMaior

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 2) = "1T"
                ArrayAcordePadrão(Index, 4, 4) = "2"
                ArrayAcordePadrão(Index, 4, 5) = "3"
                ArrayAcordePadrão(Index, 5, 3) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "m(7M)"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "1T"
                ArrayAcordePadrão(Index, 2, 2) = "2"
                ArrayAcordePadrão(Index, 4, 5) = "3"
                ArrayAcordePadrão(Index, 5, 3) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "m(7M)/" & BaixoDoAcordeTerçaMenor
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3T"
                ArrayAcordePadrão(Index, 5, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "m(7M)/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 5
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 3, 1) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 4, 4) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "m(7M)/" & BaixoDoAcordeQuinta

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 4 : ArrayAcordePadrão(Index, 1, 4) = "T"
                ArrayAcordePadrão(Index, 2, 2) = "2"
                ArrayAcordePadrão(Index, 3, 1) = "3"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "m(7M)/" & BaixoDoAcordeSétimaMaior
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 4
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ImagemAcorde(Index) = TonalidadeAcorde & "m(7M)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 2, 1) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 5, 2) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ImagemAcorde(Index) = TonalidadeAcorde & "m(7M)/" & BaixoDoAcordeTerçaMenor
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim"

            ElseIf FamiliaAcorde = "Tétrades menores com sétima – " & TonalidadeAcorde & " " & BaixoDoAcordeTerçaMenor & " " & BaixoDoAcordeQuinta & " " & BaixoDoAcordeSétima OrElse NumeraçãoFamiliaAcorde = 76 Then

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "PS" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 5) = "T"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 3, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6" 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7/" & BaixoDoAcordeTerçaMenor

                Index += 1
                ArrayAcordePadrão(Index, 1, 5) = "1S"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 2, 4) = "3T"
                ArrayAcordePadrão(Index, 3, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6/" & BaixoDoAcordeQuinta 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste4"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 4) = "S" : ArrayAcordePadrão(Index, 1, 6) = "T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6/" & BaixoDoAcordeSétima 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7/" & BaixoDoAcordeSétima
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 2, 5) = "2"
                ArrayAcordePadrão(Index, 2, 6) = "3S"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 5
                If TonalidadeAcorde <> "C" Then
                    ArrayAcordePadrão(Index, 8, 2) = "15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 0, 4) = "1"
                ArrayAcordePadrão(Index, 1, 2) = "2"
                ArrayAcordePadrão(Index, 1, 3) = "3S"
                ArrayAcordePadrão(Index, 1, 5) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6/" & BaixoDoAcordeSétima 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7/" & BaixoDoAcordeSétima

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 2) = "PT" : LinhaFinalDaSeta(Index) = 5
                ArrayAcordePadrão(Index, 4, 5) = "2S"
                ArrayAcordePadrão(Index, 5, 3) = "3"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "1T"
                ArrayAcordePadrão(Index, 2, 2) = "2S"
                ArrayAcordePadrão(Index, 4, 3) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6" 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7/" & BaixoDoAcordeTerçaMenor
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "1S"
                ArrayAcordePadrão(Index, 3, 2) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3T"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6/" & BaixoDoAcordeQuinta 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7/" & BaixoDoAcordeQuinta
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 5
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1S"
                ArrayAcordePadrão(Index, 3, 1) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6/" & BaixoDoAcordeQuinta 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7/" & BaixoDoAcordeQuinta

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 4 : ArrayAcordePadrão(Index, 1, 4) = "T"
                ArrayAcordePadrão(Index, 2, 1) = "2"
                ArrayAcordePadrão(Index, 2, 2) = "3S"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6/" & BaixoDoAcordeSétima 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7/" & BaixoDoAcordeSétima
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 4 : ArrayAcordePadrão(Index, 1, 4) = "S"
                ArrayAcordePadrão(Index, 3, 2) = "3"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 2, 1) = "2S"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 4, 2) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "6" 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7/" & BaixoDoAcordeTerçaMenor
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            ElseIf FamiliaAcorde = "Tétrades menores com sétima e quinta diminuta – " & TonalidadeAcorde & " " & BaixoDoAcordeTerçaMenor & " " & BaixoDoAcordeQuintaDiminuida & " " & BaixoDoAcordeSétima OrElse NumeraçãoFamiliaAcorde = 77 Then

                Index += 1
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "PS" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 5) = "T"
                ArrayAcordePadrão(Index, 2, 6) = "2"
                ArrayAcordePadrão(Index, 3, 4) = "3"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "m6" 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7(b5)/" & BaixoDoAcordeTerçaMenor

                Index += 1
                ArrayAcordePadrão(Index, 1, 3) = "P" : LinhaFinalDaSeta(Index) = 6 : ArrayAcordePadrão(Index, 1, 5) = "S"
                ArrayAcordePadrão(Index, 2, 4) = "2T"
                ArrayAcordePadrão(Index, 3, 6) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "m6/" & BaixoDoAcordeQuintaDiminuida 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7(b5)/" & BaixoDoAcordeQuintaDiminuida
                ArrayAcordePadrão(Index, 1, 0) = "Traste4"

                Index += 1
                If TonalidadeAcorde = "F" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 5) = "1"
                ArrayAcordePadrão(Index, 2, 3) = "2"
                ArrayAcordePadrão(Index, 2, 4) = "3S"
                ArrayAcordePadrão(Index, 2, 6) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "m6/" & BaixoDoAcordeSétima 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7(b5)/" & BaixoDoAcordeSétima
                ArrayAcordePadrão(Index, 1, 0) = "Traste7"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 2, 4) = "2"
                ArrayAcordePadrão(Index, 2, 5) = "3"
                ArrayAcordePadrão(Index, 2, 6) = "4S"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 2) = ""
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "m6/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7(b5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 5
                ArrayAcordePadrão(Index, 1, 4) = "1"
                ArrayAcordePadrão(Index, 3, 2) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3S"
                ArrayAcordePadrão(Index, 3, 5) = "4T"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "m6/" & BaixoDoAcordeSétima 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7(b5)/" & BaixoDoAcordeSétima
                ArrayAcordePadrão(Index, 1, 0) = "Traste11"

                Index += 1
                If TonalidadeAcorde = "A" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-45" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                ElseIf TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                End If
                ArrayAcordePadrão(Index, 3, 2) = "1T"
                ArrayAcordePadrão(Index, 3, 4) = "2"
                ArrayAcordePadrão(Index, 4, 3) = "3"
                ArrayAcordePadrão(Index, 4, 5) = "4S"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "m6/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7(b5)"

                Index += 1
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "1T"
                ArrayAcordePadrão(Index, 2, 2) = "2S"
                ArrayAcordePadrão(Index, 3, 5) = "3"
                ArrayAcordePadrão(Index, 4, 3) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "m6" 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7(b5)/" & BaixoDoAcordeTerçaMenor
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "1S"
                ArrayAcordePadrão(Index, 2, 2) = "2"
                ArrayAcordePadrão(Index, 3, 3) = "3T"
                ArrayAcordePadrão(Index, 4, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "m6/" & BaixoDoAcordeQuintaDiminuida 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7(b5)/" & BaixoDoAcordeQuintaDiminuida
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 5
                If TonalidadeAcorde = "B" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1S"
                ArrayAcordePadrão(Index, 2, 1) = "2"
                ArrayAcordePadrão(Index, 3, 2) = "3T"
                ArrayAcordePadrão(Index, 3, 4) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "m6/" & BaixoDoAcordeQuintaDiminuida 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7(b5)/" & BaixoDoAcordeQuintaDiminuida

                Index += 1
                ArrayAcordePadrão(Index, 1, 3) = "1"
                ArrayAcordePadrão(Index, 2, 4) = "2T"
                ArrayAcordePadrão(Index, 3, 1) = "3"
                ArrayAcordePadrão(Index, 3, 2) = "4S"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "m6/" & BaixoDoAcordeSétima 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7(b5)/" & BaixoDoAcordeSétima
                ArrayAcordePadrão(Index, 1, 0) = "Traste4"

                Index += 1
                If TonalidadeAcorde = "E" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 1) = "PT" : LinhaFinalDaSeta(Index) = 4 : ArrayAcordePadrão(Index, 1, 4) = "S"
                ArrayAcordePadrão(Index, 2, 2) = "2"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "m6/" & TonalidadeAcorde 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7(b5)"
                ArrayAcordePadrão(Index, 1, 0) = "Traste8"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 2, 1) = "2S"
                ArrayAcordePadrão(Index, 2, 4) = "3"
                ArrayAcordePadrão(Index, 4, 2) = "4"
                ArrayAcordePadrão(Index, 7, 5) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "m6" 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7(b5)/" & BaixoDoAcordeTerçaMenor
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 5
                If TonalidadeAcorde = "G" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 4) = "1T"
                ArrayAcordePadrão(Index, 2, 2) = "2S"
                ArrayAcordePadrão(Index, 2, 6) = "3"
                ArrayAcordePadrão(Index, 3, 5) = "4"
                ArrayAcordePadrão(Index, 7, 1) = "" : ArrayAcordePadrão(Index, 7, 3) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "m6" 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7(b5)/" & BaixoDoAcordeTerçaMenor
                ArrayAcordePadrão(Index, 1, 0) = "Traste5"

                Index += 1
                If TonalidadeAcorde = "D" Then
                    ArrayAcordePadrão(Index, 8, 2) = "-15" 'Indica o quanto este acorde deve ser deslocado para cima ou para baixo
                    ArrayAcordePadrão(Index, 8, 3) = "-1" 'Indica o quanto os números que ilustram o dedilhado devem aumentar ou diminuir
                End If
                ArrayAcordePadrão(Index, 1, 3) = "1T"
                ArrayAcordePadrão(Index, 2, 1) = "2S"
                ArrayAcordePadrão(Index, 2, 4) = "3"
                ArrayAcordePadrão(Index, 2, 5) = "4"
                ArrayAcordePadrão(Index, 7, 2) = "" : ArrayAcordePadrão(Index, 7, 6) = ""
                ArrayAcordePadrão(Index, 8, 4) = "x" 'indica que o acorde é invertido
                ArrayAcordePadrão(Index, 8, 5) = "= " & BaixoDoAcordeTerçaMenor & "m6" 'indica o nome do acorde sinônimo
                ImagemAcorde(Index) = TonalidadeAcorde & "m7(b5)/" & BaixoDoAcordeTerçaMenor
                ArrayAcordePadrão(Index, 1, 0) = "Traste10"

                Index += 1
                ArrayAcordePadrão(Index, 0, 0) = "Fim" 'após o último acorde, isso fará com que o loop seja finalizado

            End If


            Desenhar()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click
        Me.WindowState = CType(1, FormWindowState) '1 é para minimizar
    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click
        Me.Close()
    End Sub

    Private Sub SubstituirString()

        Try

            For a = 0 To 27
                For b = 0 To 6
                    If IntervalosAcorde(a, b) = StringSubstituição(0) Then IntervalosAcorde(a, b) = StringSubstituição(1)
                Next
            Next

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub Label3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label3.Click

        Try

            If My.Settings.NovoValorDesenharBolinhasNotasAcordes = True Then
                My.Settings.NovoValorDesenharBolinhasNotasAcordes = False
            Else
                My.Settings.NovoValorDesenharBolinhasNotasAcordes = True
            End If
            Desenhar()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Label4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label4.Click

        Try

            If My.Settings.NovoValorDesenharAcordesMaisUsados = True Then
                My.Settings.NovoValorDesenharAcordesMaisUsados = False
            Else
                My.Settings.NovoValorDesenharAcordesMaisUsados = True
            End If
            Desenhar()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Label5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label5.Click

        Try

            If My.Settings.NovoValorLocalizaçãoFundamentalDeReferenciaNoAcorde = True Then
                My.Settings.NovoValorLocalizaçãoFundamentalDeReferenciaNoAcorde = False
            Else
                My.Settings.NovoValorLocalizaçãoFundamentalDeReferenciaNoAcorde = True
            End If
            Desenhar()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Label8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label8.Click
        ExibirNomeIntervalos()
    End Sub

    Private Sub ExibirNomeIntervalos()
        Try

            If My.Settings.NovoValorExibirNomeIntervalos = True Then
                My.Settings.NovoValorExibirNomeIntervalos = False
            Else
                My.Settings.NovoValorExibirNomeIntervalos = True
                My.Settings.NovoValorExibirNomeDasNotasAcordes = False
            End If
            Desenhar()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Label9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label9.Click

        Try

            If My.Settings.NovoValorCopiarImagemDosAcordes = True Then
                My.Settings.NovoValorCopiarImagemDosAcordes = False
            Else
                My.Settings.NovoValorCopiarImagemDosAcordes = True
            End If
            Desenhar()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Label6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label6.Click
        Anterior()
    End Sub

    Private Sub Label7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label7.Click
        Proximo()
    End Sub

    Private Sub Proximo()
        ImagemCopiada(0) = Nothing : ImagemCopiada(1) = Nothing : ImagemCopiada(2) = Nothing
        FamiliaAcorde = ""
        NumeraçãoFamiliaAcorde += 1
        If NumeraçãoFamiliaAcorde > 77 Then NumeraçãoFamiliaAcorde = 1
        LocalizaNomeFamiliaAcordes()
        RedefiniçãoTonalidade()
    End Sub

    Private Sub Anterior()
        ImagemCopiada(0) = Nothing : ImagemCopiada(1) = Nothing : ImagemCopiada(2) = Nothing
        FamiliaAcorde = ""
        NumeraçãoFamiliaAcorde -= 1
        If NumeraçãoFamiliaAcorde < 1 Then NumeraçãoFamiliaAcorde = 77
        LocalizaNomeFamiliaAcordes()
        RedefiniçãoTonalidade()
    End Sub

    Private Sub RedefiniçãoTonalidade()

        Try
            'executa as subrotinas das "tonalidades", para redefinição das variáveis AjustePosiçãoTonalidadeDoAcorde e NotasAcordeIndiceLinha
            'as quais são alteradas quando se entra nos quatro primeiros itens do menu "Pedagogia do Violão"

            If TonalidadeAcorde = "C" Then
                TonalidadeC()
            ElseIf TonalidadeAcorde = "D" Then
                TonalidadeD()
            ElseIf TonalidadeAcorde = "E" Then
                TonalidadeE()
            ElseIf TonalidadeAcorde = "F" Then
                TonalidadeF()
            ElseIf TonalidadeAcorde = "G" Then
                TonalidadeG()
            ElseIf TonalidadeAcorde = "A" Then
                TonalidadeA()
            ElseIf TonalidadeAcorde = "B" Then
                TonalidadeB()
            End If

        Catch ex As Exception
            MsgBox(ex.ToString)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub LocalizaNomeFamiliaAcordes()

        Try

            If NumeraçãoFamiliaAcorde >= 1 AndAlso NumeraçãoFamiliaAcorde <= 13 Then

                For Each Item As ToolStripMenuItem In MenuStrip1.Items
                    If TypeOf (Item) Is ToolStripMenuItem Then 'ignora tudo o que não for ToolStripMenuItem
                        If CInt(Item.Tag) = NumeraçãoFamiliaAcorde Then
                            FamiliaAcorde = Item.Text
                        End If
                        subLocalizaNomeFamiliaAcordes(Item)
                    End If
                Next

            ElseIf NumeraçãoFamiliaAcorde >= 14 AndAlso NumeraçãoFamiliaAcorde <= 28 Then

                For Each Item As ToolStripMenuItem In MenuStrip5.Items
                    If TypeOf (Item) Is ToolStripMenuItem Then 'ignora tudo o que não for ToolStripMenuItem
                        If CInt(Item.Tag) = NumeraçãoFamiliaAcorde Then
                            FamiliaAcorde = Item.Text
                        End If
                        subLocalizaNomeFamiliaAcordes(Item)
                    End If
                Next

            ElseIf NumeraçãoFamiliaAcorde = 29 Then

                For Each Item As ToolStripMenuItem In MenuStrip3.Items
                    If TypeOf (Item) Is ToolStripMenuItem Then 'ignora tudo o que não for ToolStripMenuItem
                        If CInt(Item.Tag) = NumeraçãoFamiliaAcorde Then
                            FamiliaAcorde = Item.Text
                        End If
                        subLocalizaNomeFamiliaAcordes(Item)
                    End If
                Next

            ElseIf NumeraçãoFamiliaAcorde >= 30 AndAlso NumeraçãoFamiliaAcorde <= 53 Then

                For Each Item As ToolStripMenuItem In MenuStrip2.Items
                    If TypeOf (Item) Is ToolStripMenuItem Then 'ignora tudo o que não for ToolStripMenuItem
                        If CInt(Item.Tag) = NumeraçãoFamiliaAcorde Then
                            FamiliaAcorde = Item.Text
                        End If
                        subLocalizaNomeFamiliaAcordes(Item)
                    End If
                Next

            ElseIf NumeraçãoFamiliaAcorde = 54 Then

                For Each Item As ToolStripMenuItem In MenuStrip4.Items
                    If TypeOf (Item) Is ToolStripMenuItem Then 'ignora tudo o que não for ToolStripMenuItem
                        If CInt(Item.Tag) = NumeraçãoFamiliaAcorde Then
                            FamiliaAcorde = Item.Text
                        End If
                        subLocalizaNomeFamiliaAcordes(Item)
                    End If
                Next

            ElseIf NumeraçãoFamiliaAcorde >= 55 AndAlso NumeraçãoFamiliaAcorde <= 77 Then

                For Each Item As ToolStripMenuItem In MenuStrip6.Items
                    If TypeOf (Item) Is ToolStripMenuItem Then 'ignora tudo o que não for ToolStripMenuItem
                        If CInt(Item.Tag) = NumeraçãoFamiliaAcorde Then
                            FamiliaAcorde = Item.Text
                        End If
                        subLocalizaNomeFamiliaAcordes(Item)
                    End If
                Next

            End If

        Catch ex As Exception
            MsgBox(ex.ToString)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Public Sub subLocalizaNomeFamiliaAcordes(ByVal ToolstripMenuItem As ToolStripMenuItem)

        Try

            'Faz a mesma coisa que a rotina do load
            For ii = 0 To ToolstripMenuItem.DropDownItems.Count - 1
                If TypeOf (ToolstripMenuItem.DropDownItems.Item(ii)) Is ToolStripMenuItem Then 'ignora tudo o que não for ToolStripMenuItem
                    If CInt(ToolstripMenuItem.DropDownItems.Item(ii).Tag) = NumeraçãoFamiliaAcorde Then
                        FamiliaAcorde = ToolstripMenuItem.DropDownItems.Item(ii).Text
                    End If
                    'Aqui é a parte mais importante. Isso se chama recursão, onde um método
                    'chama ele mesmo, e passa o próprio item que está sendo adicionado ao listbox
                    'Assim ele adiciona o item e chama a mesma rotina para adicionar os seus respectivos
                    'subItems
                    subLocalizaNomeFamiliaAcordes(CType(ToolstripMenuItem.DropDownItems.Item(ii), Windows.Forms.ToolStripMenuItem))
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.ToString)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub CopiarAcorde()

        Try

            If ImagemCopiada(0) IsNot Nothing Then

                Dim gr2 As Graphics
               Dim FaceBit2 As Bitmap
                FaceBit2 = New Bitmap(ImagemCopiada(1))
                Me.SetBitmap(FaceBit2, TransAmount)


                gr2 = Graphics.FromImage(FaceBit2)
                Rect1.DashStyle = DashStyle.Dash

                If ValorEsquerda = -1 Then
                    ImagemCopiada(2) = ImagemCopiada(0) 'clicando numa área sem acorde, então irá copiar todos os acordes
                    gr2.DrawRectangle(Rect1, 42, 165, CInt(ImagemCopiada(2).Width / FatorEscala), CInt(ImagemCopiada(2).Height / FatorEscala))
                Else
                    'copiará somente o acorde clicado
                    ImagemCopiada(2) = ImagemCopiada(0).Clone(New Rectangle(ValorEsquerda * FatorEscala, ValorTopo * FatorEscala, 96 * FatorEscala, 121 * FatorEscala), ImagemCopiada(0).PixelFormat)
                    gr2.DrawRectangle(Rect1, ValorEsquerda + 42, ValorTopo + 165, 96, 121)
                End If

                ImagemCopiada(2).SetResolution(96 * FatorEscala, 96 * FatorEscala)
                Clipboard.SetDataObject(ImagemCopiada(2))
                If My.Settings.NovoValorSalvarAcordes = True Then ImagemCopiada(2).Save("C:\Imagens Acordes Copiados\Acorde copiado - Fator " & FatorEscala & ".png", ImageFormat.Png)


                gr2.TextRenderingHint = TextRenderingHint.ClearTypeGridFit
                CenterTextAt(gr2, "Imagem Copiada.    Fator: " & FatorEscala & ".    " & ImagemCopiada(2).HorizontalResolution & " dpi.    Tamanho: " & ImagemCopiada(2).Width & "x" & ImagemCopiada(2).Height & " px.    Reduza para " & FormatNumber(100 / FatorEscala, 2) & "%", CInt(Me.Width / 2), 750)

                Me.SetBitmap(FaceBit2, TransAmount)

            End If

        Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
            ValorEsquerda = -1
            CopiarAcorde()
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Acordes_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

        Try

            ' A tecla pressionada está na nossa lista? 
            If Not lista.Contains(e.KeyCode) Then
                ' Não! A adicionamos à lista.
                lista.Add(e.KeyCode)


                If e.KeyCode = Keys.Right OrElse e.KeyCode = Keys.Down OrElse e.KeyCode = Keys.PageDown OrElse e.KeyCode = Keys.NumPad6 OrElse e.KeyCode = Keys.NumPad2 Then
                    Proximo()
                ElseIf e.KeyCode = Keys.Left OrElse e.KeyCode = Keys.Up OrElse e.KeyCode = Keys.PageUp OrElse e.KeyCode = Keys.NumPad8 OrElse e.KeyCode = Keys.NumPad4 Then
                    Anterior()
                ElseIf e.KeyCode = Keys.Add Then
                    My.Settings.NovoValorFatorEscala += 1
                    GerarAcordes()
                ElseIf e.KeyCode = Keys.Subtract Then
                    If My.Settings.NovoValorFatorEscala > 1 Then My.Settings.NovoValorFatorEscala -= 1
                    GerarAcordes()
                ElseIf e.Modifiers = Nothing Then
                    If e.KeyCode = Keys.C Then
                        TonalidadeC()
                    ElseIf e.KeyCode = Keys.D Then
                        TonalidadeD()
                    ElseIf e.KeyCode = Keys.E Then
                        TonalidadeE()
                    ElseIf e.KeyCode = Keys.F Then
                        TonalidadeF()
                    ElseIf e.KeyCode = Keys.G Then
                        TonalidadeG()
                    ElseIf e.KeyCode = Keys.A Then
                        TonalidadeA()
                    ElseIf e.KeyCode = Keys.B Then
                        TonalidadeB()
                    ElseIf e.KeyCode = Keys.I Then
                        ExibirNomeIntervalos()
                    ElseIf e.KeyCode = Keys.N Then
                        ExibirNomeDasNotasAcordes()
                    ElseIf e.KeyCode = Keys.V Then
                        ExibirNotasBraçoViolão()
                    ElseIf e.KeyCode = Keys.M Then
                        ExibirIntervalosBraçoViolão()
                    ElseIf e.KeyCode = Keys.J Then
                        TelaAcordesModoJogo.Show()
                        GerarNovoAcordeJogo()
                    End If
                ElseIf e.Modifiers = Keys.Control Then
                    If e.KeyCode = Keys.F1 Then
                        ValorEsquerda = 0
                        ValorTopo = 0
                    ElseIf e.KeyCode = Keys.F2 Then
                        ValorEsquerda = 128
                        ValorTopo = 0
                    ElseIf e.KeyCode = Keys.F3 Then
                        ValorEsquerda = 256
                        ValorTopo = 0
                    ElseIf e.KeyCode = Keys.F4 Then
                        ValorEsquerda = 384
                        ValorTopo = 0
                    ElseIf e.KeyCode = Keys.F5 Then
                        ValorEsquerda = 512
                        ValorTopo = 0
                    ElseIf e.KeyCode = Keys.F6 Then
                        ValorEsquerda = 640
                        ValorTopo = 0
                    ElseIf e.KeyCode = Keys.F7 Then
                        ValorEsquerda = 768
                        ValorTopo = 0
                    ElseIf e.KeyCode = Keys.F8 Then
                        ValorEsquerda = 896
                        ValorTopo = 0


                    ElseIf e.KeyCode = Keys.D1 Then
                        ValorEsquerda = 0
                        ValorTopo = 118
                    ElseIf e.KeyCode = Keys.D2 Then
                        ValorEsquerda = 128
                        ValorTopo = 118
                    ElseIf e.KeyCode = Keys.D3 Then
                        ValorEsquerda = 256
                        ValorTopo = 118
                    ElseIf e.KeyCode = Keys.D4 Then
                        ValorEsquerda = 384
                        ValorTopo = 118
                    ElseIf e.KeyCode = Keys.D5 Then
                        ValorEsquerda = 512
                        ValorTopo = 118
                    ElseIf e.KeyCode = Keys.D6 Then
                        ValorEsquerda = 640
                        ValorTopo = 118
                    ElseIf e.KeyCode = Keys.D7 Then
                        ValorEsquerda = 768
                        ValorTopo = 118
                    ElseIf e.KeyCode = Keys.D8 Then
                        ValorEsquerda = 896
                        ValorTopo = 118


                    ElseIf e.KeyCode = Keys.Q Then
                        ValorEsquerda = 0
                        ValorTopo = 236
                    ElseIf e.KeyCode = Keys.W Then
                        ValorEsquerda = 128
                        ValorTopo = 236
                    ElseIf e.KeyCode = Keys.E Then
                        ValorEsquerda = 256
                        ValorTopo = 236
                    ElseIf e.KeyCode = Keys.R Then
                        ValorEsquerda = 384
                        ValorTopo = 236
                    ElseIf e.KeyCode = Keys.T Then
                        ValorEsquerda = 512
                        ValorTopo = 236
                    ElseIf e.KeyCode = Keys.Y Then
                        ValorEsquerda = 640
                        ValorTopo = 236
                    ElseIf e.KeyCode = Keys.U Then
                        ValorEsquerda = 768
                        ValorTopo = 236
                    ElseIf e.KeyCode = Keys.I Then
                        ValorEsquerda = 896
                        ValorTopo = 236


                    ElseIf e.KeyCode = Keys.A Then
                        ValorEsquerda = 0
                        ValorTopo = 354
                    ElseIf e.KeyCode = Keys.S Then
                        ValorEsquerda = 128
                        ValorTopo = 354
                    ElseIf e.KeyCode = Keys.D Then
                        ValorEsquerda = 256
                        ValorTopo = 354
                    ElseIf e.KeyCode = Keys.F Then
                        ValorEsquerda = 384
                        ValorTopo = 354
                    ElseIf e.KeyCode = Keys.G Then
                        ValorEsquerda = 512
                        ValorTopo = 354
                    ElseIf e.KeyCode = Keys.H Then
                        ValorEsquerda = 640
                        ValorTopo = 354
                    ElseIf e.KeyCode = Keys.J Then
                        ValorEsquerda = 768
                        ValorTopo = 354
                    ElseIf e.KeyCode = Keys.K Then
                        ValorEsquerda = 896
                        ValorTopo = 354


                    ElseIf e.KeyCode = Keys.Z Then
                        ValorEsquerda = 0
                        ValorTopo = 472
                    ElseIf e.KeyCode = Keys.X Then
                        ValorEsquerda = 128
                        ValorTopo = 472
                    ElseIf e.KeyCode = Keys.C Then
                        ValorEsquerda = 256
                        ValorTopo = 472
                    ElseIf e.KeyCode = Keys.V Then
                        ValorEsquerda = 384
                        ValorTopo = 472
                    ElseIf e.KeyCode = Keys.B Then
                        ValorEsquerda = 512
                        ValorTopo = 472
                    ElseIf e.KeyCode = Keys.N Then
                        ValorEsquerda = 640
                        ValorTopo = 472
                    ElseIf e.KeyCode = Keys.M Then
                        ValorEsquerda = 768
                        ValorTopo = 472
                    ElseIf e.KeyCode = 188 Then '188 = tecla vírgula 
                        ValorEsquerda = 896
                        ValorTopo = 472


                    Else
                        Exit Sub
                    End If

                    CopiarAcorde()

                End If

            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub Acordes_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyUp
        Try

            lista.Remove(e.KeyCode)

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Acordes_MouseWheel(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseWheel
        Try

            If e.Delta > 0 Then
                Proximo()
            Else
                Anterior()
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub TonalidadeC()
        Try

            AjustePosiçãoTonalidadeDoAcorde = 0
            NotasAcordeIndiceLinha = 0
            TonalidadeAcorde = "C"
            GerarAcordes()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub TonalidadeD()
        Try

            AjustePosiçãoTonalidadeDoAcorde = 2
            NotasAcordeIndiceLinha = 1
            TonalidadeAcorde = "D"
            GerarAcordes()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub TonalidadeE()
        Try

            AjustePosiçãoTonalidadeDoAcorde = 4
            NotasAcordeIndiceLinha = 2
            TonalidadeAcorde = "E"
            GerarAcordes()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub TonalidadeF()
        Try

            AjustePosiçãoTonalidadeDoAcorde = 5
            NotasAcordeIndiceLinha = 3
            TonalidadeAcorde = "F"
            GerarAcordes()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub TonalidadeG()
        Try

            AjustePosiçãoTonalidadeDoAcorde = -5
            NotasAcordeIndiceLinha = 4
            TonalidadeAcorde = "G"
            GerarAcordes()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub TonalidadeA()
        Try

            AjustePosiçãoTonalidadeDoAcorde = -3
            NotasAcordeIndiceLinha = 5
            TonalidadeAcorde = "A"
            GerarAcordes()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub TonalidadeB()
        Try

            AjustePosiçãoTonalidadeDoAcorde = -1
            NotasAcordeIndiceLinha = 6
            TonalidadeAcorde = "B"
            GerarAcordes()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Label10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label10.Click
        TonalidadeC()
    End Sub

    Private Sub Label11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label11.Click
        TonalidadeD()
    End Sub

    Private Sub Label12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label12.Click
        TonalidadeE()
    End Sub

    Private Sub Label13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label13.Click
        TonalidadeF()
    End Sub

    Private Sub Label14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label14.Click
        TonalidadeG()
    End Sub

    Private Sub Label15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label15.Click
        TonalidadeA()
    End Sub

    Private Sub Label16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label16.Click
        TonalidadeB()
    End Sub

    Private Sub CorIntervalos_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CorIntervalos.Click
        ColorDialog1.Color = My.Settings.NovaCorIntervalos
        ColorDialog1.ShowDialog()
        My.Settings.NovaCorIntervalos = ColorDialog1.Color
        Desenhar()
    End Sub

    Private Sub Label17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label17.Click
        Try

            If My.Settings.NovoValorExibirNumerosDedilhado = True Then
                My.Settings.NovoValorExibirNumerosDedilhado = False
            Else
                My.Settings.NovoValorExibirNumerosDedilhado = True
            End If
            Desenhar()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Label18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label18.Click
        ExibirNomeDasNotasAcordes()
    End Sub

    Private Sub ExibirNomeDasNotasAcordes()
        Try

            If My.Settings.NovoValorExibirNomeDasNotasAcordes = True Then
                My.Settings.NovoValorExibirNomeDasNotasAcordes = False
            Else
                My.Settings.NovoValorExibirNomeDasNotasAcordes = True
                My.Settings.NovoValorExibirNomeIntervalos = False
            End If
            Desenhar()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Label19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label19.Click
        Try

            If My.Settings.NovoValorSalvarAcordes = True Then
                My.Settings.NovoValorSalvarAcordes = False
            Else
                My.Settings.NovoValorSalvarAcordes = True
            End If
            Desenhar()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Label20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label20.Click
        Try

            OpçõesTelaAcordes.ShowDialog()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Label21_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Label21.Click
        ExibirNotasBraçoViolão()
    End Sub

    Private Sub ExibirNotasBraçoViolão()
        Try

            If My.Settings.NovoValorExibirNotasBraçoViolão = True Then
                My.Settings.NovoValorExibirNotasBraçoViolão = False
            Else
                My.Settings.NovoValorExibirNotasBraçoViolão = True
                My.Settings.NovoValorExibirIntervalosBraçoViolão = False
            End If
            Desenhar()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Public Sub ExibeTodosOsAcordesParaResposta()
        Try

            modoJogo = False
            Desenhar()
            modoJogo = True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Public Sub GerarNovoAcordeJogo()

        Try

            modoJogo = True

            'Gera família de acordes aletório:
            NumeraçãoFamiliaAcorde = 0
            Do While NumeraçãoFamiliaAcorde = 0
                ' gera o array de bytes randômico de 4 bytes...
                Dim randomNumber(3) As Byte

                ' Create a new instance of the RNGCryptoServiceProvider.
                Dim Gen As New Security.Cryptography.RNGCryptoServiceProvider()

                ' Fill the array with a random value.
                Gen.GetBytes(randomNumber)

                ' calcula o número baseado no valor máximo
                NumeraçãoFamiliaAcorde = Math.Abs(BitConverter.ToInt32(randomNumber, 0)) Mod (54 + 1)

                If NumeraçãoFamiliaAcorde <> 0 AndAlso My.Settings.NovoValorModoJogoTelaAcordesCifras(NumeraçãoFamiliaAcorde - 1) = "False" Then NumeraçãoFamiliaAcorde = 0

            Loop


            'Gera tonalidade aleatória:
            NotasAcordeIndiceLinha = -1
            Do While NotasAcordeIndiceLinha = -1
                ' gera o array de bytes randômico de 4 bytes...
                Dim randomNumber(3) As Byte

                ' Create a new instance of the RNGCryptoServiceProvider.
                Dim Gen As New Security.Cryptography.RNGCryptoServiceProvider()

                ' Fill the array with a random value.
                Gen.GetBytes(randomNumber)

                ' calcula o número baseado no valor máximo
                NotasAcordeIndiceLinha = Math.Abs(BitConverter.ToInt32(randomNumber, 0)) Mod (6 + 1)

                If NotasAcordeIndiceLinha <> -1 AndAlso My.Settings.NovoValorModoJogoTelaAcordesTom(NotasAcordeIndiceLinha) = "False" Then NotasAcordeIndiceLinha = -1

            Loop

            If NotasAcordeIndiceLinha = 0 Then
                AjustePosiçãoTonalidadeDoAcorde = 0
                TonalidadeAcorde = "C"
            ElseIf NotasAcordeIndiceLinha = 1 Then
                AjustePosiçãoTonalidadeDoAcorde = 2
                TonalidadeAcorde = "D"
            ElseIf NotasAcordeIndiceLinha = 2 Then
                AjustePosiçãoTonalidadeDoAcorde = 4
                TonalidadeAcorde = "E"
            ElseIf NotasAcordeIndiceLinha = 3 Then
                AjustePosiçãoTonalidadeDoAcorde = 5
                TonalidadeAcorde = "F"
            ElseIf NotasAcordeIndiceLinha = 4 Then
                AjustePosiçãoTonalidadeDoAcorde = -5
                TonalidadeAcorde = "G"
            ElseIf NotasAcordeIndiceLinha = 5 Then
                AjustePosiçãoTonalidadeDoAcorde = -3
                TonalidadeAcorde = "A"
            ElseIf NotasAcordeIndiceLinha = 6 Then
                AjustePosiçãoTonalidadeDoAcorde = -1
                TonalidadeAcorde = "B"
            End If

           ImagemCopiada(0) = Nothing : ImagemCopiada(1) = Nothing : ImagemCopiada(2) = Nothing
            FamiliaAcorde = ""
            LocalizaNomeFamiliaAcordes()
            RedefiniçãoTonalidade()

            modoJogo = False

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try

    End Sub

    Private Sub Label24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label24.Click
        ExibirIntervalosBraçoViolão()
    End Sub

    Private Sub ExibirIntervalosBraçoViolão()
        Try

            If My.Settings.NovoValorExibirIntervalosBraçoViolão = True Then
                My.Settings.NovoValorExibirIntervalosBraçoViolão = False
            Else
                My.Settings.NovoValorExibirIntervalosBraçoViolão = True
                My.Settings.NovoValorExibirNotasBraçoViolão = False
            End If
            Desenhar()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub

    Private Sub Label23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label23.Click
        Try

            TelaAcordesModoJogo.Show()
            GerarNovoAcordeJogo()

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            'Se quiser apresentar alguma mensagem sobre o erro coloque aqui
        Finally
            'esta parte do bloco é executada independentemente se algum erro acontecer ou não

        End Try
    End Sub
End Class